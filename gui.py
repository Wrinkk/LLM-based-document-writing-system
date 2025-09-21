import sys
import os
import re
import datetime
import time
import string
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- GUI 라이브러리 ---
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QTextEdit, QFileDialog,
    QListWidget, QListWidgetItem, QGroupBox, QMessageBox, QProgressBar
)
from PyQt6.QtCore import QThread, QObject, pyqtSignal, Qt
from PyQt6.QtGui import QIcon, QFont

# --- 기존 코드의 라이브러리 ---
import pyhwpx
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold
import PyPDF2
from dotenv import load_dotenv

#==============================================================================
# 기존 스크립트의 핵심 로직 (Worker 스레드에서 호출될 함수들)
#==============================================================================

def combine_pdf_texts(pdf_files_paths, worker_signal):
    """여러 PDF 파일의 텍스트를 효율적으로 합치는 함수"""
    combined_text = ""
    for file_path in pdf_files_paths:
        if not os.path.exists(file_path):
            worker_signal.emit(f"   - 경고: '{os.path.basename(file_path)}' 파일을 찾을 수 없어 건너뜁니다.")
            continue
        worker_signal.emit(f"   - 텍스트 추출 중: {os.path.basename(file_path)}")
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                raw_text = "".join(page.extract_text() + "\n" for page in pdf_reader.pages if page.extract_text())
                
                if not raw_text.strip():
                    worker_signal.emit(f"   - 경고: '{os.path.basename(file_path)}'에서 텍스트를 추출할 수 없습니다.")
                    continue

                cleaned_text = re.sub(r'[^\w\s가-힣.,!?;:\'"""''·()\[\]{} -]', ' ', raw_text)
                cleaned_text = re.sub(r'\s+', ' ', cleaned_text).strip()
                
                combined_text += f"\n\n=== {os.path.basename(file_path)}의 내용 ===\n{cleaned_text}"
                worker_signal.emit(f"   - 완료: {len(cleaned_text):,} 문자")
        except Exception as e:
            worker_signal.emit(f"   - 오류: '{os.path.basename(file_path)}' 처리 중 오류 발생: {e}")
    
    worker_signal.emit(f"\n통합 및 정제된 텍스트 총 길이: {len(combined_text):,} 문자")
    return combined_text

def create_hwp_document_with_foreword(template_path, foreword_text, output_path, version_num, worker_signal):
    """발간사가 포함된 한글 문서를 생성하는 함수"""
    hwp = None
    try:
        hwp = pyhwpx.Hwp(visible=False)
        if os.path.exists(template_path):
            hwp.Open(template_path)
            worker_signal.emit(f"   - 템플릿 파일 열기: {os.path.basename(template_path)}")
        else:
            hwp.XHwpDocuments.Add()
            worker_signal.emit("   - 새 문서 생성")
        
        hwp.MoveDocBegin()
        
        sections = re.split(r'##\s*([A-Z]{2,})', foreword_text)
        content_map = {sections[i].strip(): sections[i+1].strip() for i in range(1, len(sections), 2)}

        worker_signal.emit(f"\n   - HWPX 파일에 {version_num}번째 답변 삽입 중...")
        for marker, full_content in content_map.items():
            title_only = full_content.split('\n')[0].strip()
            hwp.MoveDocBegin()
            if hwp.find(marker, direction='AllDoc'):
                hwp.insert_text(title_only)
                worker_signal.emit(f"     ✓ '{marker}' 위치에 '{title_only[:20]}...' 삽입 완료.")
        
        worker_signal.emit("   - 문서 내 잔여 파싱 마커 제거 중...")
        hwp.MoveDocBegin()
        possible_markers = [f"{char}{char}" for char in string.ascii_uppercase]
        for marker in possible_markers:
            while hwp.find(marker, direction='AllDoc'):
                hwp.Delete()
        
        hwp.SaveAs(output_path)
        return True, f"성공적으로 저장됨: {output_path}"
    except Exception as e:
        return False, f"오류 발생: {e}"
    finally:
        if hwp:
            try:
                hwp.Quit()
            except:
                pass

#==============================================================================
# PyQt6 Worker 클래스 (백그라운드 작업 처리)
#==============================================================================

class Worker(QObject):
    progress = pyqtSignal(str)
    finished = pyqtSignal(str)
    progress_bar = pyqtSignal(int)

    def __init__(self, settings):
        super().__init__()
        self.settings = settings
        self.uploaded_files = []

    def run(self):
        try:
            self._execute_main_logic()
        except Exception as e:
            self.progress.emit(f"\n❌ 프로그램 실행 중 치명적 오류 발생: {e}")
            import traceback
            self.progress.emit(traceback.format_exc())
        finally:
            self._cleanup()

    def _cleanup(self):
        if self.uploaded_files:
            self.progress.emit("\n프로그램 정리 작업 시작...")
            self.progress.emit("업로드된 임시 파일들을 삭제합니다...")
            for file in self.uploaded_files:
                try:
                    genai.delete_file(file.name)
                    self.progress.emit(f"   ✓ 삭제: {file.display_name}")
                except Exception as e:
                    self.progress.emit(f"   ✗ 삭제 실패: {file.display_name} - {e}")
        self.finished.emit(self.settings['output_dir'])
    
    def _upload_file_concurrently(self, file_path):
        """단일 파일을 Gemini API에 업로드하고 결과를 반환하는 함수 (스레드에서 실행)"""
        if not os.path.exists(file_path):
            self.progress.emit(f"   - 경고: '{os.path.basename(file_path)}' 파일을 찾을 수 없어 건너뜁니다.")
            return None
        file_size = os.path.getsize(file_path)
        if file_size > 200 * 1024 * 1024:
            self.progress.emit(f"   - 경고: '{os.path.basename(file_path)}' 파일 크기가 200MB를 초과하여 건너뜁니다.")
            return None
        self.progress.emit(f"   - 업로드 시작: {os.path.basename(file_path)}")
        try:
            file_response = genai.upload_file(path=file_path)
            self.progress.emit(f"   - 업로드 완료: {file_response.display_name}")
            return file_response
        except Exception as e:
            self.progress.emit(f"   - 업로드 실패: {os.path.basename(file_path)} - {e}")
            return None

    def _wait_for_file_processing(self, uploaded_files_responses):
        """업로드된 파일들의 처리 완료를 대기하는 함수"""
        for file in uploaded_files_responses:
            start_time = time.time()
            while file.state.name == "PROCESSING":
                elapsed = time.time() - start_time
                if elapsed > 300:
                    self.progress.emit(f"   - {file.display_name} 처리 시간 초과")
                    break
                self.progress.emit(f"   - {file.display_name} 처리 중... 5초 대기 (경과: {elapsed:.0f}초)")
                time.sleep(5)
                file = genai.get_file(file.name)
            
            if file.state.name == "FAILED":
                self.progress.emit(f"   - 경고: {file.display_name} 처리 실패")
            elif file.state.name == "ACTIVE":
                self.progress.emit(f"   - {file.display_name} 처리 완료")

    def _execute_main_logic(self):
        """메인 실행 함수"""
        settings = self.settings
        genai.configure(api_key=settings['api_key'])
        
        self.progress.emit("=" * 60)
        self.progress.emit("울진군 발간사 생성 프로그램 시작")
        self.progress.emit(f"결과 저장 경로: {settings['output_dir']}")
        self.progress.emit("=" * 60)
        
        # 1. 텍스트 추출 그룹 처리
        combined_pdf_text = ""
        if settings['text_extract_paths']:
            self.progress.emit("\nPDF에서 텍스트를 추출합니다...")
            combined_pdf_text = combine_pdf_texts(settings['text_extract_paths'], self.progress)

        # 2. 파일 업로드 그룹 처리
        if settings['file_upload_paths']:
            self.progress.emit("\nPDF 파일을 AI에 직접 업로드합니다 (병렬 처리)...")
            with ThreadPoolExecutor(max_workers=10) as executor:
                future_to_file = {executor.submit(self._upload_file_concurrently, fp): fp for fp in settings['file_upload_paths']}
                for future in as_completed(future_to_file):
                    file_response = future.result()
                    if file_response:
                        self.uploaded_files.append(file_response)
            
            if self.uploaded_files:
                self._wait_for_file_processing(self.uploaded_files)
                self.progress.emit(f"\n총 {len(self.uploaded_files)}개 파일 업로드 완료")

        # 3. 모델 초기화
        self.progress.emit("\nAI 모델을 초기화합니다...")
        try:
            model = genai.GenerativeModel(
                'gemini-1.5-flash', 
                safety_settings={
                    HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
                    HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
                    HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
                    HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
                }
            )
            self.progress.emit("   - 모델: gemini-1.5-flash")
        except Exception as e:
            self.progress.emit(f"   - 모델 초기화 실패: {e}")
            return

        year = datetime.datetime.now().year
        quarter = ((datetime.datetime.now().month - 1) // 3 + 1)

        # 4. 5개의 다른 발간사 생성
        foreword_variations = settings['foreword_variations']
        fallback_forewords = settings['fallback_forewords']

        summarized_text = ""
        if combined_pdf_text:
            summarized_text = combined_pdf_text[:1000000] 
        
        total_steps = len(foreword_variations)
        for i, variation in enumerate(foreword_variations, 1):
            self.progress.emit(f"\n[{i}/{total_steps}] '{variation['focus']}' 발간사 생성 중...")
            self.progress_bar.emit(int((i / total_steps) * 100))
            
            max_retries = 3
            success = False
            for retry_count in range(max_retries):
                try:
                    prompt_parts = []
                    if i == 1 and retry_count == 0:
                        if self.uploaded_files: prompt_parts.extend(self.uploaded_files)
                        if summarized_text: prompt_parts.append(f"[참고 자료 요약]\n{summarized_text}")
                    
                    prompt_text = f"""
                        {year}년도 {quarter}분기 울진군 군정집 발간사를 작성해주세요.
                        **작성 지침:**
                        1. 버전 특징: 초점({variation['focus']}), 어조({variation['tone']}), 강조점({variation['emphasis']})
                        2. 필수 포함 내용: 경제, 화합, 희망 3대 키워드, 구체적 성과와 비전, 군민 감사/격려, {year}년 {quarter}분기 시의성 반영
                        3. 형식 요구사항: 10줄 내외, 공식적/품격있는 문체, 구체적/실질적 내용, 인사말로 시작하여 감사/격려로 마무리.
                           각 문장마다 파싱문자 ##AA,##BB,##CC... 를 부여. 예시: ##AA 내용 ##BB 내용
                        4. 참고: 이전 분기 발간사 내용은 이미 숙지하고 있으니, 이를 바탕으로 더 발전된 내용을 작성할 것.
                        5. 제약: 발간사 본문만 작성하고 다른 설명은 포함하지 말 것. 키워드에 '', ** 와 같은 특수 표시를 하지 말 것. 5개 버전의 내용이 서로 다르게 작성될 것.
                    """
                    prompt_parts.append(prompt_text)
                    
                    response = model.generate_content(prompt_parts)
                    
                    if response and hasattr(response, 'text') and len(response.text.strip()) > 200:
                        foreword_text = response.text.strip()
                        self.progress.emit(f"   ✓ AI 발간사 생성 완료 ({len(foreword_text)}자)")
                        
                        template_path = settings['template_paths'][i-1] if i <= len(settings['template_paths']) else ""
                        output_filename = f"발간사_{i:02d}_{variation['focus'].replace(' ', '_')}.hwpx"
                        output_path = os.path.join(settings['output_dir'], output_filename)
                        
                        success, message = create_hwp_document_with_foreword(template_path, foreword_text, output_path, i, self.progress)
                        if success:
                            self.progress.emit(f"   ✓ 한글 파일 생성: {output_filename}")
                        else:
                            self.progress.emit(f"   ✗ 한글 파일 생성 실패: {message}")
                        success = True
                        break
                    else:
                        self.progress.emit(f"   ✗ 생성된 내용이 짧거나 유효하지 않음. 재시도 {retry_count+1}/{max_retries}")
                except Exception as e:
                    self.progress.emit(f"   ✗ API 오류 발생: {str(e)[:100]}. 재시도 {retry_count+1}/{max_retries}")
                    time.sleep(5)
            
            if not success:
                self.progress.emit(f"   ✗ 최종 실패. 하드코딩된 기본 발간사를 사용합니다.")
                template_foreword = fallback_forewords[i-1]
                
                template_path = settings['template_paths'][i-1] if i <= len(settings['template_paths']) else ""
                output_filename = f"발간사_{i:02d}_{variation['focus'].replace(' ', '_')}_기본.hwpx"
                output_path = os.path.join(settings['output_dir'], output_filename)

                success, message = create_hwp_document_with_foreword(template_path, template_foreword, output_path, i, self.progress)
                if success:
                    self.progress.emit(f"   ✓ 기본 템플릿 파일 생성: {output_filename}")

            if i < total_steps:
                time.sleep(2) # 다음 요청 전 잠시 대기
        
        self.progress_bar.emit(100)


#==============================================================================
# PyQt6 MainApp 클래스 (GUI 창 및 위젯 관리)
#==============================================================================

class MainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("울진군 발간사 자동 생성 프로그램")
        self.setGeometry(100, 100, 800, 700)
        
        # 아이콘 설정 (실행 파일과 같은 위치에 icon.png가 있어야 함)
        if os.path.exists("icon.png"):
            self.setWindowIcon(QIcon("icon.png"))

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        
        self._create_settings_ui()
        self._create_file_list_ui()
        self._create_log_ui()
        self._create_action_ui()

        self.thread = None
        self.worker = None
        
        load_dotenv()
        self.api_key_input.setText(os.getenv("GEMINI_API_KEY", ""))

    def _create_settings_ui(self):
        settings_group = QGroupBox("1. 기본 설정")
        settings_layout = QVBoxLayout()

        # API Key
        api_layout = QHBoxLayout()
        api_layout.addWidget(QLabel("Gemini API Key:"))
        self.api_key_input = QLineEdit()
        self.api_key_input.setPlaceholderText(".env 파일에서 로드됨. 직접 입력하여 변경 가능")
        api_layout.addWidget(self.api_key_input)
        settings_layout.addLayout(api_layout)

        # Template Directory
        template_layout = QHBoxLayout()
        template_layout.addWidget(QLabel("템플릿 폴더:"))
        self.template_dir_input = QLineEdit(r'C:\Users\wj830\Desktop\dd')
        template_layout.addWidget(self.template_dir_input)
        self.template_btn = QPushButton("폴더 선택")
        self.template_btn.clicked.connect(self.select_template_dir)
        template_layout.addWidget(self.template_btn)
        settings_layout.addLayout(template_layout)
        
        settings_group.setLayout(settings_layout)
        self.layout.addWidget(settings_group)

    def _create_file_list_ui(self):
        files_group = QGroupBox("2. 학습 데이터(PDF) 설정")
        files_layout = QHBoxLayout()

        # Text Extract List
        extract_layout = QVBoxLayout()
        extract_layout.addWidget(QLabel("텍스트 추출 후 학습할 PDF 목록"))
        self.text_extract_list = QListWidget()
        extract_layout.addWidget(self.text_extract_list)
        extract_btn_layout = QHBoxLayout()
        add_extract_btn = QPushButton("추가")
        add_extract_btn.clicked.connect(self.add_text_extract_files)
        remove_extract_btn = QPushButton("제거")
        remove_extract_btn.clicked.connect(lambda: self.remove_selected_item(self.text_extract_list))
        extract_btn_layout.addWidget(add_extract_btn)
        extract_btn_layout.addWidget(remove_extract_btn)
        extract_layout.addLayout(extract_btn_layout)
        files_layout.addLayout(extract_layout)

        # File Upload List
        upload_layout = QVBoxLayout()
        upload_layout.addWidget(QLabel("파일 직접 업로드하여 학습할 PDF 목록"))
        self.file_upload_list = QListWidget()
        upload_layout.addWidget(self.file_upload_list)
        upload_btn_layout = QHBoxLayout()
        add_upload_btn = QPushButton("추가")
        add_upload_btn.clicked.connect(self.add_upload_files)
        remove_upload_btn = QPushButton("제거")
        remove_upload_btn.clicked.connect(lambda: self.remove_selected_item(self.file_upload_list))
        upload_btn_layout.addWidget(add_upload_btn)
        upload_btn_layout.addWidget(remove_upload_btn)
        upload_layout.addLayout(upload_btn_layout)
        files_layout.addLayout(upload_layout)
        
        files_group.setLayout(files_layout)
        self.layout.addWidget(files_group)

    def _create_log_ui(self):
        log_group = QGroupBox("3. 작업 로그")
        log_layout = QVBoxLayout()
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setFont(QFont("Courier New", 9))
        log_layout.addWidget(self.log_output)
        self.progress_bar = QProgressBar()
        log_layout.addWidget(self.progress_bar)
        log_group.setLayout(log_layout)
        self.layout.addWidget(log_group)

    def _create_action_ui(self):
        self.start_btn = QPushButton("발간사 생성 시작")
        self.start_btn.setFixedHeight(40)
        self.start_btn.setStyleSheet("background-color: #4CAF50; color: white; font-size: 16px; border-radius: 5px;")
        self.start_btn.clicked.connect(self.start_process)
        self.layout.addWidget(self.start_btn)

    def select_template_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "템플릿 폴더 선택")
        if directory:
            self.template_dir_input.setText(directory)

    def add_files_to_list(self, list_widget):
        files, _ = QFileDialog.getOpenFileNames(self, "PDF 파일 선택", "", "PDF Files (*.pdf)")
        for file in files:
            list_widget.addItem(QListWidgetItem(file))

    def add_text_extract_files(self):
        self.add_files_to_list(self.text_extract_list)

    def add_upload_files(self):
        self.add_files_to_list(self.file_upload_list)
        
    def remove_selected_item(self, list_widget):
        for item in list_widget.selectedItems():
            list_widget.takeItem(list_widget.row(item))

    def update_log(self, message):
        self.log_output.append(message)
        self.log_output.verticalScrollBar().setValue(self.log_output.verticalScrollBar().maximum())

    def set_progress_bar(self, value):
        self.progress_bar.setValue(value)

    def on_process_finished(self, output_dir):
        self.update_log("\n✅ 모든 작업 완료!")
        self.start_btn.setEnabled(True)
        self.start_btn.setText("발간사 생성 시작")
        self.start_btn.setStyleSheet("background-color: #4CAF50; color: white; font-size: 16px; border-radius: 5px;")
        
        reply = QMessageBox.information(self, "완료", f"발간사 생성이 완료되었습니다.\n결과 폴더: {output_dir}\n폴더를 여시겠습니까?",
                                      QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            os.startfile(output_dir)

    def start_process(self):
        api_key = self.api_key_input.text().strip()
        if not api_key:
            QMessageBox.warning(self, "오류", "Gemini API 키를 입력해주세요.")
            return

        template_dir = self.template_dir_input.text().strip()
        if not os.path.isdir(template_dir):
            QMessageBox.warning(self, "오류", "유효한 템플릿 폴더를 선택해주세요.")
            return

        settings = {
            'api_key': api_key,
            'template_base_dir': template_dir,
            'template_paths': [os.path.join(template_dir, f"template{i}.hwpx") for i in range(1, 6)],
            'text_extract_paths': [self.text_extract_list.item(i).text() for i in range(self.text_extract_list.count())],
            'file_upload_paths': [self.file_upload_list.item(i).text() for i in range(self.file_upload_list.count())],
            'output_dir': os.path.join(template_dir, f"발간사_결과_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"),
            'foreword_variations': [
                {"focus": "경제 발전 중심", "tone": "역동적이고 진취적인 어조", "emphasis": "울진군의 미래 성장 동력과 경제 발전 전략"},
                {"focus": "군민 화합 중심", "tone": "따뜻하고 포용적인 어조", "emphasis": "군민과의 소통과 협력, 상생의 가치"},
                {"focus": "희망과 비전 중심", "tone": "미래지향적이고 희망찬 어조", "emphasis": "울진군의 밝은 미래와 지속가능한 발전"},
                {"focus": "균형적 접근", "tone": "안정적이고 신뢰감 있는 어조", "emphasis": "경제, 화합, 희망이 균형잡힌 종합적 관점"},
                {"focus": "혁신과 변화 중심", "tone": "혁신적이고 도전적인 어조", "emphasis": "새로운 변화와 혁신을 통한 울진군의 도약"}
            ],
            'fallback_forewords': [
                # 1번째 발간사 (경제 발전 중심) 실패 시 사용될 내용
                """##AA 존경하는 울진 군민 여러분, 2025년 한 해를 마무리하며 새로운 희망을 품는 4분기를 맞이했습니다.##BB 올해 우리는 미래 경제 성장을 위한 기반을 다지며 괄목할 만한 성과를 이루어냈습니다.##CC 원자력수소 국가산업단지 조성과 같은 미래 성장 동력 확보에 매진하며, 지역 경제에 활력을 불어넣었습니다.##DD 또한, 관광 인프라를 전면 개편하고 농업 대전환을 통해 지역 특산물의 가치를 높였습니다.##EE 전통시장 활성화와 소상공인 지원으로 지역 경제의 든든한 화합을 이끌어냈습니다.##FF 이 모든 성과는 군민 여러분의 적극적인 참여와 노력이 있었기에 가능했습니다.##GG 역동적인 변화와 진취적인 도전으로, 울진은 더 큰 희망을 향해 나아가고 있습니다.##HH 앞으로도 지속적인 경제 발전 전략을 통해 군민 모두가 풍요로운 미래를 누리도록 최선을 다하겠습니다.##II 늘 함께해주시는 군민 여러분께 진심으로 감사드립니다.##JJ 다가오는 새해에도 울진의 밝은 미래를 함께 만들어 주시기를 바랍니다.""",
                # 2번째 발간사 (군민 화합 중심) 실패 시 사용될 내용
                """##AA 존경하는 울진군민 여러분, 2025년 한 해를 마무리하며 희망찬 새해를 준비하는 4분기 군정집을 발간합니다.##BB 올 한 해 울진은 군민 여러분의 지혜와 역량이 한데 모여 더욱 굳건한 공동체로 성장했습니다.##CC 어려운 여건 속에서도 흔들림 없는 경제 기반을 다지고자 끊임없이 노력하여 미래 성장 동력을 확보했습니다.##DD 이 모든 과정에서 서로를 보듬고 아끼는 화합의 정신은 우리 군이 직면한 여러 과제를 슬기롭게 극복하는 원동력이 되었습니다.##EE 군민과 함께 소통하며 만들어 온 정책들은 울진의 삶의 질을 높이고 내일의 희망을 키워냈습니다.##FF 저희는 앞으로도 군민의 삶을 최우선에 두고, 상생의 가치를 실현하는 따뜻하고 포용적인 군정을 펼쳐나가겠습니다.##GG 2026년에도 군민 여러분과 손잡고 더 큰 울진의 발전과 도약을 위한 비전을 향해 힘껏 나아가겠습니다.##HH 모든 세대가 함께 웃고, 미래를 꿈꿀 수 있는 행복한 울진을 만들기 위한 여정에 변함없는 동참을 부탁드립니다.##II 군민 여러분의 깊은 관심과 따뜻한 격려에 진심으로 감사드리며, 가정에 늘 건강과 행복이 가득하시기를 기원합니다.##JJ 감사합니다.""",
                # 3번째 발간사 (희망과 비전 중심) 실패 시 사용될 내용
                """##AA 존경하는 울진군민 여러분, 2025년 한 해를 마무리하는 제4분기 군정집을 발간하게 되어 매우 뜻깊게 생각합니다.##BB 올 한 해 군민 여러분과 함께 일궈온 값진 노력들이 풍성한 결실을 맺으며 희망찬 울진의 미래를 밝히고 있습니다.##CC 특히, 지속가능한 경제 발전을 위한 새로운 성장 동력 확보는 물론, 지역 경제에 활력을 불어넣기 위한 노력을 멈추지 않았습니다.##DD 이 모든 과정 속에서 군민들의 하나 된 화합과 참여는 울진 발전의 가장 든든한 초석이 되었습니다.##EE 우리는 더 나은 내일을 향한 분명한 비전을 가지고, 모두가 살기 좋은 울진을 만들기 위한 담대한 도전을 이어가고 있습니다.##FF 청정 자연을 기반으로 한 친환경 성장과 혁신적인 정책을 통해 울진의 지속가능한 발전 모델을 정립해 나갈 것입니다.##GG 군민 한 분 한 분의 삶에 따뜻한 희망이 스며들고, 모두가 행복을 누리는 풍요로운 울진을 향해 나아가겠습니다.##HH 지난 한 해 동안 변함없는 사랑과 성원을 보내주신 군민 여러분께 진심으로 감사드립니다.##II 다가오는 새해에도 군민 여러분과 함께 더 큰 도약을 이루어낼 수 있도록 최선을 다하겠습니다.##JJ 군민 여러분의 가정에 늘 건강과 행복이 가득하시기를 기원합니다.""",
                # 4번째 발간사 (균형적 접근) 실패 시 사용될 내용
                """##AA 존경하는 울진군민 여러분, 2025년 한 해를 마무리하는 4분기 군정집을 발간하며 인사드립니다.##BB 올 한 해도 변함없이 군정에 깊은 관심과 성원을 보내주신 군민 여러분께 진심으로 감사드립니다.##CC 견고한 지역 경제 기반을 다지기 위한 노력은 가시적인 성과로 이어져, 지속 가능한 성장의 발판을 마련하였습니다.##DD 군민의 화합된 힘은 크고 작은 어려움을 슬기롭게 극복하고, 서로를 배려하는 공동체 정신을 더욱 굳건히 하는 원동력이 되었습니다.##EE 우리는 미래 세대가 더 나은 삶을 꿈꿀 수 있는 희망찬 울진을 만들기 위해 끊임없이 고민하고 실천해왔습니다.##FF 해양 관광 활성화부터 스마트 농업 혁신, 그리고 품격 있는 문화생활에 이르기까지, 균형 잡힌 발전을 위해 다각적인 노력을 기울였습니다.##GG 다가오는 새해에는 더 큰 도약과 변화를 향해 나아가며, 군민 한 분 한 분의 삶이 더욱 풍요로워지도록 최선을 다하겠습니다.##HH 울진의 밝은 미래를 향한 여정에 늘 함께해주시는 군민 여러분의 지혜와 참여에 깊이 감사드립니다.##II 앞으로도 변함없는 사랑과 성원을 부탁드리며, 가정에 늘 건강과 행복이 가득하시기를 기원합니다.""",
                # 5번째 발간사 (혁신과 변화 중심) 실패 시 사용될 내용
                """##AA 존경하는 울진군민 여러분, 2025년 한 해를 마무리하며 새 시대를 향한 도약의 의지를 담은 군정집을 선보입니다.##BB 올 한 해 울진군은 혁신과 변화를 기치로 내걸고, 미래를 향한 과감한 도전을 멈추지 않았습니다.##CC 이러한 노력들이 결실을 맺어 지역 경제에 활력을 불어넣고, 새로운 성장 동력을 확보하는 데 주력했습니다.##DD 특히, 군민 모두가 한마음으로 뜻을 모아 이룬 화합의 가치는 어떠한 난관도 극복할 수 있는 굳건한 울진의 힘이 되었습니다.##EE 다가올 새해에는 스마트 도시 기반 구축과 미래 신산업 육성을 통해 지속 가능한 발전의 희망찬 비전을 현실로 만들어갈 것입니다.##FF 우리는 변화를 두려워하지 않고, 더 나은 울진의 내일을 위한 혁신적 시도를 계속해 나갈 것입니다.##GG 군민 여러분의 적극적인 참여와 성원이 있었기에 오늘의 성과가 가능했으며, 이는 미래를 향한 가장 강력한 추진력입니다.##HH 앞으로도 군민의 삶의 질 향상과 울진의 위상 강화를 위해 모든 역량을 집중할 것을 약속드립니다.##II 뜨거운 열정으로 함께해 주신 군민 여러분께 진심으로 감사드리며, 새해에도 변함없는 지지와 격려를 부탁드립니다.##JJ 희망찬 미래를 향한 울진의 여정에 동참해 주시길 바랍니다."""
            ]
        }
        
        os.makedirs(settings['output_dir'], exist_ok=True)
        self.log_output.clear()
        self.progress_bar.setValue(0)
        
        self.start_btn.setEnabled(False)
        self.start_btn.setText("생성 중...")
        self.start_btn.setStyleSheet("background-color: #FFA500; color: white; font-size: 16px; border-radius: 5px;")
        
        self.thread = QThread()
        self.worker = Worker(settings)
        self.worker.moveToThread(self.thread)
        
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        
        self.worker.progress.connect(self.update_log)
        self.worker.progress_bar.connect(self.set_progress_bar)
        self.worker.finished.connect(self.on_process_finished)
        
        self.thread.start()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainApp()
    main_window.show()
    sys.exit(app.exec())

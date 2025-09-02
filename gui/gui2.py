import sys
import os
import re
import datetime
import json
import traceback
import time

from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, 
                             QProgressBar, QFileDialog, QGroupBox, QComboBox)
from PyQt5.QtCore import QThread, pyqtSignal

# 필수 라이브러리가 없을 경우를 대비한 안내
try:
    import anthropic
    from pypdf import PdfReader
    import pyhwpx
except ImportError:
    print("="*60)
    print("필수 라이브러리가 설치되지 않았습니다.")
    print("터미널(명령 프롬프트)에 아래 명령어를 입력하여 설치해주세요.")
    print("pip install PyQt5 anthropic pypdf pyhwpx")
    print("="*60)
    sys.exit()

# -----------------------------------------------------------------------------
# 백그라운드에서 모든 자동화 작업을 처리하는 스레드 클래스
# -----------------------------------------------------------------------------
class DocumentProcessor(QThread):
    progress_update = pyqtSignal(str, int)
    finished = pyqtSignal(str)

    def __init__(self, keyword, api_key, hwp_path, pdf_paths, model_name):
        super().__init__()
        self.keyword = keyword
        self.api_key = api_key
        self.hwp_path = hwp_path
        self.pdf_paths = pdf_paths
        self.model_name = model_name

    def run(self):
        hwp = None
        try:
            hwp = pyhwpx.Hwp()
            self.progress_update.emit("한/글 프로그램을 시작합니다...", 5)
            
            client = anthropic.Anthropic(api_key=self.api_key)

            if not os.path.exists(self.hwp_path):
                hwp.SaveAs(self.hwp_path)
            hwp.Open(self.hwp_path)

            self.progress_update.emit("PDF 파일에서 텍스트를 추출합니다...", 10)
            pdf_context = ""
            num_pdfs = len(self.pdf_paths)
            for i, file_path in enumerate(self.pdf_paths):
                progress = 10 + int((i / num_pdfs) * 15)
                self.progress_update.emit(f"  - 처리 중: {os.path.basename(file_path)}", progress)
                try:
                    reader = PdfReader(file_path)
                    for page in reader.pages:
                        if page.extract_text():
                            pdf_context += page.extract_text() + "\n\n"
                except Exception as e:
                    self.progress_update.emit(f"  - '{os.path.basename(file_path)}' 처리 오류: {e}", progress)
            self.progress_update.emit("PDF 텍스트 추출 완료.", 25)

            conversation_history = []

            # --- 헬퍼 함수 정의 ---
            def process_text_response(ai_text, progress_start):
                self.progress_update.emit("AI 답변(대제목) 파싱 및 삽입 중...", progress_start)
                sections = re.split(r'##\s*([A-Z]{3,})', ai_text)
                if len(sections) > 1:
                    for i in range(1, len(sections), 2):
                        marker = sections[i].strip()
                        full_content = sections[i+1].strip()
                        title_only = full_content.split('\n')[0].strip()
                        hwp.MoveDocBegin()
                        while hwp.find(marker):
                            hwp.insert_text(title_only)
                            time.sleep(0.1)
            
            def process_json_response(ai_text, progress_start):
                self.progress_update.emit("AI 답변(JSON) 파싱 및 삽입 중...", progress_start)
                json_match = re.search(r'\{.*\}', ai_text, re.DOTALL)
                if json_match:
                    try:
                        json_data = json.loads(json_match.group(0))
                        for marker, content in json_data.items():
                            if isinstance(content, str):
                                hwp.MoveDocBegin()
                                while hwp.find(marker):
                                    hwp.insert_text(content)
                                    time.sleep(0.1)
                    except json.JSONDecodeError:
                        self.progress_update.emit(f"  - JSON 파싱 오류. ({ai_text[:30]}...)", progress_start)

            def ask_claude(prompt_text, progress_start):
                self.progress_update.emit(f"{progress_start}%. AI에게 요청 전송...", progress_start)
                
                if not conversation_history:
                    user_content = f"### 참고 문서:\n{pdf_context}\n\n### 첫 번째 요청:\n{prompt_text}"
                else:
                    user_content = prompt_text
                
                conversation_history.append({"role": "user", "content": user_content})

                response = client.messages.create(
                    model=self.model_name,
                    max_tokens=4096,
                    messages=conversation_history
                )
                ai_text = response.content[0].text
                conversation_history.append({"role": "assistant", "content": ai_text})
                return ai_text

            # --- 다단계 AI 요청 수행 ---
            first_prompt = f"참고 문서를 바탕으로 중복되는 업무 계획의 대제목을 생성해줘. 반드시 '## 식별자 제목' 형식으로만 답변하고, 식별자는 AAA, BBB 순서로 사용해줘."
            ai_response = ask_claude(first_prompt, 30)
            process_text_response(ai_response, 40)
            
            second_prompt = f"이전 답변과 참고 문서를 바탕으로, '{self.keyword}' 키워드에 맞는 세부 내용을 순수한 JSON 형식으로만 작성해줘. 키는 이전 답변의 대제목에 맞춰 AA1, AA2, BB1... 형식을 사용해줘."
            ai_response = ask_claude(second_prompt, 50)
            process_json_response(ai_response, 60)

            third_prompt = "지금까지의 대화와 참고 문서를 종합해서, 2025년의 주요 성과를 순수한 JSON 형식으로만 정리해줘. 키는 AC1, AC2... 형식을 사용해줘."
            ai_response = ask_claude(third_prompt, 70)
            process_json_response(ai_response, 80)
            
            fourth_prompt = "지금까지의 대화와 참고 문서를 종합해서, 2026년도 특수시책과 핵심과제를 순수한 JSON 형식으로만 제시해줘. 키는 J1(핵심과제), H1(특수시책) 형식을 사용해줘."
            ai_response = ask_claude(fourth_prompt, 85)
            process_json_response(ai_response, 90)

            # --- 연도 자동 변경 ---
            self.progress_update.emit("5. 연도를 자동으로 변경합니다...", 95)
            this_year = datetime.date.today().year
            hwp.MoveDocBegin();
            while hwp.find('YEAR'): hwp.insert_text(str(this_year + 1))
            hwp.MoveDocBegin();
            while hwp.find('YDDY'): hwp.insert_text(str(this_year))

            # --- 결과 파일 저장 (파일명 중복 방지) ---
            base, ext = os.path.splitext(self.hwp_path)
            output_path = f"{base}_결과{ext}" 
            counter = 1
            while os.path.exists(output_path):
                output_path = f"{base}_결과 ({counter}){ext}"
                counter += 1
            
            hwp.save_as(output_path)
            self.progress_update.emit(f"결과 파일 저장 완료: {os.path.basename(output_path)}", 100)
            
            self.finished.emit(f"모든 작업이 완료되었습니다!\n결과 파일: {output_path}")

        except Exception:
            error_message = f"오류 발생:\n{traceback.format_exc()}"
            self.finished.emit(error_message)
        finally:
            if hwp and hwp.api:
                hwp.Quit()

# -----------------------------------------------------------------------------
# PyQT5 메인 GUI 애플리케이션 클래스
# -----------------------------------------------------------------------------
class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_paths = []
        self.initUI()
        self.processor_thread = None

    def initUI(self):
        main_vbox = QVBoxLayout()
        self.setLayout(main_vbox)

        file_groupbox = QGroupBox("입력 파일 설정")
        file_vbox = QVBoxLayout()
        file_groupbox.setLayout(file_vbox)

        self.hwp_path_label = QLineEdit(self)
        self.hwp_path_label.setPlaceholderText("한/글 템플릿 파일을 선택하세요.")
        self.hwp_path_label.setReadOnly(True)
        hwp_button = QPushButton("템플릿 파일 선택")
        hwp_button.clicked.connect(self.select_hwp_file)
        hwp_hbox = QHBoxLayout()
        hwp_hbox.addWidget(self.hwp_path_label)
        hwp_hbox.addWidget(hwp_button)
        file_vbox.addLayout(hwp_hbox)

        self.pdf_path_label = QLineEdit(self)
        self.pdf_path_label.setPlaceholderText("AI가 참고할 PDF 파일들을 선택하세요 (다중 선택 가능).")
        self.pdf_path_label.setReadOnly(True)
        pdf_button = QPushButton("PDF 파일 선택")
        pdf_button.clicked.connect(self.select_pdf_files)
        pdf_hbox = QHBoxLayout()
        pdf_hbox.addWidget(self.pdf_path_label)
        pdf_hbox.addWidget(pdf_button)
        file_vbox.addLayout(pdf_hbox)
        
        run_groupbox = QGroupBox("실행")
        run_vbox = QVBoxLayout()
        run_groupbox.setLayout(run_vbox)

        self.model_label = QLabel('AI 모델 선택:')
        self.model_combo = QComboBox(self)
        self.model_combo.addItems([
            "claude-3-5-sonnet-20240620",
            "claude-3-opus-20240229",
            "claude-3-sonnet-20240229",
            "claude-3-haiku-20240307"
        ])
        model_hbox = QHBoxLayout()
        model_hbox.addWidget(self.model_label)
        model_hbox.addWidget(self.model_combo)
        run_vbox.addLayout(model_hbox)

        self.keyword_label = QLabel('핵심 키워드:')
        self.keyword_input = QLineEdit(self)
        self.keyword_input.setPlaceholderText('예: 도전, 혁신')
        keyword_hbox = QHBoxLayout()
        keyword_hbox.addWidget(self.keyword_label)
        keyword_hbox.addWidget(self.keyword_input)
        run_vbox.addLayout(keyword_hbox)
        
        self.run_button = QPushButton('문서 생성 시작', self)
        self.run_button.clicked.connect(self.start_processing)
        run_vbox.addWidget(self.run_button)

        progress_groupbox = QGroupBox("진행 상황")
        progress_vbox = QVBoxLayout()
        progress_groupbox.setLayout(progress_vbox)
        
        self.progress_bar = QProgressBar(self)
        self.status_log = QTextEdit(self)
        self.status_log.setReadOnly(True)
        self.status_log.setText("--- 대기 중 ---")
        progress_vbox.addWidget(self.progress_bar)
        progress_vbox.addWidget(self.status_log)

        main_vbox.addWidget(file_groupbox)
        main_vbox.addWidget(run_groupbox)
        main_vbox.addWidget(progress_groupbox)
        
        self.setWindowTitle('AI 문서 자동 생성 프로그램 (Claude.ver v2.2)')
        self.setGeometry(300, 300, 600, 550)
        self.show()

    def select_hwp_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, 'HWPX 템플릿 파일 선택', '', 'Hangul Files (*.hwp *.hwpx)')
        if fname: self.hwp_path_label.setText(fname)

    def select_pdf_files(self):
        fnames, _ = QFileDialog.getOpenFileNames(self, 'PDF 파일 선택', '', 'PDF Files (*.pdf)')
        if fnames:
            self.pdf_paths = fnames
            self.pdf_path_label.setText(f"{len(fnames)}개 파일 선택됨: {os.path.basename(fnames[0])} 등")

    def start_processing(self):
        keyword = self.keyword_input.text()
        hwp_path = self.hwp_path_label.text()
        model_name = self.model_combo.currentText()
        
        API_KEY = ""

        if not all([keyword, hwp_path, self.pdf_paths]):
            self.status_log.setText("오류: 모든 파일(HWPX, PDF)을 선택하고, 키워드를 입력해주세요.")
            return
            
        if not API_KEY or API_KEY == "YOUR_CLAUDE_API_KEY_HERE":
            self.status_log.setText("오류: 코드의 API_KEY 변수에 본인의 Claude API 키를 설정해주세요.")
            return

        self.run_button.setEnabled(False)
        self.status_log.clear()
        self.progress_bar.setValue(0)

        self.processor_thread = DocumentProcessor(keyword, API_KEY, hwp_path, self.pdf_paths, model_name)
        self.processor_thread.progress_update.connect(self.update_status)
        self.processor_thread.finished.connect(self.on_finished)
        self.processor_thread.start()

    def update_status(self, message, value):
        self.status_log.append(message)
        self.progress_bar.setValue(value)

    def on_finished(self, message):
        self.status_log.append(f"\n--- 작업 완료 ---\n{message}")
        self.progress_bar.setValue(100)
        self.run_button.setEnabled(True)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())


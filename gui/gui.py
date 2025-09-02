import sys
import os
import re
import datetime
import json
import traceback
import time
import pyhwpx
import anthropic
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold

from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, 
                             QProgressBar, QFileDialog, QGroupBox)
from PyQt5.QtCore import QThread, pyqtSignal

# 필수 라이브러리가 없을 경우를 대비한 안내
try:
    import google.generativeai as genai
    from google.generativeai.types import HarmCategory, HarmBlockThreshold
    import pyhwpx
except ImportError:
    print("="*60)
    print("필수 라이브러리가 설치되지 않았습니다.")
    print("터미널(명령 프롬프트)에 아래 명령어를 입력하여 설치해주세요.")
    print("pip install PyQt5 google-generativeai pyhwpx")
    print("="*60)
    sys.exit()

# -----------------------------------------------------------------------------
# 백그라운드에서 모든 자동화 작업을 처리하는 스레드 클래스
# -----------------------------------------------------------------------------
class DocumentProcessor(QThread):
    progress_update = pyqtSignal(str, int)
    finished = pyqtSignal(str)

    def __init__(self, keyword, api_key, hwp_path, pdf_paths):
        super().__init__()
        self.keyword = keyword
        self.api_key = api_key
        self.hwp_path = hwp_path
        self.pdf_paths = pdf_paths

    def run(self):
        hwp = None
        
        hwp = pyhwpx.Hwp()
        try:

            genai.configure(api_key=self.api_key)
            
            # --- HWP 파일 열기 ---
            if not os.path.exists(self.hwp_path):
                hwp.SaveAs(self.hwp_path)
            hwp.Open(self.hwp_path)

            # --- PDF 파일 업로드 ---
            self.progress_update.emit("PDF 파일 업로드를 시작합니다...", 10)
            uploaded_files = []
            num_pdfs = len(self.pdf_paths)
            for i, file_path in enumerate(self.pdf_paths):
                progress = 10 + int((i / num_pdfs) * 15)
                self.progress_update.emit(f"  - 업로드 중: {os.path.basename(file_path)}", progress)
                file_response = genai.upload_file(path=file_path)
                uploaded_files.append(file_response)
            self.progress_update.emit("PDF 파일 업로드 완료.", 25)

            # --- 모델 및 채팅 초기화 ---
            model = genai.GenerativeModel('gemini-2.5-flash', safety_settings={
                HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
            })
            chat = model.start_chat(history=[])

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
                self.progress_update.emit("AI 답변 파싱 및 삽입 중...", progress_start)
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
                        self.progress_update.emit("  - JSON 파싱 오류 발생. 이 단계는 건너뜁니다.", progress_start)
            # --- 5. 연도 자동 변경 ---
            self.progress_update.emit("연도를 변경합니다...", 95)
            this_year = datetime.date.today().year

            hwp.MoveDocBegin();
            while hwp.find('YEAR'):
                hwp.insert_text(str(this_year + 1))

            hwp.MoveDocBegin();
            while hwp.find('YDDY'):
                hwp.insert_text(str(this_year))

            # --- 대제목 생성 ---
            self.progress_update.emit("AI에게 대제목 생성 요청", 30)
            first_prompt =  f"""
                            PDF 내용을 참고하여 중복되는 업무 계획의 대제목을 생성해줘.
                            아래 [식별자 목록] 각각에 가장 적절한 제목을 한 줄로 할당해줘.
                            답변은 반드시 '## 식별자 제목' 형식으로만 생성해줘.
                            ** YOU DON'T MAKE 'YYY' TITLE**

                            [식별자 목록]
                            AAA, BBB, CCC, DDD, EEE, FFF, GGG, HHH, III, JJJ, KKK, LLL, ... 같은 규칙으로 순차적으로 늘어나도록 
                            """
            response = chat.send_message([*uploaded_files, first_prompt])
            if response.parts: process_text_response(response.text, 40)
            
            # --- 2. 세부내용 생성 ---
            self.progress_update.emit("AI에게 세부내용 생성 요청", 50)
            second_prompt = f"""
                            업로드 한 PDF 파일과 이전 답변을 바탕으로, '{self.keyword}' 키워드에 맞게 주제가 무너지지 않는 선에서 조화를 이루도록 세부 내용을 작성 해줘.

                            1. 기존에 생성한 핵심 주제들을 유지하면서, 새로운 키워드의 관점에서 내용을 재구성해줘.
                            2. 이전 대제목의 자식 식별자(AA1, AA2, BB1, BB2...)를 사용해줘.
                            3. **반드시 순수한 JSON 형태로만 출력해줘** (다른 설명 없이)
                            4. 이전 질문의 대제목을 잘 확인해서 연결해줘 AAA = AA1~AA6 파트 , BBB = BB1~BB6, ... 확인좀
                            5. YOU MUST MAINTAIN THAT I PROVIDED 'json'
                            6. You should only make as many as you can with a main title
                            7. 명사형종결체로 말해줘
                            - 아래 [JSON 출력 형식]을 완벽하게 따라줘.
                            


                            [JSON 출력 형식]
                            ```json
                            {{
                                "AA1": "한줄요약 내용",
                                "AA2": "성과목표 내용",
                                "AA3": "성과목표 내용",
                                "AA4": "소제목 내용",
                                "AA5": "추진배경 내용",
                                "AA6": "추진방향 내용",
                                "BB1": "한줄요약 내용",
                                "BB2": "성과목표 내용",
                                "BB3": "성과목표 내용",
                                "BB4": "소제목 내용",
                                "BB5": "추진배경 내용",
                                "BB6": "추진방향 내용",
                                "CC1": "한줄요약 내용",
                                "CC2": "성과목표 내용",
                                "CC3": "성과목표 내용",
                                "CC4": "소제목 내용",
                                "CC5": "추진배경 내용",
                                "CC6": "추진방향 내용",
                                "DD1": "한줄요약 내용",
                                "DD2": "성과목표 내용",
                                "DD3": "성과목표 내용",
                                "DD4": "소제목 내용",
                                "DD5": "추진배경 내용",
                                "DD6": "추진방향 내용",
                                "EE1": "한줄요약 내용",
                                "EE2": "성과목표 내용",
                                "EE3": "성과목표 내용",
                                "EE4": "소제목 내용",
                                "EE5": "추진배경 내용",
                                "EE6": "추진방향 내용",
                                "FF1": "한줄요약 내용",
                                "FF2": "성과목표 내용",
                                "FF3": "성과목표 내용",
                                "FF4": "소제목 내용",
                                "FF5": "추진배경 내용",
                                "FF6": "추진방향 내용",
                                "GG1": "한줄요약 내용",
                                "GG2": "성과목표 내용",
                                "GG3": "성과목표 내용",
                                "GG4": "소제목 내용",
                                "GG5": "추진배경 내용",
                                "GG6": "추진방향 내용",
                                "HH1": "한줄요약 내용",
                                "HH2": "성과목표 내용",
                                "HH3": "성과목표 내용",
                                "HH4": "소제목 내용",
                                "HH5": "추진배경 내용",
                                "HH6": "추진방향 내용",
                                "II1": "한줄요약 내용",
                                "II2": "성과목표 내용",
                                "II3": "성과목표 내용",
                                "II4": "소제목 내용",
                                "II5": "추진배경 내용",
                                "II6": "추진방향 내용",

                            }}
                            ``` 이형식 잘 유지해줘
                            """ 

            response = chat.send_message(second_prompt)
            if response.parts: process_json_response(response.text, 60)

            # --- 3. 주요 성과 생성 ---
            self.progress_update.emit("AI에게 주요 성과 생성 요청", 70)
            third_prompt =   f"""
                            지금까지의 PDF 내용을 종합해서, **2025년의 주요 성과**를 정리해줘.

                            [출력 형식]
                            - 답변은 반드시 순수한 JSON 형태로만 출력해줘 (다른 설명 없이).
                            - 각 성과 형식은 'JSON 출력 예시'의 형식에 있는 식별자를 사용해줘 (AC1, AC2, AC3...).
                            - 아래 [JSON 출력 예시]을 완벽하게 따라줘.
                            - 명사형종결체로 말해줘
                            - **'CC'식별자는 만들지마**

                            [JSON 출력 예시]
                            ```json
                            {{
                                "AC1": "1번째 주요 성과 내용 제목",
                                "AC2": "1번째 주요 성과 내용 요약",
                                "AC3": "2번째 주요 성과 내용 제목"
                                "AC4": "2번째 주요 성과 내용 요약"
                                "AC5": "3번째 주요 성과 내용 제목",
                                "AC6": "3번째 주요 성과 내용 요약",
                                "AC7": "4번째 주요 성과 내용 제목"
                                "AC8": "4번째 주요 성과 내용 요약"
                                "AC9": "5번째 주요 성과 내용 제목",
                                "BC1": "5번째 주요 성과 내용 요약",
                                "BC2": "6번째 주요 성과 내용 제목"
                                "BC3": "6번째 주요 성과 내용 요약"
                                "BC4": "7번째 주요 성과 내용 제목"
                                "BC5": "7번째 주요 성과 내용 요약"
                                "BC6": "8번째 주요 성과 내용 제목"
                                "BC7": "8번째 주요 성과 내용 요약"
                                "BC8": "9번째 주요 성과 내용 제목"
                                "BC9": "9번째 주요 성과 내용 요약"
                                "DC1": "10번째 주요 성과 내용 제목"
                                "DC2": "10번째 주요 성과 내용 요약"
                                "DC3": "11번째 주요 성과 내용 제목"
                                "DC4": "11번째 주요 성과 내용 요약"
                                "DC5": "12번째 주요 성과 내용 제목"
                                "DC6": "12번째 주요 성과 내용 요약"
                                "DC7": "13번째 주요 성과 내용 제목"
                                "DC8": "13번째 주요 성과 내용 요약"
                            }} 이형식 잘 유지해줘
                            ```
                            """
            
            response = chat.send_message(third_prompt)
            if response.parts: process_json_response(response.text, 80)
            
            # --- 4. 특수시책/핵심과제 생성 ---
            self.progress_update.emit("AI에게 특수시책/핵심과제 생성 요청", 85)
            fourth_prompt = f"""
                            지금까지의 PDF 내용을 종합해서, 2026년도에 할만한 특수시책이랑 핵심과제에 대해 제시해줘

                            [출력 형식]
                            - 답변은 반드시 순수한 JSON 형태로만 출력해줘 (다른 설명 없이).
                            - 각 성과 형식은 'JSON 출력 예시'의 형식에 있는 식별자를 사용해줘 (AC1, AC2, AC3...).
                            - 아래 [JSON 출력 예시]을 완벽하게 따라줘.
                            - 음슴체로 해줘 합니다. 말고 그냥 제공이면 제공. 확립이면 확립

                            [JSON 출력 예시]
                            ```json
                            {{
                                "J1": "1번째 핵심과제 제목",
                                "J2": "2번째 핵심과제 제목",
                                "H1": "1번째 특수시책 제목"
                            }} 이형식 잘 유지해줘
                            ```
                            """
            response = chat.send_message(fourth_prompt)
            if response.parts: process_json_response(response.text, 90)

            self.finished.emit(f"모든 작업이 완료되었습니다!\n결과 파일: {self.hwp_path}")

        except Exception:
            error_message = f"오류 발생:\n{traceback.format_exc()}"
            self.finished.emit(error_message)
        finally:
            base, ext = os.path.splitext(self.hwp_path)
            # 템플릿 파일이 아닌, 새로운 결과 파일 이름을 만듭니다.
            output_path = f"{base}_결과{ext}" 
            counter = 1
            # 동일한 파일명이 존재하면, 이름 뒤에 (숫자)를 붙입니다.
            while os.path.exists(output_path):
                output_path = f"{base}_결과 ({counter}){ext}"
                counter += 1
            
            hwp.save_as(output_path)
            # hwp.Quit()

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

        # --- 파일 선택 그룹 ---
        file_groupbox = QGroupBox("입력 파일 설정")
        file_vbox = QVBoxLayout()
        file_groupbox.setLayout(file_vbox)

        self.hwp_path_label = QLineEdit(self)
        self.hwp_path_label.setPlaceholderText("한글 템플릿 파일을 선택하세요.")
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
        
        # --- 키워드 입력 및 실행 ---
        run_groupbox = QGroupBox("실행")
        run_vbox = QVBoxLayout()
        run_groupbox.setLayout(run_vbox)

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

        # --- 진행 상황 그룹 ---
        progress_groupbox = QGroupBox("진행 상황")
        progress_vbox = QVBoxLayout()
        progress_groupbox.setLayout(progress_vbox)
        
        self.progress_bar = QProgressBar(self)
        self.status_log = QTextEdit(self)
        self.status_log.setReadOnly(True)
        self.status_log.setText("--- 대기 중 ---")
        progress_vbox.addWidget(self.progress_bar)
        progress_vbox.addWidget(self.status_log)

        # --- 메인 레이아웃에 그룹박스 추가 ---
        main_vbox.addWidget(file_groupbox)
        main_vbox.addWidget(run_groupbox)
        main_vbox.addWidget(progress_groupbox)
        
        self.setWindowTitle('LLM기반 문서 작성.v1.0')
        self.setGeometry(300, 300, 600, 500)
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
        
        if not all([keyword, hwp_path, self.pdf_paths]):
            self.status_log.setText("오류: 모든 파일(HWPX, PDF)을 선택하고 키워드를 입력해주세요.")
            return

        self.run_button.setEnabled(False)
        self.status_log.clear()
        self.progress_bar.setValue(0)

        API_KEY = "" # <<< 본인의 API 키를 꼭 입력해주세요!

        self.processor_thread = DocumentProcessor(keyword, API_KEY, hwp_path, self.pdf_paths)
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

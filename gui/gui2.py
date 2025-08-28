import sys
import os
import re
import datetime
import json
import traceback
import time

from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, 
                             QProgressBar, QFileDialog, QGroupBox)
from PyQt5.QtCore import QThread, pyqtSignal

# PyQT5가 없을 경우를 대비한 안내 메시지
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

    def __init__(self, keyword, api_key, hwp_path, pdf_paths, json_path):
        super().__init__()
        self.keyword = keyword
        self.api_key = api_key
        self.hwp_path = hwp_path
        self.pdf_paths = pdf_paths
        self.json_path = json_path

    def run(self):
        hwp = None 
        try:
            hwp = pyhwpx.Hwp()
            self.progress_update.emit("한글 프로그램을 시작합니다...", 5)

            genai.configure(api_key=self.api_key)

            if not os.path.exists(self.hwp_path):
                hwp.SaveAs(self.hwp_path)
            hwp.Open(self.hwp_path)

            # 파트 1: AI 대제목 생성
            self.progress_update.emit("--- 대제목 생성 시작 ---", 10)
            
            uploaded_files = []
            num_pdfs = len(self.pdf_paths)
            for i, file_path in enumerate(self.pdf_paths):
                progress = 15 + int((i / num_pdfs) * 20)
                self.progress_update.emit(f"PDF 업로드 중: {os.path.basename(file_path)}", progress)
                file_response = genai.upload_file(path=file_path)
                uploaded_files.append(file_response)
            self.progress_update.emit("PDF 파일 업로드 완료.", 35)

            model = genai.GenerativeModel('gemini-2.5-flash', safety_settings={
                HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
            })
            prompt_parts = [
                *uploaded_files,
                f"""
                PDF 내용을 참고하여 중복되는 업무 계획의 대제목을 생성해줘.
                아래 [식별자 목록] 각각에 가장 적절한 제목을 한 줄로 할당해줘.
                답변은 반드시 '## 식별자 제목' 형식으로만 생성해줘.

                [식별자 목록]
                AAA, BBB, CCC, DDD, EEE, FFF, GGG, HHH, III, JJJ, KKK, LLL
                """
            ]


            self.progress_update.emit("AI에게 대제목 생성을 요청...", 40)
            response = model.generate_content(prompt_parts, request_options={"timeout": 600})

            if response.parts:
                ai_text = response.text
                self.progress_update.emit("대제목 생성 완료. 문서 작성 시작", 50)

                this_year = datetime.date.today().year

                hwp.MoveDocBegin()
                while hwp.find('YYYY'): 
                    hwp.insert_text(str(this_year + 1)); 
                    time.sleep(0.2)

                hwp.MoveDocBegin()
                while hwp.find('YYYD'): 
                    hwp.insert_text(str(this_year));
                    time.sleep(0.2)

                sections = re.split(r'##\s*([A-Z]{3,})', ai_text)
                title_map = {}
                if len(sections) > 1:
                    for i in range(1, len(sections), 2):
                        marker = sections[i].strip()
                        title = sections[i+1].strip()
                        title_map[marker] = title
                
                for marker, title in title_map.items():
                    hwp.MoveDocBegin()
                    while hwp.find(marker):
                        hwp.insert_text(title)
                        time.sleep(0.2)


            # 파트 2: JSON 세부 내용 삽입
            if os.path.exists(self.json_path):
                with open(self.json_path, 'r', encoding='utf-8') as f:
                    detail_data_map = json.load(f)
                
                for main_marker, sub_items in detail_data_map.items():
                    for sub_marker, content in sub_items.items():
                        hwp.MoveDocBegin()
                        if hwp.find(sub_marker):
                            hwp.insert_text(content)
                            time.sleep(0.2)
            else:
                 self.progress_update.emit(f"[오류] JSON 파일을 찾을 수 없습니다: {self.json_path}", 85)

            # 파트 3: 연도 자동 변경

            
            self.finished.emit(f"모든 작업이 완료되었습니다!\n결과 파일: {self.hwp_path}")

        except Exception:
            error_message = f"오류 발생:\n{traceback.format_exc()}"
            self.finished.emit(error_message)
        finally:
                hwp.Save()
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

        input_groupbox = QGroupBox("입력 설정")
        input_vbox = QVBoxLayout()
        input_groupbox.setLayout(input_vbox)

        self.pdf_path_label = QLineEdit(self)
        self.pdf_path_label.setPlaceholderText("AI가 참고할 PDF 파일들을 선택하세요.")
        self.pdf_path_label.setReadOnly(True)
        pdf_button = QPushButton("PDF 파일 선택")
        pdf_button.clicked.connect(self.select_pdf_files)
        pdf_hbox = QHBoxLayout()
        pdf_hbox.addWidget(self.pdf_path_label)
        pdf_hbox.addWidget(pdf_button)
        input_vbox.addLayout(pdf_hbox)

        self.keyword_label = QLabel('핵심 키워드:')
        self.keyword_input = QLineEdit(self)
        self.keyword_input.setPlaceholderText('예: 2026년 주요 업무계획')
        keyword_hbox = QHBoxLayout()
        keyword_hbox.addWidget(self.keyword_label)
        keyword_hbox.addWidget(self.keyword_input)
        input_vbox.addLayout(keyword_hbox)
        
        self.run_button = QPushButton('문서 생성 시작', self)
        self.run_button.clicked.connect(self.start_processing)

        progress_groupbox = QGroupBox("진행 상황")
        progress_vbox = QVBoxLayout()
        progress_groupbox.setLayout(progress_vbox)
        
        self.progress_bar = QProgressBar(self)
        self.status_log = QTextEdit(self)
        self.status_log.setReadOnly(True)
        self.status_log.setText("--- 대기 중 ---")
        progress_vbox.addWidget(self.progress_bar)
        progress_vbox.addWidget(self.status_log)

        main_vbox.addWidget(input_groupbox)
        main_vbox.addWidget(self.run_button)
        main_vbox.addWidget(progress_groupbox)
        
        self.setWindowTitle('LLM 기반 문서 작성 지원프로그램 v1.0')
        self.setGeometry(300, 300, 600, 450)
        self.show()

    def select_pdf_files(self):
        fnames, _ = QFileDialog.getOpenFileNames(self, 'PDF 파일 선택 (다중 선택 가능)', '', 'PDF Files (*.pdf)')
        if fnames:
            self.pdf_paths = fnames
            self.pdf_path_label.setText(f"{len(fnames)}개 파일 선택됨: {os.path.basename(fnames[0])} 등")

    def start_processing(self):
        keyword = self.keyword_input.text()
        
        if not self.pdf_paths or not keyword:
            self.status_log.setText("오류: PDF 파일을 선택하고, 키워드를 입력해주세요.")
            return

        self.run_button.setEnabled(False)
        self.status_log.clear()
        self.progress_bar.setValue(0)

        # 고정 파일 경로 및 API 키 설정 (사용자 환경에 맞게 수정)
        API_KEY = ""
        hwp_path = r'C:\Users\USER\Desktop\gyeongji\0826\template.hwp'
        json_path = r'C:\Users\USER\Desktop\llm\LLM-based-document-writing-system\test.json'
        

        self.processor_thread = DocumentProcessor(keyword, API_KEY, hwp_path, self.pdf_paths, json_path)
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
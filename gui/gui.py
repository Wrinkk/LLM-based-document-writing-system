import sys
import os
import re
import datetime
import json
import traceback

from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTextEdit, QProgressBar
from PyQt5.QtCore import QThread, pyqtSignal

import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold
import pyhwpx

# -----------------------------------------------------------------------------
# 기존 로직을 별도의 스레드(Thread)에서 실행하기 위한 클래스
# -----------------------------------------------------------------------------
class DocumentProcessor(QThread):
    # 시그널 정의: (메시지, 진행률) 형태로 GUI에 업데이트를 보냄
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
        try:
            # --- HWP 객체 생성 ---
            hwp = pyhwpx.Hwp()
            self.progress_update.emit("한글 프로그램을 시작합니다...", 5)

            # --- API 설정 ---
            genai.configure(api_key=self.api_key)

            # --- HWP 파일 열기 ---
            if not os.path.exists(self.hwp_path):
                hwp.SaveAs(self.hwp_path)
            hwp.Open(self.hwp_path)

            # ==========================================================
            # 파트 1: AI 대제목 생성
            # ==========================================================
            self.progress_update.emit("--- 파트 1: AI 대제목 생성 시작 ---", 10)
            
            # PDF 업로드
            uploaded_files = []
            for i, file_path in enumerate(self.pdf_paths):
                if os.path.exists(file_path):
                    self.progress_update.emit(f"PDF 업로드 중: {os.path.basename(file_path)}", 15 + (i * 5))
                    file_response = genai.upload_file(path=file_path)
                    uploaded_files.append(file_response)
            self.progress_update.emit("PDF 파일 업로드 완료.", 35)

            # AI 요청
            model = genai.GenerativeModel('gemini-2.5-flash', safety_settings={
                            HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
                            HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_ONLY_HIGH,
                            HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
                            HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
            }) # 안전 설정
            prompt_parts = [
                *uploaded_files,
                        f"""
                            위에 제공된 PDF 파일 4개의 내용을 모두 참고해서 다음 작업을 수행해줘:

                            1. 먼저, 4개년 문서 전체에서 공통적으로 반복되는 **핵심 주제(테마)들을 찾아줘.** (예: 군정 기획 및 성과 관리, 재정 확보 및 운용, 적극 행정 및 주민 참여 등)

                            2. 그 다음, 네가 직접 찾아낸 이 **핵심 주제들을 대제목으로 사용**해서, 대제목에 대한 주요 업무 추진계획을 종합적으로 정리해줘.

                            3. 각 대제목 앞에는 반드시 '## ' (더블샵 +대제목식별자+공백)을 붙여서 답변을 생성해줘. (예: ## 군정 기획 및 성과 관리)

                            4. 제목과 세부 내용은 반드시 줄바꿈으로 분리해줘.

                            [출력 형식]
                            답변은 반드시 아래와 같은 형식으로만 생성해줘. 각 항목 앞에는 '## ' (더블샵+공백)을 붙여야 해.

                            ## AAA [AI가 생성한 세부 제목]
                            - 세부 내용 1
                            - 세부 내용 2

                            ## BBB [AI가 생성한 세부 제목]
                            - 세부 내용 1
                            - 세부 내용 2
                            ... 와 같이 계속

                            [대제목 식별자 목록]
                            AAA, BBB, CCC, DDD, EEE, FFF, GGG, HHH, III, JJJ, KKK, LLL, ... 같은 규칙으로 순차적으로 늘어나도록 
                        """
            ]
            self.progress_update.emit("AI에게 대제목 생성을 요청합니다...", 40)
            response = model.generate_content(prompt_parts, request_options={"timeout": 600})

            # AI 답변 파싱 및 대제목 삽입
            if response.parts:
                ai_text = response.text
                self.progress_update.emit("AI 대제목 생성 완료. 문서에 삽입을 시작합니다.", 50)
                sections = re.split(r'##\s*([A-Z]{3,})', ai_text)

                content_map = {}
                if len(sections) > 1:
                    for i in range(1, len(sections), 2):
                        marker = sections[i].strip()
                        full_content = sections[i+1].strip()
                        content_map[marker] = full_content                
                

                for marker, title in content_map.items():
                    hwp.MoveDocBegin()
                    if hwp.find(marker):
                        hwp.insert_text(title)
                        self.progress_update.emit(f"'{marker}'에 제목 삽입 완료.", 60)
            else:
                self.progress_update.emit("[AI 오류] AI가 대제목을 생성하지 않았습니다.", 60)

            # ==========================================================
            # 파트 2: JSON 세부 내용 삽입
            # ==========================================================
            self.progress_update.emit("--- 파트 2: JSON 세부 내용 삽입 시작 ---", 70)
            if os.path.exists(self.json_path):
                with open(self.json_path, 'r', encoding='utf-8') as f:
                    detail_data_map = json.load(f)
                
                for main_marker, sub_items in detail_data_map.items():
                    for sub_marker, content in sub_items.items():
                        hwp.MoveDocBegin()
                        if hwp.find(sub_marker):
                            hwp.insert_text(content)
                            self.progress_update.emit(f"'{sub_marker}'에 세부 내용 삽입 완료.", 80)
            else:
                 self.progress_update.emit(f"[오류] JSON 파일을 찾을 수 없습니다: {self.json_path}", 80)

            # ==========================================================
            # 파트 3: 연도 자동 변경
            # ==========================================================
            self.progress_update.emit("--- 파트 3: 연도 자동 변경 시작 ---", 90)
            next_year = str(datetime.date.today().year + 1)
            
            hwp.MoveDocBegin()
            while hwp.find('YYYY'): hwp.insert_text(next_year)
            hwp.MoveDocBegin()
            while hwp.find('YYYD'): hwp.insert_text(next_year)
            self.progress_update.emit("연도 자동 변경 완료.", 95)
            
            # --- 안전한 종료 ---
            hwp.Save()
            # hwp.Quit()
            self.finished.emit(f"모든 작업이 완료되었습니다!\n결과 파일: {self.hwp_path}")

        except Exception as e:
            # 상세한 오류 메시지를 GUI로 전달
            error_message = f"오류 발생:\n{traceback.format_exc()}"
            self.finished.emit(error_message)


# -----------------------------------------------------------------------------
# PyQT5 메인 GUI 애플리케이션 클래스
# -----------------------------------------------------------------------------
class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.processor_thread = None

    def initUI(self):
        # --- 위젯 생성 ---
        self.keyword_label = QLabel('핵심 키워드:')
        self.keyword_input = QLineEdit(self)
        self.keyword_input.setPlaceholderText('예: 문서 자동화')
        
        self.run_button = QPushButton('문서 생성 시작', self)
        
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        
        self.status_log = QTextEdit(self)
        self.status_log.setReadOnly(True)
        self.status_log.setText("--- 대기 중 ---")

        # --- 레이아웃 설정 ---
        vbox = QVBoxLayout()
        self.setLayout(vbox)

        hbox_input = QHBoxLayout()
        hbox_input.addWidget(self.keyword_label)
        hbox_input.addWidget(self.keyword_input)

        vbox.addLayout(hbox_input)
        vbox.addWidget(self.run_button)
        vbox.addWidget(self.progress_bar)
        vbox.addWidget(self.status_log)
        
        # --- 이벤트 연결 ---
        self.run_button.clicked.connect(self.start_processing)

        # --- 창 설정 ---
        self.setWindowTitle('AI 문서 자동 생성 프로그램')
        self.setGeometry(300, 300, 500, 400)
        self.show()

    def start_processing(self):
        keyword = self.keyword_input.text()
        if not keyword:
            self.status_log.setText("오류: 핵심 키워드를 입력해주세요.")
            return

        # 버튼 비활성화 (중복 실행 방지)
        self.run_button.setEnabled(False)
        self.status_log.clear()
        self.progress_bar.setValue(0)

        # 사용자 설정 (이 부분은 실제 환경에 맞게 수정해야 합니다)
        API_KEY = ""
        hwp_path = r'C:\Users\USER\Desktop\gyeongji\0826\template.hwp'
        json_path = r'C:\Users\USER\Desktop\gyeongji\0826\llm\LLM-based-document-writing-system\backend\models\test.json'
        pdf_files_paths = [
            r"C:\Users\USER\Desktop\gyeongji\0826\2021.pdf",
            r"C:\Users\USER\Desktop\gyeongji\0826\2022.pdf",
            r"C:\Users\USER\Desktop\gyeongji\0826\2023.pdf",
            r"C:\Users\USER\Desktop\gyeongji\0826\2024.pdf"
        ]


        # 스레드 생성 및 시작
        self.processor_thread = DocumentProcessor(keyword, API_KEY, hwp_path, pdf_files_paths, json_path)
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
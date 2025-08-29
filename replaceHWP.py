# import pyhwpx
# import google.generativeai as genai

# hwp = pyhwpx.Hwp()
# hwp.Open(r'C:\Users\USER\Desktop\gyeongji\0826\sample2.hwpx')

# prompt_keyword = input("키워드 입력: " )

# genai.configure(api_key="")

# model = genai.GenerativeModel('gemini-2.5-flash') 

# response = model.generate_content(prompt_keyword + "좀 더 고급스럽게 해줘")

# print(response.text)
# hwp.MoveDocBegin()
# hwp.find("AAA")
# hwp.insert_text(response.text)

import pyhwpx
import google.generativeai as genai
import os
import re
from google.generativeai.types import HarmCategory, HarmBlockThreshold
import datetime 
import json 
import time

hwp = pyhwpx.Hwp()

try:

    # [사용자 설정 1] 작업할 HWPX 템플릿 파일 경로
    hwp_path = r'C:\Users\USER\Desktop\gyeongji\0826\template.hwp'
    if not os.path.exists(hwp_path):
        print(f"템플릿 파일이 없어 새로 생성합니다: {hwp_path}")
        hwp.SaveAs(hwp_path)
    hwp.Open(hwp_path)

    # [사용자 설정 2] 본인의 Gemini API 키 입력
    API_KEY = "" 
    genai.configure(api_key=API_KEY)

    # [사용자 설정 3] AI가 참고할 PDF 파일들의 경로
    pdf_files_paths = [
        r"C:\Users\USER\Desktop\gyeongji\0826\2021.pdf",
        r"C:\Users\USER\Desktop\gyeongji\0826\2022.pdf",
        r"C:\Users\USER\Desktop\gyeongji\0826\2023.pdf",
        r"C:\Users\USER\Desktop\gyeongji\0826\2024.pdf",
        r"C:\Users\USER\Desktop\gyeongji\0826\2025.pdf"
    ]

    # --- 2. PDF 파일 업로드 ---
    uploaded_files = []
    print("PDF 파일 업로드를 시작합니다...")
    for file_path in pdf_files_paths:
        if os.path.exists(file_path):
            print(f"  - 업로드 중: {os.path.basename(file_path)}")
            file_response = genai.upload_file(path=file_path)
            uploaded_files.append(file_response)
            print(f"  - 업로드 완료: {file_response.display_name}")
        else:
            print(f"  - 경고: '{file_path}' 파일을 찾을 수 없어 건너뜁니다.")
    print("PDF 파일 업로드가 완료되었습니다.\n")

    # --- 3. 사용자 입력 및 AI 요청 ---
    prompt_keyword = input("키워드 입력: ")

    model = genai.GenerativeModel(
        'gemini-2.5-flash',
        safety_settings={
            HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
            HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_ONLY_HIGH,
            HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
            HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
        }
    )
    # '{prompt_keyword}'
    # AI가 답변을 파싱하기 좋은 형태로 생성하도록 요청 프롬프트를 구체화
    prompt_parts = [
        *uploaded_files,
        f"""
        위에 제공된 PDF 파일 5개의 내용을 모두 참고해서 다음 작업을 수행해줘:

        1. 먼저, 5개년 문서 전체에서 공통적으로 반복되는 **핵심 주제(테마)들을 찾아줘.** (예: 군정 기획 및 성과 관리, 재정 확보 및 운용, 적극 행정 및 주민 참여, 감사 및 공직 윤리 확립,정보통신 인프라 구축 및 관리, 디지털 역량 강화 및 포용, 군정 홍보 및 지역 이미지 제고, 법무·송무 및 규제 개선 등)

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

    print("\nAI에게 답변 생성을 요청합니다...")
    response = model.generate_content(prompt_parts, request_options={"timeout": 600})
    history = model.start_chat(history=[]) 
    
    # --- 4. AI 답변 파싱 및 HWPX에 순차 삽입 ---
    if response.parts:
        ai_text = response.text
        print("\n--- AI 생성 답변 ---\n", ai_text)

        # '## AAA' 등을 기준으로 텍스트를 분리
        sections = re.split(r'##\s*([A-Z]{3,})', ai_text)
        print (sections)
        
        content_map = {}
        if len(sections) > 1:
            for i in range(1, len(sections), 2):
                marker = sections[i].strip()
                full_content = sections[i+1].strip()
                content_map[marker] = full_content

        print("\nHWPX 파일에 파싱된 답변을 삽입합니다...")
        for marker, full_content in content_map.items():
            # 전체 내용에서 첫 번째 줄(제목)만 추출
            title_only = full_content.split('\n')[0].strip()

            hwp.MoveDocBegin() 
            while hwp.find(marker):
                hwp.insert_text(title_only)
                time.sleep(0.1)
                print(f"  - 성공: '{marker}' 위치에 제목 '{title_only}'을(를) 삽입했습니다.")
       
    else:
        print("\n[오류] AI가 답변을 생성하지 않았습니다. 안전 필터에 의해 차단되었을 수 있습니다.")
        print("차단 피드백:", response.prompt_feedback)

    hwp.MoveDocBegin() 
    while (hwp.find('YYYY')):
        date = datetime.date.today()
        years = str(date.year + 1)
        hwp.insert_text(years)

    hwp.MoveDocBegin() 
    while (hwp.find('YYYD')):
        date = datetime.date.today()
        years = str(date.year)
        hwp.insert_text(years)


    # 1. JSON 파일 경로 설정
    json_path = r'C:\Users\USER\Desktop\llm\LLM-based-document-writing-system\test.json' 

    # 2. JSON 파일 불러오기 및 파싱
    if not os.path.exists(json_path):
        print(f"[오류] JSON 파일을 찾을 수 없습니다: {json_path}")
    else:
        with open(json_path, 'r', encoding='utf-8') as f:
            detail_data_map = json.load(f)

        # 3. HWPX 문서에 파싱된 내용 순차 삽입 (Replace)
        for main_marker, sub_items in detail_data_map.items():
            for sub_marker, content in sub_items.items():
                hwp.MoveDocBegin()
                print(sub_marker) 
                print(content) 

                while hwp.find(sub_marker):
                    hwp.insert_text(content)
                    time.sleep(0.1)



finally:
    # --- 5. 안전한 종료 ---
    # try 블록에서 오류가 발생하든 안 하든, 이 부분은 반드시 실행됩니다.
    print("\n프로그램을 안전하게 종료합니다...")
    # hwp.Save()
    # hwp.Quit()

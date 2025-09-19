import pyhwpx
import google.generativeai as genai
import os
import re
from google.generativeai.types import HarmCategory, HarmBlockThreshold
import datetime 
import json 
import time

hwp = pyhwpx.Hwp()


def process_ai_response(ai_text, question_num):
        """AI 답변을 파싱하고 HWPX에 삽입하는 함수"""
        print(f"\n--- {question_num}번째 질문 AI 생성 답변 ---\n", ai_text)

        # '## AAA' 등을 기준으로 텍스트를 분리
        sections = re.split(r'##\s*([A-Z]{3,})', ai_text)
        print(f"{question_num}번째 질문 파싱 결과:", sections)
        
        content_map = {}
        if len(sections) > 1:
            for i in range(1, len(sections), 2):
                marker = sections[i].strip()
                full_content = sections[i+1].strip()
                content_map[marker] = full_content

        print(f"\nHWPX 파일에 {question_num}번째 질문 파싱된 답변을 삽입합니다...")
        for marker, full_content in content_map.items():
            # 전체 내용에서 첫 번째 줄(제목)만 추출
            title_only = full_content.split('\n')[0].strip()

            hwp.MoveDocBegin() 
            while hwp.find(marker):
                hwp.insert_text(title_only)
                time.sleep(0.1)
                print(f"  - 성공: '{marker}' 위치에 제목 '{title_only}'을(를) 삽입했습니다.")

def extract_json_from_text(text):
        """텍스트에서 JSON 부분만 추출하는 함수"""
        # JSON 코드 블록에서 추출 시도
        json_match = re.search(r'```json\s*(\{.*?\})\s*```', text, re.DOTALL)
        if json_match:
            return json_match.group(1)
        
        # 일반적인 중괄호로 둘러싸인 JSON 추출 시도
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            return json_match.group(0)
        
        return None

def process_json_response(ai_text, question_num):
        """JSON 형태의 AI 답변을 파싱하고 HWPX에 삽입하는 함수"""
        print(f"\n--- {question_num}번째 질문 AI JSON 답변 처리 ---")
        
        # JSON 추출
        json_text = extract_json_from_text(ai_text)
        if not json_text:
            print("JSON 형태를 찾을 수 없습니다. 일반 텍스트로 처리합니다.")
            process_ai_response(ai_text, question_num)
            return
        
        try:
            # JSON 파싱
            json_data = json.loads(json_text)
            print(f"JSON 파싱 성공: {len(json_data)}개 항목")
            
            # JSON 데이터를 파일로 저장
            json_save_path = r'C:\Users\USER\Desktop\llm\LLM-based-document-writing-system\ai_response.json'
            with open(json_save_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            print(f"JSON 데이터를 파일로 저장: {json_save_path}")
            
            # HWPX에 JSON 데이터 삽입
            print(f"\nHWPX 파일에 JSON 데이터를 삽입합니다...")
            for marker, content in json_data.items():
                hwp.MoveDocBegin()
                print(f"  - 처리 중: {marker}")
                
                while hwp.find(marker):
                    hwp.insert_text(content)
                    time.sleep(0.1)
                    print(f"  - 성공: '{marker}' 위치에 '{content}'을(를) 삽입했습니다.")
                    
        except json.JSONDecodeError as e:
            print(f"JSON 파싱 오류: {e}")
            print("일반 텍스트로 처리합니다.")
            process_ai_response(ai_text, question_num)

try:

    # [사용자 설정 1] 작업할 HWPX 템플릿 파일 경로
    hwp_path = r'C:\Users\USER\Desktop\gyeongji\0826\template.hwp'
    if not os.path.exists(hwp_path):
        print(f"템플릿 파일이 없어 새로 생성합니다: {hwp_path}")
        hwp.SaveAs(hwp_path)
    hwp.Open(hwp_path)

    # [사용자 설정 2] 본인의 Gemini API 키 입력
    API_KEY = "AIzaSyDz5V-imSFlRcC8dfNI_fmQpmEXOotMxX0" 
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

    # --- 3. 모델 초기화 ---
    model = genai.GenerativeModel(
        'gemini-2.5-flash',
        safety_settings={
            HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
            HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_ONLY_HIGH,
            HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
            HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
        }
    )
    
    # 채팅 세션 시작 (대화 히스토리 유지를 위해)
    chat = model.start_chat(history=[])
    

    # --- 날짜 처리  ---
    hwp.MoveDocBegin() 
    while (hwp.find('YEAR')):
        date = datetime.date.today()
        years = str(date.year + 1)
        hwp.insert_text(years)

    hwp.MoveDocBegin() 
    while (hwp.find('YDDY')):
        date = datetime.date.today()
        years = str(date.year)
        hwp.insert_text(years)

    # --- 4-1. 첫 번째 질문 및 AI 요청 ---
    prompt_keyword = input("첫 번째 키워드 입력: ")

    # 첫 번째 질문 프롬프트
    first_prompt_parts = [
        *uploaded_files,
        f"""
        PDF 내용을 참고하여 중복되는 업무 계획의 대제목을 생성해줘.
        아래 [식별자 목록] 각각에 가장 적절한 제목을 한 줄로 할당해줘.
        답변은 반드시 '## 식별자 제목' 형식으로만 생성해줘.
        YOU DON'T MAKE 'YYY' TITLE

        [식별자 목록]
        AAA, BBB, CCC, DDD, EEE, FFF, GGG, HHH, III, JJJ, KKK, LLL, ... 같은 규칙으로 순차적으로 늘어나도록 
        """
    ]

    print("\nAI에게 답변 생성을 요청...")
    first_response = chat.send_message(first_prompt_parts, request_options={"timeout": 600})
    
    # 첫 번째 답변 처리
    if first_response.parts:
        process_ai_response(first_response.text, 1)
    else:
        print("\n[오류] 첫 번째 질문에서 AI가 답변을 생성하지 않았습니다.")
        print("차단 피드백:", first_response.prompt_feedback)

    # --- 4-2. 두 번째 질문 및 AI 요청 (JSON 처리) ---
    print("\n" + "="*50)
    second_keyword = prompt_keyword
    print(chat.history)
    
    if second_keyword.strip():  
        # 두 번째 질문 프롬프트 (JSON 형태 요청)
        second_prompt = [
            *uploaded_files,
            f"""
            업로드 한 PDF 파일과 이전 답변을 바탕으로, 키워드에 맞게 주제가 무너지지 않는 선에서 세부 내용을 작성 해줘: '{second_keyword}'

            1. 기존에 생성한 핵심 주제들을 유지하면서, 새로운 키워드의 관점에서 내용을 재구성해줘.
            2. 이전 대제목의 자식 식별자(AA1, AA2, BB1, BB2...)를 사용해줘.
            3. **반드시 순수한 JSON 형태로만 출력해줘** (다른 설명 없이)
            4. 이전 질문의 대제목을 잘 확인해서 연결해줘 AAA = AA1~AA6 파트 , BBB = BB1~BB6, ... 확인좀
            5. YOU MUST MAINTAIN THAT I PROVIDED 'json'
            6. You should only make as many as you can with a main title
            - 아래 [JSON 출력 형식]을 완벽하게 따라줘.
            
            새로운 키워드: '{second_keyword}'

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
        ]

        print("\nAI에게 답변 생성을 요청합니다...")
        second_response = chat.send_message(second_prompt, request_options={"timeout": 600})
        
        # 두 번째 답변을 JSON으로 처리
        if second_response.parts:
            process_json_response(second_response.text, 2)
            print(second_response.text)
        else:
            print("\n[오류] 두 번째 질문에서 AI가 답변을 생성하지 않았습니다.")
            print("차단 피드백:", second_response.prompt_feedback)
    else:
        print("")

    # --- 4-3. 세 번째 질문: 2025년 주요 성과 정리 (JSON) ---
    print("\n" + "="*50)
    
    # 세 번째 질문을 위한 프롬프트
    third_prompt = [
        *uploaded_files,
        f"""
        지금까지의 PDF 내용을 종합해서, **2025년의 주요 성과**를 정리해줘.
        
        [출력 형식]
        - 답변은 반드시 순수한 JSON 형태로만 출력해줘 (다른 설명 없이).
        - 각 성과 형식은 'JSON 출력 예시'의 형식에 있는 식별자를 사용해줘 (AC1, AC2, AC3...).
        - 아래 [JSON 출력 예시]을 완벽하게 따라줘.
        - 음슴체로 해줘 합니다. 말고 그냥 제공이면 제공. 확립이면 확립
        - 'CC'식별자는 만들지마

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
    ]

    third_response = chat.send_message(third_prompt, request_options={"timeout": 600})
    
    # 세 번째 답변을 JSON으로 처리
    if third_response.parts:
        process_json_response(third_response.text, 3)
    else:
        print("\n[오류] 질문에 AI가 답변하지 않았습니다.")
        print("차단 피드백:", third_response.prompt_feedback)

        # --- 4-3. 세 번째 질문: 2025년 주요 성과 정리 (JSON) ---
    print("\n" + "="*50)
    
    # 4번째 질문을 위한 프롬프트
    fourth_prompt = [
        *uploaded_files,
        f"""
        지금까지의 PDF 내용을 종합해서, 2026년도 특수시책이랑 핵심과제를 적어줘
        
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
    ]

    fourth_response = chat.send_message(fourth_prompt, request_options={"timeout": 600})
    
    # 세 번째 답변을 JSON으로 처리
    if fourth_response.parts:
        process_json_response(fourth_response.text, 3)
    else:
        print("\n[오류] 세 번째 질문에 AI가 답변하지 않았습니다.")
        print("차단 피드백:", fourth_response.prompt_feedback)



    # # --- 6. 기존 JSON 파일 처리 (선택사항) ---
    # json_path = r'C:\Users\USER\Desktop\llm\LLM-based-document-writing-system\test.json' 

    # if os.path.exists(json_path):
    #     print(f"\n기존 JSON 파일도 처리합니다: {json_path}")
    #     with open(json_path, 'r', encoding='utf-8') as f:
    #         detail_data_map = json.load(f)

    #     # HWPX 문서에 파싱된 내용 순차 삽입 (Replace)
    #     print("기존 JSON 데이터를 HWPX에 삽입합니다...")
    #     for main_marker, sub_items in detail_data_map.items():
    #         for sub_marker, content in sub_items.items():
    #             hwp.MoveDocBegin()
    #             print(f"  - 처리 중: {sub_marker}")

    #             while hwp.find(sub_marker):
    #                 hwp.insert_text(content)
    #                 time.sleep(0.1)
    # else:
    #     print(f"기존 JSON 파일이 없습니다: {json_path}")

finally:
    # --- 7. 안전한 종료 ---
    print("\n프로그램을 안전하게 종료합니다...")
    hwp.save_as("output")
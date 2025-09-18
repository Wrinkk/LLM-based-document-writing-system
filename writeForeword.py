import pyhwpx
import google.generativeai as genai
import os
import re
from google.generativeai.types import HarmCategory, HarmBlockThreshold
import datetime 
import json 
import time
import PyPDF2
from io import StringIO

hwp = pyhwpx.Hwp()

def extract_text_from_pdf(pdf_path):
    """PDF 파일에서 텍스트를 추출하는 함수"""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text += page.extract_text() + "\n"
        return text
    except Exception as e:
        print(f"PDF 텍스트 추출 오류 ({pdf_path}): {e}")
        return ""

def combine_pdf_texts(pdf_files_paths):
    """여러 PDF 파일의 텍스트를 하나로 합치는 함수"""
    combined_text = ""
    
    print("PDF 파일에서 텍스트를 추출하고 통합합니다...")
    for file_path in pdf_files_paths:
        if os.path.exists(file_path):
            print(f"  - 텍스트 추출 중: {os.path.basename(file_path)}")
            pdf_text = extract_text_from_pdf(file_path)
            
            if pdf_text.strip():
                combined_text += f"\n\n=== {os.path.basename(file_path)} ===\n"
                combined_text += pdf_text
                print(f"  - 추출 완료: {len(pdf_text)} 문자")
            else:
                print(f"  - 경고: '{file_path}'에서 텍스트를 추출할 수 없습니다.")
        else:
            print(f"  - 경고: '{file_path}' 파일을 찾을 수 없어 건너뜁니다.")
    
    print(f"통합된 텍스트 총 길이: {len(combined_text)} 문자\n")
    
    # 통합된 텍스트를 파일로 저장 (선택사항)
    combined_text_path = r'C:\Users\USER\Desktop\gyeongji\0826\combined_pdfs.txt'
    try:
        with open(combined_text_path, 'w', encoding='utf-8') as f:
            f.write(combined_text)
        print(f"통합된 텍스트를 파일로 저장: {combined_text_path}")
    except Exception as e:
        print(f"텍스트 파일 저장 오류: {e}")
    
    return combined_text

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

    # --- 2. PDF 텍스트 통합 ---
    combined_pdf_text = combine_pdf_texts(pdf_files_paths)
    
    if not combined_pdf_text.strip():
        print("경고: 추출된 PDF 텍스트가 없습니다. 프로그램을 종료합니다.")
        exit()

    # --- 3. 모델 초기화 ---
    model = genai.GenerativeModel(
        'gemini-1.5-pro',  # 긴 텍스트 처리를 위해 pro 모델 사용
        safety_settings={
            HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
            HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_ONLY_HIGH,
            HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
            HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
        }
    )
    
    # 채팅 세션 시작 (대화 히스토리 유지를 위해)
    chat = model.start_chat(history=[])

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

    # 첫 번째 질문 프롬프트 (통합된 PDF 텍스트 사용)
    first_prompt = f"""
다음은 여러 PDF 파일에서 추출한 통합 텍스트 내용입니다:

{combined_pdf_text}

위 내용을 참고하여 중복되는 업무 계획의 대제목을 생성해줘.
아래 [식별자 목록] 각각에 가장 적절한 제목을 한 줄로 할당해줘.
답변은 반드시 '## 식별자 제목' 형식으로만 생성해줘.
YOU DON'T MAKE 'YYY' TITLE

[식별자 목록]
AAA, BBB, CCC, DDD, EEE, FFF, GGG, HHH, III, JJJ, KKK, LLL, ... 같은 규칙으로 순차적으로 늘어나도록 
"""

    print("\nAI에게 답변 생성을 요청...")
    first_response = chat.send_message(first_prompt, request_options={"timeout": 600})
    
    # 첫 번째 답변 처리
    if first_response.parts:
        process_ai_response(first_response.text, 1)
    else:
        print("\n[오류] 첫 번째 질문에서 AI가 답변을 생성하지 않았습니다.")
        print("차단 피드백:", first_response.prompt_feedback)

    # --- 4-2. 두 번째 질문 및 AI 요청 (JSON 처리) ---
    print("\n" + "="*50)
    second_keyword = prompt_keyword
    print("채팅 히스토리:", len(chat.history), "개 메시지")
    
    if second_keyword.strip():  
        # 두 번째 질문 프롬프트 (JSON 형태 요청)
        second_prompt = f"""
다음은 여러 PDF 파일에서 추출한 통합 텍스트 내용입니다:

{combined_pdf_text}

위 PDF 내용과 이전 답변을 바탕으로, 키워드에 맞게 주제가 무너지지 않는 선에서 세부 내용을 작성 해줘: '{second_keyword}'

1. 기존에 생성한 핵심 주제들을 유지하면서, 새로운 키워드의 관점에서 내용을 재구성해줘.
2. 이전 대제목의 자식 식별자(AA1, AA2, BB1, BB2...)를 사용해줘.
3. **반드시 순수한 JSON 형태로만 출력해줘** (다른 설명 없이)
4. 이전 질문의 대제목을 잘 확인해서 연결해줘 AAA = AA1~AA6 파트 , BBB = BB1~BB6, ... 확인좀
5. YOU MUST MAINTAIN THAT I PROVIDED 'json'
6. You should only make as many as you can with a main title
- 아래 [JSON 출력 형식]을 완벽하게 따라줘.

새로운 키워드: '{second_keyword}'

[JSON 출력 형식]

""" 

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
        print("두 번째 키워드가 입력되지 않았습니다.")

    # --- 4-3. 세 번째 질문: 2025년 주요 성과 정리 (JSON) ---
    print("\n" + "="*50)
    
    # 세 번째 질문을 위한 프롬프트
    third_prompt = f"""
여러 PDF 파일에서 추출한 통합 텍스트 내용입니다:

{combined_pdf_text}

"""

    third_response = chat.send_message(third_prompt, request_options={"timeout": 600})
    
    # 세 번째 답변을 JSON으로 처리
    if third_response.parts:
        process_json_response(third_response.text, 3)
    else:
        print("\n[오류] 세 번째 질문에 AI가 답변하지 않았습니다.")
        print("차단 피드백:", third_response.prompt_feedback)

    # --- 4-4. 네 번째 질문: 2026년도 특수시책 및 핵심과제 ---
    print("\n" + "="*50)
    
    # 4번째 질문을 위한 프롬프트
    fourth_prompt = f"""
        다음은 여러 PDF 파일에서 추출한 통합 텍스트 내용입니다:

        {combined_pdf_text}

            """

    fourth_response = chat.send_message(fourth_prompt, request_options={"timeout": 600})
    
    # 네 번째 답변을 JSON으로 처리
    if fourth_response.parts:
        process_json_response(fourth_response.text, 4)
    else:
        print("\n[오류] 네 번째 질문에 AI가 답변하지 않았습니다.")
        print("차단 피드백:", fourth_response.prompt_feedback)

finally:
    # --- 7. 안전한 종료 ---
    print("\n프로그램을 안전하게 종료합니다...")
    hwp.save_as("output")
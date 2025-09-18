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


finally:
    # --- 7. 안전한 종료 ---
    print("\n프로그램을 안전하게 종료합니다...")
    hwp.save_as("output")
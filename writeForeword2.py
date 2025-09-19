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

def combine_pdf_texts(pdf_files_paths):
    """여러 PDF 파일의 텍스트를 효율적으로 합치는 함수 (추출과 정제를 한 번에 처리)"""
    combined_text = ""
    
    print("PDF 파일에서 텍스트 추출 및 정제를 시작합니다...")
    for file_path in pdf_files_paths:
        if not os.path.exists(file_path):
            print(f"  - 파일 찾을 수 없음: '{os.path.basename(file_path)}'")
            continue

        print(f"  - 처리 중: {os.path.basename(file_path)}")
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                raw_text = ""
                for page in pdf_reader.pages:
                    if page.extract_text():
                        raw_text += page.extract_text() + "\n"
                
                if not raw_text.strip():
                    print(f"  - 경고: '{os.path.basename(file_path)}'에서 텍스트를 추출 불가")
                    continue

                cleaned_text = re.sub(r'[^가-힣A-Za-z0-9\s\.,\?!()\[\]{}:;\'"""''·]', '', raw_text)
                
                combined_text += f"\n\n=== {os.path.basename(file_path)}의 내용 ===\n"
                combined_text += cleaned_text
                print(f"  - 완료: {len(cleaned_text)} 문자")

        except Exception as e:
            print(f"  - 오류: '{os.path.basename(file_path)}' 처리 중 오류 발생: {e}")
            
    print(f"\n통합 및 정제된 텍스트 총 길이: {len(combined_text)} 문자")
    return combined_text

def wait_for_file_processing(uploaded_files):
    """업로드된 파일들의 처리 완료를 대기하는 함수"""
    print("\n업로드된 파일들의 처리 완료를 대기 중...")
    for file in uploaded_files:
        while file.state.name == "PROCESSING":
            print(f"  - {file.display_name} 처리 중... 10초 대기")
            time.sleep(10)
            file = genai.get_file(file.name)
        
        if file.state.name == "FAILED":
            raise ValueError(f"파일 처리 실패: {file.display_name}")
        
        print(f"  - {file.display_name} 처리 완료")

def create_hwp_document_with_foreword(template_path, foreword_text, output_path, version_num):
    """발간사가 포함된 한글 문서를 생성하는 함수"""
    try:
        hwp = pyhwpx.Hwp(visible=False) 
        
        # 템플릿 파일 열기
        if os.path.exists(template_path):
            hwp.Open(template_path)
        else:
            # 새 문서 생성
            hwp.XHwpDocuments.Add()
        
        # 문서 시작으로 이동
        hwp.MoveDocBegin()
        
        # 발간사 제목 삽입
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = f"발간사 (버전 {version_num})\n\n"
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
        
        # 발간사 내용 삽입
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = foreword_text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
        
        # 'TEST' 텍스트가 있으면 발간사로 교체
        hwp.MoveDocBegin()
        while hwp.find('TEST'):
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = foreword_text
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
        
        # 파일 저장
        hwp.SaveAs(output_path)
        hwp.Quit()
        
        return True, f"성공적으로 저장됨: {output_path}"
        
    except Exception as e:
        try:
            hwp.Quit()
        except:
            pass
        return False, f"오류 발생: {e}"

# 전역 변수
hwp = None
uploaded_files = []

try:
    # [사용자 설정 1] 작업할 HWPX 템플릿 파일 경로
    hwp_path = r'C:\Users\USER\Desktop\gyeongji\0826\template.hwp'
    
    # 결과 파일들을 저장할 디렉토리 생성
    output_dir = os.path.join(os.path.dirname(hwp_path), "발간사_결과")
    os.makedirs(output_dir, exist_ok=True)
    
    # [사용자 설정 2] 본인의 Gemini API 키 입력
    API_KEY = ""  # 여기에 실제 API 키를 입력하세요
    
    genai.configure(api_key=API_KEY)

    # 텍스트를 추출할 PDF 파일 목록
    text_extract_paths = [
        r"C:\Users\USER\Downloads\hwp\희망울진 군정집(2025년 1분기).pdf",
        r"C:\Users\USER\Downloads\hwp\희망울진 군정집(2025년 2분기)_압.pdf",
        r"C:\Users\USER\Downloads\hwp\희망울진 군정집(2025년 3분기).pdf",
    ]

    # 파일 채로 AI에 업로드할 PDF 파일 목록
    file_upload_paths = [
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(건설과).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(경제교통과).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(기획예산실)(9.8. 수정).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(농기계임대사업소).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(농업기술센터)(수정).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(농정과).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(도시새마을과).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(맑은물사업소).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(문화관광과).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(민원과)(9.8. 수정).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(보건소).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(복지정책과).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(사회복지과).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(산림과).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(수소국가산업추진단).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(안전재난과).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(왕피천공원사업소).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(울진군의료원).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(원전에너지과).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(인구정책과).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(재무과)(9.8. 수정).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(정책홍보실)(9.11. 수정).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(체육진흥과).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(총무과)(9.8. 수정).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(해양수산과).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(환경위생과).pdf",
        r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(환동해산업연구원).pdf"
    ]

    # --- 1. 텍스트 추출 그룹 처리 ---
    combined_pdf_text = ""
    if text_extract_paths:
        combined_pdf_text = combine_pdf_texts(text_extract_paths)

    # --- 2. 파일 업로드 그룹 처리 ---
    if file_upload_paths:
        print("\nPDF 파일을 AI에 직접 업로드합니다...")
        for file_path in file_upload_paths:
            if os.path.exists(file_path):
                print(f"  - 업로드 중: {os.path.basename(file_path)}")
                try:
                    file_size = os.path.getsize(file_path)
                    if file_size > 200 * 1024 * 1024:  # 200MB
                        print(f"  - 경고: '{os.path.basename(file_path)}'가 200MB를 초과하여 건너뜁니다.")
                        continue
                    
                    file_response = genai.upload_file(path=file_path)
                    uploaded_files.append(file_response)
                    print(f"  - 업로드 완료: {file_response.display_name}")
                except Exception as e:
                    print(f"  - 업로드 실패: {os.path.basename(file_path)} - {e}")
            else:
                print(f"  - 경고: '{file_path}' 파일을 찾을 수 없어 건너뜁니다.")
        
        if uploaded_files:
            wait_for_file_processing(uploaded_files)
            print("파일 업로드가 완료되었습니다.")
        else:
            print("업로드된 파일이 없습니다.")

    # --- 3. 모델 초기화 ---
    try:
        model = genai.GenerativeModel(
            'gemini-2.5-flash',
            safety_settings={
                HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
            }
        )
    except Exception as e:
        print(f"모델 초기화 실패: {e}")
        model = genai.GenerativeModel('gemini-1.5-flash')

    year = datetime.datetime.now().year
    quarter = ((datetime.datetime.now().month - 1) // 3 + 1 ) + 1

    # --- 4. 5개의 다른 발간사 생성 ---
    foreword_variations = [
        {
            "focus": "경제 발전 중심",
            "tone": "역동적이고 진취적인 어조",
            "emphasis": "울진군의 미래 성장 동력과 경제 발전 전략"
        },
        {
            "focus": "군민 화합 중심", 
            "tone": "따뜻하고 포용적인 어조",
            "emphasis": "군민과의 소통과 협력, 상생의 가치"
        },
        {
            "focus": "희망과 비전 중심",
            "tone": "미래지향적이고 희망찬 어조", 
            "emphasis": "울진군의 밝은 미래와 지속가능한 발전"
        },
        {
            "focus": "균형적 접근",
            "tone": "안정적이고 신뢰감 있는 어조",
            "emphasis": "경제, 화합, 희망이 균형잡힌 종합적 관점"
        },
        {
            "focus": "혁신과 변화 중심",
            "tone": "혁신적이고 도전적인 어조",
            "emphasis": "새로운 변화와 혁신을 통한 울진군의 도약"
        }
    ]

    print(f"\n=== 5개의 발간사 버전을 생성합니다 ===")
    
    generated_forewords = []
    
    for i, variation in enumerate(foreword_variations, 1):
        print(f"\n--- {i}번째 발간사 생성 중 ({variation['focus']}) ---")
        
        # 프롬프트 구성
        first_prompt_parts = []
        
        if uploaded_files:
            first_prompt_parts.extend(uploaded_files)
        
        if combined_pdf_text:
            first_prompt_parts.append(f"다음은 추출된 텍스트 내용입니다:\n{combined_pdf_text}")
        
        prompt_text = f"""
            업로드된 PDF 파일들을 모두 종합적으로 참고해서 다음 요구사항에 맞는 '{year}년도 {quarter}분기 발간사'를 작성해줘:

            **이번 버전의 특별 요구사항:**
            - 주요 초점: {variation['focus']}
            - 어조: {variation['tone']}
            - 강조점: {variation['emphasis']}

            **공통 요구사항:**
            1. '경제', '화합', '희망'을 핵심 주제로 포함해서 작성해줘.
            2. 경제 부분에는 울진군의 지속가능한 성장 동력을 완성해나가는 과정이 잘 드러나도록 작성해줘.
            3. 군민 화합을 강조하고 미래에 대한 희망적인 메시지로 마무리해줘.
            4. 공식적이고 품격 있는 어조로 작성하되, 적절한 분량(6-10줄)으로 작성해줘.
            5. 다른 버전들과는 차별화된 독특한 관점과 표현을 사용해줘.

            발간사 내용만 작성해주고, 다른 설명은 생략해줘.
        """
        
        first_prompt_parts.append(prompt_text)
        
        try:
            response = model.generate_content(
                first_prompt_parts,
                generation_config=genai.types.GenerationConfig(
                    temperature=0.8,  # 창의성을 위해 약간 높게 설정
                    top_p=0.9,
                    top_k=40,
                    max_output_tokens=100000,
                ),
            )
            
            if response.parts and hasattr(response, 'text'):
                foreword_text = response.text.strip()
                generated_forewords.append(foreword_text)
                
                print(f"✅ {i}번째 발간사 생성 완료")
                print(f"미리보기: {foreword_text[:100]}...")
                
                # 한글 파일 생성
                output_filename = f"발간사_{i}_{variation['focus'].replace(' ', '_')}.hwp"
                output_path = os.path.join(output_dir, output_filename)
                
                success, message = create_hwp_document_with_foreword(
                    hwp_path, foreword_text, output_path, i
                )
                
                if success:
                    print(f"✅ 한글 파일 저장 성공: {output_filename}")
                else:
                    print(f"❌ 한글 파일 저장 실패: {message}")
                
            else:
                print(f"❌ {i}번째 발간사 생성 실패")
                
        except Exception as e:
            print(f"❌ {i}번째 발간사 생성 중 오류: {e}")
        
        # 요청 간 잠시 대기 (API 제한 방지)
        if i < len(foreword_variations):
            print("   잠시 대기 중...")
            time.sleep(3)
    
    # --- 5. 결과 요약 출력 ---
    
    for i, foreword in enumerate(generated_forewords, 1):
        print(f"\n--- {i}번째 발간사 ({foreword_variations[i-1]['focus']}) ---")
        print(foreword)
        print("-" * 50)

except Exception as e:
    print(f"\n프로그램 실행 중 오류 발생: {e}")
    
finally:
    # --- 안전한 종료 ---
    print("\n프로그램을 안전하게 종료합니다...")
    
    # 업로드된 파일들 정리
    try:
        if uploaded_files:
            print("업로드된 임시 파일들을 정리합니다...")
            for file in uploaded_files:
                try:
                    genai.delete_file(file.name)
                    print(f"임시 파일 삭제: {file.display_name}")
                except Exception as delete_error:
                    print(f"파일 삭제 실패 {file.display_name}: {delete_error}")
    except Exception as cleanup_error:
        print(f"파일 정리 중 오류: {cleanup_error}")
    
    print("프로그램 종료 완료")
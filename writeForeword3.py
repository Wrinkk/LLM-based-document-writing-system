import pyhwpx
import google.generativeai as genai
import os
import re
from google.generativeai.types import HarmCategory, HarmBlockThreshold
import datetime
import time
import PyPDF2
from concurrent.futures import ThreadPoolExecutor, as_completed
import string
from dotenv import load_dotenv

def combine_pdf_texts(pdf_files_paths):
    """여러 PDF 파일의 텍스트를 효율적으로 합치는 함수 (추출과 정제를 한 번에 처리)"""

    combined_text = ""
    
    for file_path in pdf_files_paths:
        if not os.path.exists(file_path):
            print(f"  - 경고: '{os.path.basename(file_path)}' 파일을 찾을 수 없어 건너뜁니다.")
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
                    print(f"  - 경고: '{os.path.basename(file_path)}'에서 텍스트를 추출할 수 없습니다.")
                    continue

                # 정제 과정 개선: 필요한 특수문자 유지
                cleaned_text = re.sub(r'[^\w\s가-힣.,!?;:\'"""''·()\[\]{} -]', ' ', raw_text)
                cleaned_text = re.sub(r'\s+', ' ', cleaned_text)  # 중복 공백 제거
                
                combined_text += f"\n\n=== {os.path.basename(file_path)}의 내용 ===\n"
                combined_text += cleaned_text
                print(f"  - 완료: {len(cleaned_text):,} 문자")

        except Exception as e:
            print(f"  - 오류: '{os.path.basename(file_path)}' 처리 중 오류 발생: {e}")
            
    print(f"\n통합 및 정제된 텍스트 총 길이: {len(combined_text):,} 문자")
    return combined_text

def wait_for_file_processing(uploaded_files, max_wait_time=300):
    """업로드된 파일들의 처리 완료를 대기하는 함수"""

    for file in uploaded_files:
        start_time = time.time()
        while file.state.name == "PROCESSING":
            elapsed = time.time() - start_time
            if elapsed > max_wait_time:
                print(f"  - {file.display_name} 처리 시간 초과")
                break
            
            print(f"  - {file.display_name} 처리 중... 10초 대기 (경과: {elapsed:.0f}초)")
            time.sleep(10)
            file = genai.get_file(file.name)
        
        if file.state.name == "FAILED":
            print(f"  - 경고: {file.display_name} 처리 실패")
            continue
        elif file.state.name == "ACTIVE":
            print(f"  - {file.display_name} 처리 완료")

def upload_file_concurrently(file_path):
        """단일 파일을 Gemini API에 업로드하고 결과를 반환하는 함수 (스레드에서 실행)"""

        if not os.path.exists(file_path):
            print(f"  - 경고: '{os.path.basename(file_path)}' 파일을 찾을 수 없어 건너뜁니다.")
            return None

        file_size = os.path.getsize(file_path)
        if file_size > 200 * 1024 * 1024:  # 200MB
            print(f"  - 경고: '{os.path.basename(file_path)}' 파일 크기가 200MB를 초과하여 건너뜁니다.")
            return None

        print(f"  - 업로드 시작: {os.path.basename(file_path)}")
        try:
            file_response = genai.upload_file(path=file_path)
            print(f"  - 업로드 완료: {file_response.display_name}")
            return file_response
        except Exception as e:
            print(f"  - 업로드 실패: {os.path.basename(file_path)} - {e}")
            return None

def create_hwp_document_with_foreword(template_path, foreword_text, output_path, version_num):
    """발간사가 포함된 한글 문서를 생성하는 함수"""

    hwp = None
    try:
        hwp = pyhwpx.Hwp(visible=True)
        
        # 템플릿 파일 열기
        if os.path.exists(template_path):
            hwp.Open(template_path)
            print(f"템플릿 파일 열기: {os.path.basename(template_path)}")
        else:
            # 새 문서 생성
            hwp.XHwpDocuments.Add()
            print("새 문서 생성")
        
        # 문서 시작으로 이동
        hwp.MoveDocBegin()
        
        
        sections = re.split(r'##\s*([A-Z][0-9])', foreword_text)
        print(f"{version_num}번째 질문 파싱 결과:", sections)
        
        content_map = {}
        if len(sections) > 1:
            for i in range(1, len(sections), 2):
                marker = sections[i].strip()
                full_content = sections[i+1].strip()
                print(full_content)
                content_map[marker] = full_content

        print(f"\nHWPX 파일에 {version_num}번째 질문 파싱된 답변을 삽입합니다...")
        for marker, full_content in content_map.items():
            # 전체 내용에서 첫 번째 줄(제목)만 추출
            title_only = full_content.split('\n')[0].strip()

            hwp.MoveDocBegin() 
            while hwp.find(marker,direction='AllDoc'):

                hwp.insert_text(title_only)
                time.sleep(0.1)
                print(f"  - 성공: '{marker}' 위치에 제목 '{title_only}'을(를) 삽입했습니다.")

        hwp.MoveDocBegin() # 문서 시작으로 이동
        
        # 한/글의 '찾아 바꾸기' 기능을 정규식 모드로 실행
        possible_markers = [f"{char}{num}" for char in string.ascii_uppercase for num in range(1, 10)]
        

        for marker in possible_markers:
            while hwp.find(marker, direction='AllDoc'):
                hwp.Erase()
                hwp.DeleteLine()
                hwp.DeleteLine()

        while hwp.find("#",direction='AllDoc'):
            hwp.Erase()


        # 파일 저장
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

def main():
    """메인 실행 함수"""
    # 전역 변수
    uploaded_files = []

    load_dotenv() 
    
    try:
        # ===== 사용자 설정 섹션 =====
        # [사용자 설정 1] 템플릿 파일들 경로 설정
        template_base_dir = r'C:\Users\USER\Desktop\llm\LLM-based-document-writing-system\template'
        template_paths = [
            os.path.join(template_base_dir, f"template1.hwpx")
        ]
        
              # [사용자 설정 3] API 실패 시 사용할 5가지 기본 발간사 내용 (여기를 채워주세요)
        fallback_forewords = [
            # 1번째 발간사 (경제 발전 중심) 실패 시 사용될 내용
            """##AA 존경하는 울진 군민 여러분, 2025년 한 해를 마무리하며 새로운 희망을 품는 4분기를 맞이했습니다. 
               ##BB 올해 우리는 미래 경제 성장을 위한 기반을 다지며 괄목할  만한 성과를 이루어냈습니다.
               ##CC 원자력수소 국가산업단지 조성과 같은 미래 성장 동력 확보에 매진하며, 지역 경제에 활력을 불어넣었습니다. 
               ##DD 또한, 관광 인프라를 전면 개편하고 농업 대전환을 통해 지역 특산물의 가치를 높였습니다. 
               ##EE 전통시장 활성화와 소상공인 지원으로 지역 경제의 든든한 화합을 이끌어냈습니다. 
               ##FF 이 모든 성과는 군민 여러분의 적극적인 참여와 노력이 있었기에 가능했습니다. 
               ##GG 역동적인 변화와 진취적인 도전으로, 울진은 더 큰 희망을 향해 나아가고 있습니다. 
               ##HH  앞으로도 지속적인 경제 발전 전략을 통해 군민 모두가 풍요로운 미래를 누리도록 최선을 다하겠습니다. 
               ##II 늘 함께해주시는 군민 여러분께 진심으로 감사드립니다. 
               ##JJ 다가오는 새해에도 울진의 밝은 미래를 함께 만들어 주시기를 바랍니다.""",

        ]



        API_KEY = os.getenv('GENAI_API_KEY') 
        
        if not API_KEY:
            raise ValueError("API 키를 입력해주세요. API_KEY 변수에 실제 키를 설정하세요.")
        
        # 템플릿 파일들 존재 확인
        available_templates = []
        for i, template_path in enumerate(template_paths, 1):
            if os.path.exists(template_path):
                available_templates.append(template_path)
            else:
                print(f"  ✗ template{i}.hwp 파일 없음: {template_path}")
        
        if not available_templates:
            print("경고: 사용 가능한 템플릿이 없습니다. 새 문서로 생성합니다.")
        
        # 결과 파일들을 저장할 디렉토리 생성
        output_dir = os.path.join(template_base_dir, f"발간사_결과_{datetime.datetime.now().strftime('%M%S')}")
        os.makedirs(output_dir, exist_ok=True)
        print(f"\n결과 저장 경로: {output_dir}")
        
        genai.configure(api_key=API_KEY)

        # ===== PDF 파일 경로 설정 =====
        # 텍스트를 추출할 PDF 파일 목록
        text_extract_paths = [
            # r"C:\Users\USER\Downloads\hwp\압축\희망울진 군정집(2025년 1분기).pdf",
            # r"C:\Users\USER\Downloads\hwp\압축\희망울진 군정집(2025년 2분기).pdf",
            # r"C:\Users\USER\Downloads\hwp\압축\희망울진 군정집(2025년 3분기).pdf",
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
            r"C:\Users\USER\Downloads\hwp\압축\2026년 주요업무보고(환동해산업연구원).pdf",
            r"C:\Users\USER\Downloads\hwp\압축\2023년도 시정연설문.pdf",
            r"C:\Users\USER\Downloads\hwp\압축\2025년 시정연설.pdf",
            r"C:\Users\USER\Downloads\hwp\압축\2024년도 시정연설문(최종).pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2022년 10월 정례 조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2022년 12월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2022년 9월 월례 조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2023년 10월 정례조회 월례사 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2023년 10월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2023년 11월 정례조회 월례사 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2023년 11월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2023년 2월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2023년 3월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2023년 4월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2023년 5월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2023년 6월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2023년 8월 정례조회 월례사 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2023년 8월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2023년 9월 정례조회 월례사 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2023년 9월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2023년 송년사(최종).pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2024년 10월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2024년 11월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2024년 12월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2024년 2월 정례조회 월례사 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2024년 2월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2024년 3월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2024년 4월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2024년 5월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2024년 6월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2024년 8월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2024년 9월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2024년 송년사.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2024년 신년사(최종).pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2025년 3월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2025년 4월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2025년 5월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2025년 6월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2025년 7월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2025년 8월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2025년 9월 정례조회 훈시 1부.pdf",
            r"C:\Users\USER\Downloads\hwp\민선8기 분량정리\pdf\2025년 신년사.pdf",


        ]

        # ===== 1. 텍스트 추출 그룹 처리 =====
        combined_pdf_text = ""
        if text_extract_paths:
            combined_pdf_text = combine_pdf_texts(text_extract_paths)

        # ===== 2. 파일 업로드 그룹 처리 (병렬 처리 적용) =====
        if file_upload_paths:

            with ThreadPoolExecutor(max_workers=27) as executor: 
                future_to_file = {executor.submit(upload_file_concurrently, file_path): file_path 
                                  for file_path in file_upload_paths} 
                
                for future in as_completed(future_to_file):
                    file_path = future_to_file[future]
                    try:
                        file_response = future.result() # 업로드 결과 가져오기
                        if file_response:
                            uploaded_files.append(file_response)
                    except Exception as exc:
                        print(f"  - '{os.path.basename(file_path)}' 업로드 중 예외 발생: {exc}")

            if uploaded_files:
                wait_for_file_processing(uploaded_files)
                print(f"\n총 {len(uploaded_files)}개 파일 업로드 완료")

        # ===== 3. 모델 초기화 =====
        try:
            model = genai.GenerativeModel(
                'gemini-2.5-flash',  # 최신 모델 사용
                safety_settings={
                    HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
                    HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
                    HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
                    HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
                }
            )
        except Exception as e:
            print(f"기본 모델 초기화 실패: {e}")
            model = genai.GenerativeModel('gemini-1.5-flash')
            print("대체 모델 사용: gemini-1.5-flash")

        year = datetime.datetime.now().year + 1
        month = datetime.datetime.now().month
        quarter = ((datetime.datetime.now().month - 1) // 3 + 1) + 1

        foreword_variations = [
            {
                "focus": "경제 발전 중심",
                "tone": "역동적이고 진취적인 어조",
                "emphasis": "울진군의 미래 성장 동력과 경제 발전 전략"
            }
        ]
        
        generated_forewords = []
        
        # 토큰 절약을 위해 텍스트 데이터 축약
        summarized_text = ""
        if combined_pdf_text:
            summarized_text = combined_pdf_text[:500000]  
            if len(combined_pdf_text) > 500000:
                summarized_text += "\n...(추가 내용 생략)..."
        
        for i, variation in enumerate(foreword_variations, 1):
            
            max_retries = 3
            retry_count = 0
            success = False
            
            while retry_count < max_retries and not success:
                try:
                    # 프롬프트 구성
                    prompt_parts = []
                    
                    if i == 1 and retry_count == 0:
                        if uploaded_files:
                            prompt_parts.extend(uploaded_files)
                        
                        if summarized_text:
                            prompt_parts.append(f"[참고 자료 요약]\n{summarized_text}")
                    
                    prompt_text = f"""
                                    {year}년도 울진군 시장연설문을 작성해주세요. 60갑자 제대로 확인해서 적어주세요.

                                    당신은 '화합으로 새로운 희망울진'을 군정 비전으로 삼고 있는 손병복 울진군수입니다. 군민과 군의회를 존중하며, 울진의 미래에 대한 확신과 비전을 담아 연설문을 작성해야 합니다.

                                    * 작성 지침
                                    2025년에 열리는 울진군의회 제2차 정례회에서 '2026년도 예산안'을 제출하며 발표할 시정연설문을 작성해 주세요.
                                    대상: 울진 군민과 군의회 의원
                                    목적: 2025년의 주요 군정 성과를 보고하고, 이를 바탕으로 수립된 2026년도 군정 운영 방향과 핵심 사업들을 설명하여 예산안에 대한 이해와 협조를 구하는 것입니다.

                                    2. 필수 포함 내용
                                       - 경제, 화합, 희망 중심의 내용
                                       - 울진군의 구체적 성과와 비전 제시
                                       - 군민에 대한 감사와 격려
                                        '2026년 주요 업무계획' 보고서를 핵심 자료로 활용
                                        '2025년 주요성과'은 지난 성과와 성과에 대한 상세 내용을 설명
                                        '2026년 주요 업무 추진계획' 부분은 내년도 계획을 설명
                                        - 공백포함 15,000자 이상 20,000자 이내으로 작성

                                    3. 형식 요구사항
                                       - 공식적이고 품격있는 문체
                                       - 주요사업 내용에 대해 구체적 설명하고, 어떻게 진행했는지 심도있는 내용으로 작성
                                       - 강조할때는 '**' 문자 말고 '##'로 강조    
                                       - 각 문단마다 파싱문자 ##A1~A9,##B1~B9,##C1~C9,##D1~D9, ... 부여 (하나의 알파벳에 대해 숫자는 1부터 9까지 순서대로, 빠짐없이 모두 사용)
                                       - 연설문의 전체적인 구조, 문체, 어조는 함께 첨부된 2023년, 2024년, 2025년 시정연설문을 참고하여 일관성을 유지
                                       - 어조 및 형식(Tone & Format)
                                          어조: 군민과 군의회를 존중하는 정중하고 진솔한 어조를 사용하되, 미래 비전에 대해서는 자신감 있고 희망적인 어조를 사용해 주세요.
                                          문체: 격식을 갖춘 문어체로 작성해 주세요.
                                          사자성어: 연설의 마지막 부분에 군정 운영의 의지를 나타낼 수 있는 적절한 사자성어를 포함해 주세요.

                                    """
                    
                    prompt_parts.append(prompt_text)
                    
                    response = model.generate_content(
                        prompt_parts,
                        generation_config=genai.types.GenerationConfig(
                            temperature=0.8,
                            top_p=0.9,
                            top_k=40,
                            max_output_tokens=70000,
                        ),
                    )
                    
                    if response and hasattr(response, 'text'):
                        foreword_text = response.text.strip()
                        
                        # 응답 검증
                        if len(foreword_text) < 5000:
                            print(f"  - 경고: 생성된 텍스트가 너무 짧습니다. 재시도...")
                            retry_count += 1
                            continue
                        
                        generated_forewords.append(foreword_text)
                        
                        print(f"  ✓ 발간사 생성 완료 ({len(foreword_text)}자)")
                        print(f"  미리보기: {foreword_text[:80]}...")
                        
                        # 한글 파일 생성
                        template_to_use = available_templates[i-1] if i <= len(available_templates) else (available_templates[0] if available_templates else "")
                        output_filename = f"{year}년_시장연설문.hwp"
                        output_path = os.path.join(output_dir, output_filename)
                        
                        file_success, message = create_hwp_document_with_foreword(
                            template_to_use, foreword_text, output_path, i
                        )
                        
                        if file_success:
                            print(f"  ✓ 한글 파일 생성: {output_filename}")
                        else:
                            print(f"  ✗ 한글 파일 생성 실패: {message}")
                        
                        success = True
                        
                    else:
                        print(f"  ✗ 응답 없음. 재시도 {retry_count+1}/{max_retries}")
                        retry_count += 1
                        
                except Exception as e:
                    error_message = str(e)
                    print(f"  ✗ 오류 발생: {error_message[:100]}")
                    
                    # API 할당량 초과 처리
                    if "429" in error_message or "quota" in error_message.lower():
                        retry_match = re.search(r'retry in (\d+\.?\d*)', error_message)
                        wait_time = float(retry_match.group(1)) + 5 if retry_match else 30
                        print(f"  ⏳ API 한도 초과. {wait_time:.0f}초 대기...")
                        time.sleep(wait_time)
                    else:
                        time.sleep(5)
                    
                    retry_count += 1
            
            if not success:
                
                # 실패 시, 해당 순서(i)에 맞는 기본 발간사 사용
                template_foreword = fallback_forewords[i-1] # i는 1부터 시작하므로 인덱스는 i-1
                generated_forewords.append(template_foreword)
                
                template_to_use = template_paths[i-1] if i <= len(template_paths) else ""
                output_filename = f"발간사_{i:02d}_{variation['focus'].replace(' ', '_')}_기본.hwp"
                output_path = os.path.join(output_dir, output_filename)
                
                file_success, message = create_hwp_document_with_foreword(template_to_use, template_foreword, output_path, i)
                if file_success:
                    print(f"  ✓ 기본 템플릿 파일 생성: {output_filename}")
                
                generated_forewords.append(template_foreword)
                
                # 기본 템플릿으로 파일 생성
                template_to_use = available_templates[0] if available_templates else ""
                output_filename = f"발간사_{i:02d}_{variation['focus'].replace(' ', '_')}_기본.hwp"
                output_path = os.path.join(output_dir, output_filename)
                
                file_success, message = create_hwp_document_with_foreword(
                    template_to_use, template_foreword, output_path, i
                )
                
                if file_success:
                    print(f"  ✓ 기본 템플릿 파일 생성: {output_filename}")
            
            # 다음 요청 전 대기
            if i < len(foreword_variations):
                print("  다음 발간사 생성을 위해 대기 중...")
                time.sleep(5)
        
        # ===== 5. 결과 출력 =====
        print(f"\n{'=' * 60}")
        print("생성된 발간사 전체 내용")
        print(f"{'=' * 60}")
        
        for i, foreword in enumerate(generated_forewords, 1):
            print("-" * 40)
            print(foreword)
            print("-" * 40)
        
        print(f"결과 파일 위치: {output_dir}")
        
    except Exception as e:
        print(f"\n❌ 프로그램 실행 중 치명적 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        # ===== 정리 작업 =====
        print("\n프로그램 정리 작업 시작...")
        
        # 업로드된 파일들 정리
        if uploaded_files:
            print("업로드된 임시 파일들을 삭제합니다...")
            for file in uploaded_files:
                try:
                    genai.delete_file(file.name)
                    print(f"  ✓ 삭제: {file.display_name}")
                except Exception as e:
                    print(f"  ✗ 삭제 실패: {file.display_name} - {e}")

        print("\n프로그램 종료")
        print("=" * 60)

if __name__ == "__main__":
    main()
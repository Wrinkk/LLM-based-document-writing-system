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
        hwp = pyhwpx.Hwp(visible=False)
        
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
        
        
        sections = re.split(r'##\s*([A-Z]{2,})', foreword_text)
        # print(f"{version_num}번째 질문 파싱 결과:", sections)
        
        content_map = {}
        if len(sections) > 1:
            for i in range(1, len(sections), 2):
                marker = sections[i].strip()
                full_content = sections[i+1].strip()
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

                # 문서에 남아있을 수 있는 미사용 파싱 마커 (예: ##CC, ##DD 등)를 모두 찾아 삭제합니다.
        print("\n    최종 문서에서 잔여 파싱 마커를 제거합니다...")
        hwp.MoveDocBegin() # 문서 시작으로 이동
        
        # 한/글의 '찾아 바꾸기' 기능을 정규식 모드로 실행
        possible_markers = [f"{char}{char}" for char in string.ascii_uppercase]
        for marker in possible_markers:
            # hwp.Find는 찾으면 True, 못 찾으면 False를 반환합니다.
            # 문서 전체를 계속 반복하며 해당 마커가 더 이상 없을 때까지 찾아서 지웁니다.
            while hwp.find(marker, direction='AllDoc'):
                # 찾은 마커를 빈 문자열로 대체 (삭제)
                hwp.Delete()
                print("삭제완료")
        
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
        template_base_dir = r'C:\Users\wj830\Desktop\dd'
        template_paths = [
            os.path.join(template_base_dir, f"template{i}.hwpx") for i in range(1, 6)
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
            
            # 2번째 발간사 (군민 화합 중심) 실패 시 사용될 내용
            """##AA 존경하는 울진군민 여러분, 2025년 한 해를 마무리하며 희망찬 새해를 준비하는 4분기 군정집을 발간합니다.
               ##BB 올 한 해 울진은 군민 여러분의 지혜와 역량이 한데 모여 더욱 굳건한 공동체로 성장했습니다.
               ##CC 어려운 여건 속에서도 흔들림 없는 경제 기반을 다지고자 끊임없이 노력하여 미래 성장 동력을 확보했습니다.
               ##DD 이 모든 과정에서 서로를 보듬고 아끼는 화합의 정신은 우리 군이 직면한 여러 과제를 슬기롭게 극복하는 원동력이 되었습니다.
               ##EE 군민과 함께 소통하며 만들어 온 정책들은 울진의 삶의 질을 높이고 내일의 희망을 키워냈습니다.
               ##FF 저희는 앞으로도 군민의 삶을 최우선에 두고, 상생의 가치를 실현하는 따뜻하고 포용적인 군정을 펼쳐나가겠습니다.
               ##GG 2026년에도 군민 여러분과 손잡고 더 큰 울진의 발전과 도약을 위한 비전을 향해 힘껏 나아가겠습니다.
               ##HH 모든 세대가 함께 웃고, 미래를 꿈꿀 수 있는 행복한 울진을 만들기 위한 여정에 변함없는 동참을 부탁드립니다.
               ##II 군민 여러분의 깊은 관심과 따뜻한 격려에 진심으로 감사드리며, 가정에 늘 건강과 행복이 가득하시기를 기원합니다.
               ##JJ 감사합니다.""",
                           
            # 3번째 발간사 (희망과 비전 중심) 실패 시 사용될 내용
            """##AA 존경하는 울진군민 여러분, 2025년 한 해를 마무리하는 제4분기 군정집을 발간하게 되어 매우 뜻깊게 생각합니다.
               ##BB 올 한 해 군민 여러분과 함께 일궈온 값진 노 력들이 풍성한 결실을 맺으며 희망찬 울진의 미래를 밝히고 있습니다.
               ##CC 특히, 지속가능한 경제 발전을 위한 새로운 성장 동력 확보는 물론, 지역 경제에 활력을 불어넣기 위한 노력을 멈추지 않았습니다.
               ##DD 이 모든 과정 속에서 군민들의 하나 된 화합과 참여는 울진 발전의 가장 든든한 초석이 되었습니다. 
               ##EE 우리는 더 나은 내일을 향한 분명한 비전을 가지고, 모두가 살기 좋은 울진을 만들기 위한 담대한 도전을 이어가고 있습니다.
               ##FF 청정 자연을 기반으로 한 친환경 성장과 혁신적인 정책을 통해 울진의 지속가능한 발전 모델을 정립해 나갈 것입니다.
               ##GG 군민 한 분 한 분의 삶에 따뜻한 희망이 스며들고, 모두가 행복을 누리는 풍요로운 울진을 향해 나아가겠습니다. 
               ##HH 지난 한 해 동안 변함없는 사랑과 성원을 보내주신 군민 여러분께 진심으로 감사드립니다. 
               ##II 다가오는 새해에도 군민 여러분과 함께 더 큰 도약을 이루어낼 수 있도록 최선을 다하겠습니다. 
               ##JJ 군민 여러분의 가정에 늘 건강과 행복이 가득하시기를 기원합니다.""",
               
            # 4번째 발간사 (균형적 접근) 실패 시 사용될 내용
            """##AA 존경하는 울진군민 여러분, 2025년 한 해를 마무리하는 4분기 군정집을 발간하며 인사드립니다. 
               ##BB 올 한 해도 변함없이 군정에 깊은 관심과 성원을 보내주신 군민 여러분께 진심으로 감사드립니다. 
               ##CC 견고한 지역 경제 기반을 다지기 위한 노력은 가시적인 성과로 이어져, 지속 가능한 성장의 발판을 마련하였습니다.
               ##DD 군민의 화합된 힘은 크고 작은 어려움을 슬기롭게 극복하고, 서로를 배려하는 공동체 정신을 더욱 굳건히 하는 원동력이 되었습니다. 
               ##EE 우리는 미래 세대가 더 나은 삶을 꿈꿀 수 있는 희망찬 울진을 만들기 위해 끊임없이 고민하고 실천해왔습니다. 
               ##FF 해양 관광 활성화부터 스마트 농업 혁신, 그리고 품격 있는 문화생활에 이르기까지, 균형 잡힌 발전을 위해 다각적인 노력을 기울였습니다. 
               ##GG 다가오는 새해에는 더 큰 도약과 변화를 향해 나아가며, 군민 한 분 한 분의 삶이 더욱 풍요로워지도록 최선을 다하겠습니다.
               ##HH 울진의 밝은 미래를 향한 여정에 늘 함께해주시는 군민 여러분의 지혜와 참여에 깊이 감사드립니다. 
               ##II 앞으로도 변함없는 사랑과 성원을 부탁드리며, 가정에 늘 건강과 행복이 가득하시기를 기원합니다.""",
               
            # 5번째 발간사 (혁신과 변화 중심) 실패 시 사용될 내용
            """##AA 존경하는 울진군민 여러분, 2025년 한 해를 마무리하며 새 시대를 향한 도약의 의지를 담은 군정집을 선보입니다.
               ##BB 올 한 해 울진군은 혁신과 변화를 기치로 내걸고, 미래를 향한 과감한 도전을 멈추지 않았습니다.
               ##CC 이러한 노력들이 결실을 맺어 지역 경제에 활력을 불어넣고, 새로운 성장 동력을 확보하는 데 주력했습니다.
               ##DD 특히, 군민 모두가 한마음으로 뜻을 모아 이룬 화합의 가치는 어떠한 난관도 극복할 수 있는 굳건한 울진의 힘이 되었습니다.
               ##EE 다가올 새해에는 스마트 도시 기반 구축과 미래 신산업 육성을 통해 지속 가능한 발전의 희망찬 비전을 현실로 만들어갈 것입니다.
               ##FF 우리는 변화를 두려워하지 않고, 더 나은 울진의 내일을 위한 혁신적 시도를 계속해 나갈 것입니다.
               ##GG 군민 여러분의 적극적인 참여와 성원이 있었기에 오늘의 성과가 가능했으며, 이는 미래를 향한 가장 강력한 추진력입니다.
               ##HH 앞으로도 군민의 삶의 질 향상과 울진의 위상 강화를 위해 모든 역량을 집중할 것을 약속드립니다.
               ##II 뜨거운 열정으로 함께해 주신 군민 여러분께 진심으로 감사드리며, 새해에도 변함없는 지지와 격려를 부탁드립니다.
               ##JJ 희망찬 미래를 향한 울진의 여정에 동참해 주시길 바랍니다.""",
        ]


        # [사용자 설정 2] 본인의 Gemini API 키 입력
        API_KEY = os.getenv("GEMINI_API_KEY")  # 여기에 실제 API 키를 입력하세요
        print(f"현재 작업 폴더: {os.getcwd()}")
        print(f".env 파일 존재 여부: {os.path.exists('.env')}")
        
        if not API_KEY:
            raise ValueError("API 키를 입력해주세요. API_KEY 변수에 실제 키를 설정하세요.")
        
        # ===== 초기 설정 및 검증 =====
        print("=" * 60)
        print("울진군 발간사 생성 프로그램 시작")
        print("=" * 60)
        
        # 템플릿 파일들 존재 확인
        print("\n템플릿 파일들을 확인합니다...")
        available_templates = []
        for i, template_path in enumerate(template_paths, 1):
            if os.path.exists(template_path):
                print(f"  ✓ template{i}.hwp 파일 존재")
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
            r"C:\Users\wj830\Desktop\dd\llm_data\희망울진 군정집(2025년 1분기).pdf",
            # r"C:\Users\wj830\Desktop\dd\llm_data\희망울진 군정집(2025년 2분기).pdf",
            # r"C:\Users\wj830\Desktop\dd\llm_data\희망울진 군정집(2025년 3분기).pdf",
        ]

        # 파일 채로 AI에 업로드할 PDF 파일 목록
        file_upload_paths = [
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(건설과).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(경제교통과).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(기획예산실)(9.8. 수정).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(농기계임대사업소).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(농업기술센터)(수정).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(농정과).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(도시새마을과).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(맑은물사업소).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(문화관광과).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(민원과)(9.8. 수정).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(보건소).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(복지정책과).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(사회복지과).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(산림과).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(수소국가산업추진단).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(안전재난과).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(왕피천공원사업소).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(울진군의료원).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(원전에너지과).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(인구정책과).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(재무과)(9.8. 수정).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(정책홍보실)(9.11. 수정).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(체육진흥과).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(총무과)(9.8. 수정).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(해양수산과).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(환경위생과).pdf",
            r"C:\Users\wj830\Desktop\dd\llm_data\2026년 주요업무보고(환동해산업연구원).pdf"
        ]

        # ===== 1. 텍스트 추출 그룹 처리 =====
        combined_pdf_text = ""
        if text_extract_paths:
            combined_pdf_text = combine_pdf_texts(text_extract_paths)

        # ===== 2. 파일 업로드 그룹 처리 (병렬 처리 적용) =====
        if file_upload_paths:
            print("\nPDF 파일을 AI에 직접 업로드합니다 (병렬 처리)...")

            with ThreadPoolExecutor(max_workers=27) as executor: # 예시: 5개 파일 동시 업로드
                # 각 파일에 대해 업로드 작업을 제출
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

        year = datetime.datetime.now().year
        month = datetime.datetime.now().month
        quarter = ((datetime.datetime.now().month - 1) // 3 + 1) + 1

        # ===== 4. 5개의 다른 발간사 생성 =====
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
        
        generated_forewords = []
        
        # 토큰 절약을 위해 텍스트 데이터 축약
        summarized_text = ""
        if combined_pdf_text:
            summarized_text = combined_pdf_text[:500000]  # 15,000자로 제한
            if len(combined_pdf_text) > 500000:
                summarized_text += "\n...(추가 내용 생략)..."
        
        for i, variation in enumerate(foreword_variations, 1):
            print(f"\n[{i}/5] {variation['focus']} 발간사 생성 중...")
            
            max_retries = 3
            retry_count = 0
            success = False
            
            while retry_count < max_retries and not success:
                try:
                    # 프롬프트 구성
                    prompt_parts = []
                    
                    # 첫 번째 요청에만 파일들과 텍스트 포함
                    # 파일은 한 번 업로드하면 모델이 기억하므로 매번 보낼 필요 없습니다.
                    # 텍스트 요약본은 토큰 사용량이 크므로 첫 요청에만 포함.
                    if i == 1 and retry_count == 0:
                        if uploaded_files:
                            prompt_parts.extend(uploaded_files)
                        
                        if summarized_text:
                            prompt_parts.append(f"[참고 자료 요약]\n{summarized_text}")
                    
                    prompt_text = f"""
                                    {year}년도 {quarter}분기 울진군 군정집 발간사를 작성해주세요.

                                    **작성 지침:**
                                    1. 버전 특징
                                       - 초점: {variation['focus']}
                                       - 어조: {variation['tone']}
                                       - 강조점: {variation['emphasis']}

                                    2. 필수 포함 내용
                                       - 경제, 화합, 희망 3대 키워드 자연스럽게 포함 
                                       - 울진군의 구체적 성과와 비전 제시
                                       - 군민에 대한 감사와 격려
                                       - {year}년 {quarter}분기 시의성 반영
                                       

                                    3. 형식 요구사항
                                       - 10줄 분량 
                                       - 공식적이고 품격있는 문체
                                       - 구체적이고 실질적인 내용
                                       - 인사말로 시작, 감사/격려로 마무리
                                       - 각 문장마다 파싱문자 ##AA,##BB,##CC,##DD,##EE,##FF,##GG,##HH,##II,##JJ,##KK,##LL, .. 부여
                                       - 예시 ##AA 내용 ##BB 내용 ##CC 내용

                                    4. 이전 분기 발간사 참고내용
                                    
                                    * 2분기 내용
                                   - 아이를 낳고 키우기 좋은 도시는
                                     삶의 기준을 바꾸는 우리의 소중한 선택입니다.
                                     
                                     울진은 지금 그런 아름다운 도시가 되기 위해
                                     변화의 길을 힘차게 걷고 있습니다.
                                     
                                     아이들이 마음껏 뛰놀며 꿈을 키울 수 있는 자연과 환경,
                                     다자녀 유공 수당 지급 등으로 아이 키우는 부담 경감,
                                     부모가 안심하고 맡길 수 있는 다정한 돌봄,
                                     이웃이 함께 아이를 키우는 믿음직한 공동체 울진.
                                     
                                     그리고 오랜 세월 울진을 지켜오신
                                     어르신들이 존중받고, 건강한 노후를 보내실 수 있는 도시.
                                     목욕비, 이·미용비 지원, 경로당 식사 제공, 어르신 일자리 확대 등
                                     체감할 수 있는 맞춤형 복지를 통해
                                     어르신들의 일상에도 따뜻한 변화를 만들어가고 있습니다.
                                     
                                     아이와 어르신이 함께 웃고,
                                     모든 세대가 함께 어우러지는 복지공동체 울진.
                                     이곳에서 아이들의 꿈이 자라고, 울진의 내일도 함께 자라납니다.
                                     
                                     우리 아이들과 어르신 모두에게 더 나은 울진을
                                     물려주기 위한 노력은 오늘도, 앞으로도 계속될 것입니다.
                                     
                                     아름다운 내일을 꿈꾸며 나아가는
                                     이 길에 군민 여러분의 동참과 성원을 바랍니다.
                                     감사합니다.
                                     
                                     2025년 5월 울진군수 손병복

                                    * 3분기 내용
                                    군민들이 스스로 울진군의 주인이라고 생각하고,
                                    공직자들이 군정 운영의 주인이라고 생각하며
                                    책임감을 가지고 자신의 자리에서 최선을 다해야 할 것입니다.
                                    
                                    모두가 주인의식을 가질 수 있도록 하기 위해 군민들을 섬기는 군정,
                                    공직자들이 자긍심을 가지고 일할 수 있는
                                    환경 마련에 최선을 다하고 있습니다.

                                    * 발간사 본문만 작성하고 다른 설명은 포함하지마.
                                    * 그리고 키워드에 '', ** ** 이런표시 하지마.
                                    * 5개 내용 다 비슷하게 적지마 
                                """
                    
                    prompt_parts.append(prompt_text)
                    
                    response = model.generate_content(
                        prompt_parts,
                        generation_config=genai.types.GenerationConfig(
                            temperature=0.8,
                            top_p=0.9,
                            top_k=40,
                            max_output_tokens=5000,
                        ),
                    )
                    
                    if response and hasattr(response, 'text'):
                        foreword_text = response.text.strip()
                        
                        # 응답 검증
                        if len(foreword_text) < 400:
                            print(f"  - 경고: 생성된 텍스트가 너무 짧습니다. 재시도...")
                            retry_count += 1
                            continue
                        
                        generated_forewords.append(foreword_text)
                        
                        print(f"  ✓ 발간사 생성 완료 ({len(foreword_text)}자)")
                        print(f"  미리보기: {foreword_text[:80]}...")
                        
                        # 한글 파일 생성
                        template_to_use = available_templates[i-1] if i <= len(available_templates) else (available_templates[0] if available_templates else "")
                        output_filename = f"발간사_{i:02d}_{variation['focus'].replace(' ', '_')}.hwp"
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
                print(f"  ✗ 최종 실패. 하드코딩된 기본 발간사를 사용합니다.")
                
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
            print(f"\n[{i}번째 발간사 - {foreword_variations[i-1]['focus']}]")
            print("-" * 40)
            print(foreword)
            print("-" * 40)
        
        print(f"\n✅ 모든 작업 완료!")
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
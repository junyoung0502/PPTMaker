import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import base64

def create_ppt(title, text_content):
    """텍스트를 두 줄씩 나누어 PPT 생성 (상단 배치 + 볼드 + 스타일 통일)"""
    
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor

    # 새 프레젠테이션 생성
    prs = Presentation()
    
    # 슬라이드 크기 설정 (16:9 비율)
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    # 빈 레이아웃 사용
    blank_slide_layout = prs.slide_layouts[6]
    
    # 제목 슬라이드 생성
    title_slide = prs.slides.add_slide(blank_slide_layout)
    fill = title_slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 0, 0)
    
    # 제목 텍스트박스 (위치 상단 이동)
    title_box = title_slide.shapes.add_textbox(
        Inches(1), Inches(1.5), Inches(11.33), Inches(2.5)  # Y값 2.5 → 1.5
    )
    title_frame = title_box.text_frame
    title_frame.text = title
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.alignment = PP_ALIGN.CENTER
    title_run = title_paragraph.runs[0]
    title_run.font.name = '맑은 고딕'
    title_run.font.size = Pt(48)
    title_run.font.color.rgb = RGBColor(255, 255, 255)
    title_run.font.bold = True
    
    # 텍스트를 줄바꿈으로 분할
    lines = [line.strip() for line in text_content.split('\n') if line.strip()]
    
    # 두 줄씩 묶어서 슬라이드 생성
    for i in range(0, len(lines), 2):
        line1 = lines[i] if i < len(lines) else ""
        line2 = lines[i + 1] if i + 1 < len(lines) else ""
        
        if line1 or line2:
            slide = prs.slides.add_slide(blank_slide_layout)
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(0, 0, 0)
            
            # 텍스트박스 (위치 상단 이동)
            textbox = slide.shapes.add_textbox(
                Inches(1), Inches(1), Inches(11.33), Inches(3)  # Y값 2 → 1
            )
            text_frame = textbox.text_frame
            text_frame.clear()
            
            # 각 줄을 별도의 Paragraph로 추가하고 스타일 동일하게 적용
            for idx, line in enumerate([line1, line2]):
                if not line:
                    continue
                p = text_frame.add_paragraph() if idx > 0 else text_frame.paragraphs[0]
                p.text = line
                p.alignment = PP_ALIGN.CENTER
                for run in p.runs:
                    run.font.name = '맑은 고딕'
                    run.font.size = Pt(36)
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.bold = True
    
    return prs

# def create_ppt(title, text_content):
#     """텍스트를 두 줄씩 나누어 PPT 생성"""
    
#     # 새 프레젠테이션 생성
#     prs = Presentation()
    
#     # 슬라이드 크기 설정 (16:9 비율)
#     prs.slide_width = Inches(13.33)
#     prs.slide_height = Inches(7.5)
    
#     # 빈 레이아웃 사용 (레이아웃 인덱스 6은 보통 빈 슬라이드)
#     blank_slide_layout = prs.slide_layouts[6]
    
#     # 제목 슬라이드 생성
#     title_slide = prs.slides.add_slide(blank_slide_layout)
    
#     # 제목 슬라이드 배경을 검은색으로 설정
#     background = title_slide.background
#     fill = background.fill
#     fill.solid()
#     fill.fore_color.rgb = RGBColor(0, 0, 0)
    
#     # 제목 텍스트박스 추가
#     title_box = title_slide.shapes.add_textbox(
#         Inches(1), Inches(2.5), Inches(11.33), Inches(2.5)
#     )
#     title_frame = title_box.text_frame
#     title_frame.text = title
    
#     # 제목 텍스트 스타일 설정
#     title_paragraph = title_frame.paragraphs[0]
#     title_paragraph.alignment = PP_ALIGN.CENTER
#     title_run = title_paragraph.runs[0]
#     title_run.font.name = '맑은 고딕'
#     title_run.font.size = Pt(48)
#     title_run.font.color.rgb = RGBColor(255, 255, 255)
#     title_run.font.bold = True
    
#     # 텍스트를 줄바꿈으로 분할
#     lines = [line.strip() for line in text_content.split('\n') if line.strip()]
    
#     # 두 줄씩 묶어서 슬라이드 생성
#     for i in range(0, len(lines), 2):
#         line1 = lines[i] if i < len(lines) else ""
#         line2 = lines[i + 1] if i + 1 < len(lines) else ""
        
#         # 빈 줄이 아닌 경우에만 슬라이드 생성
#         if line1 or line2:
#             slide = prs.slides.add_slide(blank_slide_layout)
            
#             # 슬라이드 배경을 검은색으로 설정
#             background = slide.background
#             fill = background.fill
#             fill.solid()
#             fill.fore_color.rgb = RGBColor(0, 0, 0)
            
#             # 텍스트 내용 결합
#             content = line1
#             if line2:
#                 content += "\n" + line2
            
#             # 텍스트박스 추가 (가운데 상단에 위치)
#             textbox = slide.shapes.add_textbox(
#                 Inches(1), Inches(1), Inches(11.33), Inches(3)
#             )
#             text_frame = textbox.text_frame
#             text_frame.text = content
            
#             # 텍스트 스타일 설정
#             paragraph = text_frame.paragraphs[0]
#             paragraph.alignment = PP_ALIGN.CENTER
            
#             for run in paragraph.runs:
#                 run.font.name = '맑은 고딕'
#                 run.font.size = Pt(36)
#                 run.font.color.rgb = RGBColor(255, 255, 255)
#                 run.font.bold = True  # 볼드 처리
    
#     return prs

def get_ppt_download_link(prs, filename):
    """PPT 파일 다운로드 링크 생성"""
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    
    b64 = base64.b64encode(buffer.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{filename}">📥 PPT 파일 다운로드</a>'
    return href

def split_text_preview(text_content):
    """텍스트를 두 줄씩 나누어 미리보기용 리스트 반환"""
    lines = [line.strip() for line in text_content.split('\n') if line.strip()]
    slides = []
    
    for i in range(0, len(lines), 2):
        line1 = lines[i] if i < len(lines) else ""
        line2 = lines[i + 1] if i + 1 < len(lines) else ""
        
        if line1 or line2:
            content = line1
            if line2:
                content += "\n" + line2
            slides.append(content)
    
    return slides

# Streamlit 앱 설정
st.set_page_config(
    page_title="텍스트 to PPT 생성기",
    page_icon="📊",
    layout="wide"
)

# 제목
st.title("📊 텍스트 to PPT 생성기")
st.markdown("---")

# 설명
st.markdown("""
### 🎯 사용법
1. **PPT 제목**을 입력하세요
2. **텍스트 내용**을 입력하세요 (두 줄씩 자동으로 분할됩니다)
3. **미리보기**를 확인하세요
4. **PPT 생성** 버튼을 눌러 다운로드하세요

💡 **팁**: 각 슬라이드는 검은 배경에 흰 글씨로, 가운데 상단에 표시됩니다.
""")

st.markdown("---")

# 두 개의 컬럼으로 레이아웃 구성
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📝 입력")
    
    # PPT 제목 입력
    ppt_title = st.text_input(
        "PPT 제목", 
        value="꽃들도",
        help="프레젠테이션의 제목을 입력하세요"
    )
    
    # 텍스트 내용 입력
    default_text = """이 곳에 생명샘 솟아나
눈물 골짝 지나갈 때에
머잖아 열매 맺히고
웃음 소리 넘쳐나리라
이 곳에 생명샘 솟아나
눈물 골짝 지나갈 때에
머잖아 열매 맺히고
웃음 소리 넘쳐나리라
꽃들도 구름도
바람도 넓은 바다도
찬양하라 찬양하라 예수를
하늘을 울리며 노래해
나의 영혼아
은혜의 주 은혜의 주 은혜의 주
그날에 하늘이 열리고
모든 이가 보게 되리라
마침내 꽃들이 피고
영광의 주가 오시리라
꽃들도 구름도
바람도 넓은 바다도
찬양하라 찬양하라 예수를
하늘을 울리며 노래해
나의 영혼아
은혜의 주 은혜의 주 은혜의 주"""
    
    text_content = st.text_area(
        "텍스트 내용", 
        value=default_text,
        height=400,
        help="두 줄씩 분할될 텍스트를 입력하세요"
    )

with col2:
    st.subheader("👀 미리보기")
    
    if text_content.strip():
        slides = split_text_preview(text_content)
        
        st.info(f"총 {len(slides) + 1}개의 슬라이드가 생성됩니다 (제목 슬라이드 포함)")
        
        # 제목 슬라이드 미리보기
        st.markdown("**슬라이드 1 (제목)**")
        st.markdown(f"""
        <div style='background-color: black; color: white; padding: 30px; text-align: center; border-radius: 10px; margin-bottom: 20px;'>
            <h2 style='color: white; margin: 0;'>{ppt_title}</h2>
        </div>
        """, unsafe_allow_html=True)
        
        # 내용 슬라이드들 미리보기
        for idx, slide_content in enumerate(slides, 2):
            st.markdown(f"**슬라이드 {idx}**")
            formatted_content = slide_content.replace('\n', '<br>')
            st.markdown(f"""
            <div style='background-color: black; color: white; padding: 30px; text-align: center; border-radius: 10px; margin-bottom: 15px;'>
                <div style='font-size: 18px; line-height: 1.6;'>{formatted_content}</div>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.warning("텍스트를 입력하면 미리보기가 표시됩니다.")

# PPT 생성 버튼
st.markdown("---")
if st.button("🎯 PPT 생성 및 다운로드", type="primary", use_container_width=True):
    if not text_content.strip():
        st.error("텍스트를 입력해주세요!")
    else:
        with st.spinner('PPT를 생성하고 있습니다...'):
            try:
                # PPT 생성
                prs = create_ppt(ppt_title, text_content)
                
                # 다운로드 링크 생성
                filename = f"{ppt_title}.pptx"
                download_link = get_ppt_download_link(prs, filename)
                
                st.success("PPT가 성공적으로 생성되었습니다!")
                st.markdown(download_link, unsafe_allow_html=True)
                
                # 추가 정보
                slides_count = len(split_text_preview(text_content)) + 1
                st.info(f"✅ {slides_count}개의 슬라이드가 생성되었습니다.")
                
            except Exception as e:
                st.error(f"PPT 생성 중 오류가 발생했습니다: {str(e)}")

# 사이드바에 추가 정보
with st.sidebar:
    st.header("📚 도움말")
    st.markdown("""
    ### 기능 설명
    - **자동 분할**: 텍스트를 두 줄씩 자동 분할
    - **검은 배경**: 모든 슬라이드가 검은 배경
    - **흰색 텍스트**: 가운데 상단에 흰색으로 표시
    - **제목 슬라이드**: 첫 번째 슬라이드는 제목용
    
    ### 팁
    - 빈 줄은 자동으로 제거됩니다
    - 한 줄만 남는 경우도 슬라이드로 생성됩니다
    - 미리보기에서 결과를 먼저 확인해보세요
    """)
    
    st.markdown("---")
    st.markdown("**문의**: jylee0005@gmail.com")
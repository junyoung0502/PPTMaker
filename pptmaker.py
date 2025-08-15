import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import base64

def create_ppt(title, text_content, font_name, font_size):
    """텍스트를 분할하여 PPT 생성 (간단한 방식)"""
    
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
    
    # 제목 텍스트박스
    title_box = title_slide.shapes.add_textbox(
        Inches(1), Inches(1.5), Inches(11.33), Inches(2.5)
    )
    title_frame = title_box.text_frame
    title_frame.text = title
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.alignment = PP_ALIGN.CENTER
    title_run = title_paragraph.runs[0]
    title_run.font.name = font_name
    title_run.font.size = Pt(font_size + 6)
    title_run.font.color.rgb = RGBColor(255, 255, 255)
    title_run.font.bold = True
    
    # 텍스트를 슬라이드 단위로 분할 (간단한 방식)
    slides_content = parse_text_to_slides(text_content)
    
    # 각 슬라이드 생성
    for slide_text in slides_content:
        if slide_text.strip():  # 빈 슬라이드 제외
            slide = prs.slides.add_slide(blank_slide_layout)
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(0, 0, 0)
            
            # 텍스트박스 추가
            textbox = slide.shapes.add_textbox(
                Inches(1), Inches(0.5), Inches(11.33), Inches(3)
            )
            text_frame = textbox.text_frame
            text_frame.text = slide_text
            
            # 텍스트 스타일 적용
            for paragraph in text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER
                for run in paragraph.runs:
                    run.font.name = font_name
                    run.font.size = Pt(font_size)
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.bold = True
    
    return prs

def parse_text_to_slides(text_content):
    """텍스트를 슬라이드별로 분할하는 간단한 함수"""
    lines = [line.strip() for line in text_content.split('\n') if line.strip()]
    slides = []
    current_slide = []
    
    for line in lines:
        if line == '---':  # 구분자 발견
            if current_slide:  # 현재 슬라이드에 내용이 있으면 저장
                slides.append('\n'.join(current_slide))
                current_slide = []
        else:
            current_slide.append(line)
            
            # 두 줄이 쌓이면 자동으로 슬라이드 완성
            if len(current_slide) == 2:
                slides.append('\n'.join(current_slide))
                current_slide = []
    
    # 마지막에 남은 내용 처리
    if current_slide:
        slides.append('\n'.join(current_slide))
    
    return slides

def get_ppt_download_link(prs, filename):
    """PPT 파일 다운로드 링크 생성"""
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    
    b64 = base64.b64encode(buffer.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{filename}">📥 PPT 파일 다운로드</a>'
    return href

def split_text_preview(text_content):
    """텍스트를 분할하여 미리보기용 리스트 반환 (간단한 방식)"""
    return parse_text_to_slides(text_content)

# Streamlit 앱 설정
st.set_page_config(
    page_title="Euodia lyrics PPT",
    page_icon="📊",
    layout="wide"
)

# 제목
st.title("Euodia lyrics PPT")
st.markdown("---")

# 설명
st.markdown("""
### 사용법
1. **PPT 제목**을 입력하세요.
2. **텍스트 내용**을 입력하세요.
   - 기본적으로 두 줄씩 자동 분할됩니다.
     (빈줄은 자동으로 무시합니다.)
   - 한 줄만 넣고 싶으면 해당 줄 다음에 `---`를 입력하세요.
   - 오류가 발생할 수 있으니 한번 더 확인 부탁드립니다.
3. **폰트와 글자 크기**를 선택하세요.
   - 글자 크기: 12~100
4. **미리보기**를 확인하세요.
5. **PPT 생성** 버튼을 눌러 다운로드하세요.
""")

st.markdown("---")

# 세 개의 컬럼으로 레이아웃 구성
col1, col2, col3 = st.columns([1, 0.8, 1])

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
---
눈물 골짝 지나갈 때에
머잖아 열매 맺히고
웃음 소리 넘쳐나리라
꽃들도 구름도
바람도 넓은 바다도
찬양하라 찬양하라 예수를
하늘을 울리며 노래해
나의 영혼아
은혜의 주 은혜의 주 은혜의 주
---
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
        help="두 줄씩 분할될 텍스트를 입력하세요. 한 줄만 넣고 싶으면 다음 줄에 '---'를 입력하세요."
    )

with col2:
    st.subheader("⚙️ 설정")
    
    # 폰트 선택
    font_options = {
        '맑은 고딕': 'Malgun Gothic',
        '굴림': 'Gulim',
        '돋움': 'Dotum',
        '바탕': 'Batang',
        'Arial': 'Arial',
        'Times New Roman': 'Times New Roman',
        'Calibri': 'Calibri',
        'Helvetica': 'Helvetica'
    }
    
    selected_font_display = st.selectbox(
        "폰트 선택",
        list(font_options.keys()),
        index=0,
        help="사용할 폰트를 선택하세요"
    )
    
    selected_font = font_options[selected_font_display]
    
    # 글자 크기 선택 방식
    size_input_method = st.radio(
        "글자 크기 입력 방식",
        ["슬라이더", "직접 입력"],
        horizontal=True,
        help="슬라이더 또는 직접 입력 중 선택하세요"
    )
    
    if size_input_method == "슬라이더":
        font_size = st.slider(
            "글자 크기 (pt)",
            min_value=12,
            max_value=100,
            value=54,
            step=2,
            help="본문 글자 크기를 선택하세요 (제목은 자동으로 6pt 더 크게)"
        )
    else:
        font_size = st.number_input(
            "글자 크기 (pt)",
            min_value=12,
            max_value=100,
            value=54,
            step=1,
            help="본문 글자 크기를 직접 입력하세요 (제목은 자동으로 6pt 더 크게)"
        )
    
    st.markdown("---")
    st.markdown("**현재 설정**")
    st.write(f"**폰트**: {selected_font_display}")
    st.write(f"**본문 크기**: {font_size}pt")
    st.write(f"**제목 크기**: {font_size + 6}pt")

with col3:
    st.subheader("👀 미리보기")
    
    if text_content.strip():
        slides = split_text_preview(text_content)
        
        st.info(f"총 {len(slides) + 1}개의 슬라이드가 생성됩니다 (제목 슬라이드 포함)")
        
        # 제목 슬라이드 미리보기
        st.markdown("**슬라이드 1 (제목)**")
        st.markdown(f"""
        <div style='background-color: black; color: white; padding: 30px; text-align: center; border-radius: 10px; margin-bottom: 20px; font-family: {selected_font};'>
            <h2 style='color: white; margin: 0; font-size: {min(font_size + 6, 32)}px;'>{ppt_title}</h2>
        </div>
        """, unsafe_allow_html=True)
        
        # 내용 슬라이드들 미리보기
        for idx, slide_content in enumerate(slides, 2):
            st.markdown(f"**슬라이드 {idx}**")
            formatted_content = slide_content.replace('\n', '<br>')
            preview_size = min(font_size * 0.4, 20)  # 미리보기용 크기 조정
            st.markdown(f"""
            <div style='background-color: black; color: white; padding: 20px; text-align: center; border-radius: 10px; margin-bottom: 15px; font-family: {selected_font};'>
                <div style='font-size: {preview_size}px; line-height: 1.6; font-weight: bold;'>{formatted_content}</div>
            </div>
            """, unsafe_allow_html=True)
            
        # # 디버깅 정보 (개발자용)
        # with st.expander("🔍 분할 결과 확인 (디버깅용)"):
        #     st.write("**원본 텍스트 줄 수:**", len([line for line in text_content.split('\n') if line.strip()]))
        #     st.write("**생성된 슬라이드 수:**", len(slides))
        #     for i, slide in enumerate(slides):
        #         st.write(f"슬라이드 {i+2}: {repr(slide)}")  # repr로 줄바꿈 문자 확인
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
                prs = create_ppt(ppt_title, text_content, selected_font, font_size)
                
                # 다운로드 링크 생성
                filename = f"{ppt_title}.pptx"
                download_link = get_ppt_download_link(prs, filename)
                
                st.success("PPT가 성공적으로 생성되었습니다!")
                st.markdown(download_link, unsafe_allow_html=True)
                
                # 추가 정보
                slides_count = len(split_text_preview(text_content)) + 1
                st.info(f"✅ {slides_count}개의 슬라이드가 생성되었습니다.")
                st.info(f"📝 폰트: {selected_font_display}, 크기: {font_size}pt")
                
            except Exception as e:
                st.error(f"PPT 생성 중 오류가 발생했습니다: {str(e)}")

# 사이드바에 추가 정보
with st.sidebar:
    st.header("도움말")
    
    st.markdown("---")
    st.markdown("**문의**: jylee0005@gmail.com")
    


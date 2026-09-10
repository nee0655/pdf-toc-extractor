import streamlit as st
import pdfplumber
import re
import pandas as pd
import io

# ---------------------------------------------------------
# Streamlit 기본 설정 (1GB 대용량 업로드 지원)
# ---------------------------------------------------------
st.set_page_config(page_title="PDF 사업명 & 쪽 번호 정밀 추출기 (1GB 지원)", page_icon="📊", layout="wide")

# 대용량 파일 처리를 위한 메모리 한도 안내
st.title("📊 PDF 사업명 & 쪽 번호 정밀 추출기 (대용량 1GB)")
st.caption("사업명 영역과 쪽 번호 영역을 각각 지정하여 왜곡된 글자까지 정밀 추출합니다. (최대 1GB 파일 업로드 지원)")

# ---------------------------------------------------------
# Excel 파일 생성 함수
# ---------------------------------------------------------
def create_excel_file(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='사업명_목차_리스트')
        worksheet = writer.sheets['사업명_목차_리스트']
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len * 1.5, 12)
        worksheet.column_dimensions['B'].width = 50
        worksheet.column_dimensions['D'].width = 55
    return output.getvalue()

# ---------------------------------------------------------
# PDF 파싱 함수 (기존 복합 키워드 + 신규/전환사업 패턴 복원)
# ---------------------------------------------------------
def parse_pdf_titles_comprehensive(pdf_file, x_start, x_end, page_y_start, page_y_end, title_y_start, title_y_end):
    results = []
    
    # 1. 예산서 특화 키워드 전체 복원 (신규/전환사업 등 확장 키워드 통합)
    keywords = [
        r'신규\s*[\/\,]\s*전환\s*사업', r'계속\s*[\/\,]\s*전환\s*사업',
        r'신규\s*[\/\,]\s*전환', r'계속\s*[\/\,]\s*전환', r'전환\s*사업', r'전환사업', r'전환',
        r'신규\s*[\/\,]\s*소멸\s*사업', r'신규\s*[\/\,]\s*소멸',
        r'계속\s*[\/\,]\s*소멸\s*기금', r'계속\s*[\/\,]\s*소멸',
        r'소멸\s*기금', r'소멸기금', r'계속', r'신규', r'소멸', r'자체', r'보조'
    ]
    pattern_str = r'[\(\[\<]\s*(' + '|'.join(keywords) + r')\s*[\)\]\>]'
    keyword_pattern = re.compile(pattern_str)

    visited_titles = set()

    with pdfplumber.open(pdf_file) as pdf:
        for pdf_page_idx, page in enumerate(pdf.pages):
            pw, ph = page.width, page.height
            
            # A. 하단 지정 Y축에서 쪽 번호 스캔
            page_bbox = (
                pw * (x_start / 100.0),
                ph * (page_y_start / 100.0),
                pw * (x_end / 100.0),
                ph * (page_y_end / 100.0)
            )
            try:
                page_crop = page.crop(page_bbox)
                page_crop_text = page_crop.extract_text() or ""
            except Exception:
                page_crop_text = ""
            
            page_match = re.search(r'(?:-\s*|\[)?(\d{1,4})(?:\s*-|\])?', page_crop_text)
            detected_page = int(page_match.group(1)) if page_match else (pdf_page_idx + 1)

            # B. 상단 지정 Y축에서 사업명 제목 스캔
            title_bbox = (
                0,
                ph * (title_y_start / 100.0),
                pw,
                ph * (title_y_end / 100.0)
            )
            try:
                title_crop = page.crop(title_bbox)
                title_crop_text = title_crop.extract_text(layout=False) or ""
            except Exception:
                title_crop_text = page.extract_text(layout=False) or ""

            lines = [line.strip() for line in title_crop_text.split('\n') if line.strip()]
            
            for line in lines:
                clean_line = re.sub(r'\s+', ' ', line).strip()
                
                # 대괄호/꺾쇠괄호를 소괄호()로 포맷 통일
                formatted_line = re.sub(r'[\[\<](.*?)[\]\>]', r'(\1)', clean_line)
                
                # [조건 1] 확장 키워드 패턴 매칭
                is_keyword_matched = bool(keyword_pattern.search(formatted_line))
                
                # [조건 2] 보조 매칭: 괄호()와 /가 동시에 들어가 있고 단어 길이가 150자 이내인 제목
                is_slash_bracket_matched = ('(' in formatted_line and ')' in formatted_line and '/' in formatted_line)
                
                if (is_keyword_matched or is_slash_bracket_matched) and len(formatted_line) <= 150:
                    formatted_title = re.sub(r'\s+', ' ', formatted_line)
                    formatted_title = re.sub(r'^[ㆍ•■□▶-]\s*', '', formatted_title)
                    
                    if formatted_title not in visited_titles:
                        visited_titles.add(formatted_title)
                        results.append({
                            'pdf_page': pdf_page_idx + 1,
                            'detected_page': detected_page,
                            'title': formatted_title
                        })

    return results

# ---------------------------------------------------------
# UI 구동부
# ---------------------------------------------------------
uploaded_file = st.file_uploader("PDF 파일 업로드 (대용량 1GB 파일 지원)", type=["pdf"])

if uploaded_file is not None:
    st.sidebar.header("📌 1. 사업명 제목 Y축 영역 지정")
    st.sidebar.caption("페이지 상단~중단에서 사업명 제목이 위치한 영역을 지정하세요.")
    title_y_start = st.sidebar.slider("사업명 시작 Y축 (%)", min_value=0, max_value=80, value=0)
    title_y_end = st.sidebar.slider("사업명 종료 Y축 (%)", min_value=10, max_value=90, value=50)

    st.sidebar.divider()

    st.sidebar.header("🎯 2. 쪽 번호 Y축/X축 영역 지정")
    st.sidebar.caption("페이지 하단에서 쪽 번호가 적힌 영역을 지정하세요.")
    page_y_start = st.sidebar.slider("쪽 번호 시작 Y축 (%)", min_value=70, max_value=99, value=90)
    page_y_end = st.sidebar.slider("쪽 번호 종료 Y축 (%)", min_value=75, max_value=100, value=100)

    col_x1, col_x2 = st.sidebar.columns(2)
    with col_x1:
        x_start_pct = st.number_input("X축 시작 (%)", min_value=0, max_value=100, value=25)
    with col_x2:
        x_end_pct = st.number_input("X축 종료 (%)", min_value=0, max_value=100, value=75)

    # 5페이지 텍스트 크롭 즉시 검증
    st.subheader("🔍 지정 영역 텍스트 검증 (앞 5페이지)")
    
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            max_p = min(5, len(pdf.pages))
            tabs = st.tabs([f"{i+1}페이지" for i in range(max_p)])
            
            for idx, tab in enumerate(tabs):
                with tab:
                    p = pdf.pages[idx]
                    
                    p_bbox = (p.width * (x_start_pct / 100.0), p.height * (page_y_start / 100.0), p.width * (x_end_pct / 100.0), p.height * (page_y_end / 100.0))
                    page_txt = p.crop(p_bbox).extract_text() or ""
                    
                    t_bbox = (0, p.height * (title_y_start / 100.0), p.width, p.height * (title_y_end / 100.0))
                    title_txt = p.crop(t_bbox).extract_text() or ""
                    
                    col_t, col_p = st.columns([2, 1])
                    with col_t:
                        st.info(f"🏷️ **사업명 영역 읽은 내용:**\n\n`{title_txt.strip().replace('\n', ' ')[:100]}...`")
                    with col_p:
                        st.success(f"📌 **쪽 번호 영역 읽은 글자:**\n\n`{page_txt.strip().replace('\n', ' ')}`")
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")

    st.divider()

    # 전체 추출 실행 버튼
    st.subheader("🚀 사업명 및 쪽 번호 전체 추출")
    run_button = st.button("🚀 설정한 영역으로 전체 문서 엑셀 추출하기", use_container_width=True, type="primary")

    if run_button or 'parsed_cache' in st.session_state:
        if run_button:
            with st.spinner("전체 대용량 PDF 문서에서 정밀 키워드로 분석 중입니다..."):
                st.session_state['parsed_cache'] = parse_pdf_titles_comprehensive(
                    uploaded_file, x_start_pct, x_end_pct, page_y_start, page_y_end, title_y_start, title_y_end
                )

        parsed_data = st.session_state.get('parsed_cache', [])

        if parsed_data:
            st.success(f"총 {len(parsed_data)}개의 사업명을 발견 및 추출했습니다!")
            
            initial_df = pd.DataFrame([
                {
                    "연번": idx + 1,
                    "사업명": item['title'],
                    "PDF 순서": f"{item['pdf_page']}p",
                    "추출된 쪽 번호": item['detected_page']
                }
                for idx, item in enumerate(parsed_data)
            ])

            edited_df = st.data_editor(
                initial_df,
                column_config={
                    "연번": st.column_config.NumberColumn(disabled=True),
                    "사업명": st.column_config.TextColumn(disabled=True, width="large"),
                    "PDF 순서": st.column_config.TextColumn(disabled=True),
                    "추출된 쪽 번호": st.column_config.NumberColumn(min_value=1, step=1, required=True)
                },
                hide_index=True,
                use_container_width=True
            )

            edited_df["최종 목차 문자열"] = edited_df.apply(
                lambda row: f"{row['사업명']}-{row['추출된 쪽 번호']}", axis=1
            )

            raw_lines = "\n".join(edited_df['최종 목차 문자열'].tolist())

            st.divider()
            st.subheader("📥 엑셀 파일 다운로드")
            
            col1, col2 = st.columns(2)
            
            with col1:
                export_df = edited_df[["연번", "사업명", "추출된 쪽 번호", "최종 목차 문자열"]].rename(
                    columns={"추출된 쪽 번호": "페이지 번호"}
                )
                excel_bytes = create_excel_file(export_df)
                st.download_button(
                    label="📊 엑셀(XLSX) 파일로 다운로드",
                    data=excel_bytes,
                    file_name=f"{uploaded_file.name.replace('.pdf', '')}_목차.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col2:
                st.download_button(
                    label="📝 텍스트(TXT) 파일로 다운로드",
                    data=raw_lines.encode('utf-8'),
                    file_name=f"{uploaded_file.name.replace('.pdf', '')}_목록.txt",
                    mime="text/plain"
                )
        else:
            st.error("지정된 영역에서 사업명을 찾지 못했습니다.")
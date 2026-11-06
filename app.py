import streamlit as st
import pdfplumber
import re
import pandas as pd
import io

# ---------------------------------------------------------
# Streamlit 기본 설정
# ---------------------------------------------------------
st.set_page_config(page_title="PDF 사업명 & 쪽 번호 정밀 추출기", page_icon="📊", layout="wide")

st.title("📊 PDF 사업명 & 쪽 번호 정밀 추출기")
st.caption("설정한 글자 크기(Font Size) 이상이거나, 키워드/기호 조건에 해당하면 모두 정밀 추출합니다.")

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
# 하이브리드 파싱 함수 (글자 크기 이상 OR 당초 키워드/기호 조건)
# ---------------------------------------------------------
def parse_pdf_titles_hybrid(pdf_bytes, x_start, x_end, page_y_start, page_y_end, title_y_start, title_y_end, min_font_size):
    results = []
    
    # 키워드 목록
    keywords = [
        r'신규\s*[\/\,]\s*전환\s*사업', r'계속\s*[\/\,]\s*전환\s*사업',
        r'신규\s*[\/\,]\s*전환', r'계속\s*[\/\,]\s*전환', r'전환\s*사업', r'전환사업', r'전환',
        r'신규\s*[\/\,]\s*소멸\s*사업', r'신규\s*[\/\,]\s*소멸',
        r'계속\s*[\/\,]\s*소멸\s*기금', r'계속\s*[\/\,]\s*소멸',
        r'소멸\s*기금', r'소멸기금', r'계속', r'신규', r'소멸', r'자체', r'보조'
    ]
    # [수정] \vert{} 문법 오류를 올바른 | (OR) 정규식으로 수정
    pattern_str = r'[\(\[\<]\s*(' + '|'.join(keywords) + r')\s*[\)\]\>]'
    keyword_pattern = re.compile(pattern_str)

    visited_titles = set()

    # [수정] io.BytesIO를 통해 어느 환경에서나 안전하게 PDF 오픈
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pdf_page_idx, page in enumerate(pdf.pages):
            pw, ph = page.width, page.height
            
            # 1. 쪽 번호 스캔 영역 (하단 지정 Y축)
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
            
            # [수정] 쪽 번호 추출 정규식 이스케이프 오류 수정
            page_match = re.search(r'(?:-\s*|\[)?(\d{1,4})(?:\s*-\s*|\])?', page_crop_text)
            detected_page = int(page_match.group(1)) if page_match else (pdf_page_idx + 1)

            # 2. 사업명 영역 스캔 (글자 크기 및 단어 속성 포함)
            title_bbox = (
                0,
                ph * (title_y_start / 100.0),
                pw,
                ph * (title_y_end / 100.0)
            )
            
            try:
                title_crop = page.crop(title_bbox)
                words = title_crop.extract_words(extra_attrs=["size"])
            except Exception:
                words = []

            if not words:
                continue

            # 라인 단위(Y좌표 기준) 그룹화
            lines_words = {}
            lines_max_size = {}
            
            for w in words:
                top_key = round(w["top"], 1)
                lines_words.setdefault(top_key, []).append(w["text"])
                lines_max_size[top_key] = max(lines_max_size.get(top_key, 0), w.get("size", 0))

            for top_key in sorted(lines_words.keys()):
                full_line = " ".join(lines_words[top_key]).strip()
                clean_line = re.sub(r'\s+', ' ', full_line)
                formatted_line = re.sub(r'[\[\<](.*?)[\]\>]', r'(\1)', clean_line)
                
                line_max_font = lines_max_size.get(top_key, 0)
                
                # [판정 조건 A] 지정한 글자 크기 이상인 경우
                is_large_enough = (line_max_font >= min_font_size)
                
                # [판정 조건 B] 키워드 또는 괄호/기호 조건
                is_keyword_matched = bool(keyword_pattern.search(formatted_line))
                has_bracket_or_slash = ('(' in formatted_line and ')' in formatted_line) or ('/' in formatted_line)
                
                if (is_large_enough or is_keyword_matched or has_bracket_or_slash) and len(formatted_line) <= 150:
                    formatted_title = re.sub(r'^[ㆍ•■□▶-]\s*', '', formatted_line)
                    
                    if len(formatted_title) >= 2 and formatted_title not in visited_titles:
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
uploaded_file = st.file_uploader("PDF 파일 업로드", type=["pdf"])

if uploaded_file is not None:
    # 안전한 파일 바이트 추출
    pdf_bytes = uploaded_file.getvalue()

    st.sidebar.header("📏 1. 사업명 글자 크기(Font Size) 설정")
    st.sidebar.caption("제목으로 간주할 글자 크기 기준을 정하세요.")
    min_font_size = st.sidebar.slider("최소 글자 크기 (pt)", min_value=8.0, max_value=25.0, value=11.0, step=0.5)

    st.sidebar.divider()

    st.sidebar.header("📌 2. 사업명 제목 Y축 영역 지정")
    title_y_start = st.sidebar.slider("사업명 시작 Y축 (%)", min_value=0, max_value=80, value=0)
    title_y_end = st.sidebar.slider("사업명 종료 Y축 (%)", min_value=10, max_value=90, value=50)

    st.sidebar.divider()

    st.sidebar.header("🎯 3. 쪽 번호 Y축/X축 영역 지정")
    page_y_start = st.sidebar.slider("쪽 번호 시작 Y축 (%)", min_value=70, max_value=99, value=90)
    page_y_end = st.sidebar.slider("쪽 번호 종료 Y축 (%)", min_value=75, max_value=100, value=100)

    col_x1, col_x2 = st.sidebar.columns(2)
    with col_x1:
        x_start_pct = st.number_input("X축 시작 (%)", min_value=0, max_value=100, value=25)
    with col_x2:
        x_end_pct = st.number_input("X축 종료 (%)", min_value=0, max_value=100, value=75)

    # 5페이지 실시간 검증
    st.subheader("🔍 지정 영역 & 글자 크기 검증 (앞 5페이지)")
    
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            max_p = min(5, len(pdf.pages))
            tabs = st.tabs([f"{i+1}페이지" for i in range(max_p)])
            
            for idx, tab in enumerate(tabs):
                with tab:
                    p = pdf.pages[idx]
                    
                    # 쪽 번호 크롭
                    p_bbox = (p.width * (x_start_pct / 100.0), p.height * (page_y_start / 100.0), p.width * (x_end_pct / 100.0), p.height * (page_y_end / 100.0))
                    page_txt = p.crop(p_bbox).extract_text() or ""
                    
                    # 사업명 크롭 & 글자 크기
                    t_bbox = (0, p.height * (title_y_start / 100.0), p.width, p.height * (title_y_end / 100.0))
                    try:
                        t_crop = p.crop(t_bbox)
                        words = t_crop.extract_words(extra_attrs=["size"])
                        detected_items = [f"{w['text']}({w['size']:.1f}pt)" for w in words]
                        sample_txt = " ".join(detected_items) if detected_items else "(포착된 단어 없음)"
                    except Exception:
                        sample_txt = "(인식 불가)"
                    
                    col_t, col_p = st.columns([2, 1])
                    with col_t:
                        st.info(f"🏷️ **사업명 Y축 영역 단어 및 크기:**\n\n`{sample_txt[:120]}`")
                    with col_p:
                        st.success(f"📌 **쪽 번호 영역 읽은 글자:**\n\n`{page_txt.strip().replace('\n', ' ')}`")
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")

    st.divider()

    # 전체 추출 버튼
    st.subheader("🚀 사업명 및 쪽 번호 전체 추출")
    run_button = st.button("🚀 설정한 모든 조건으로 전체 문서 추출하기", use_container_width=True, type="primary")

    if run_button or 'parsed_cache' in st.session_state:
        if run_button:
            with st.spinner("하이브리드 조건(글자 크기 OR 키워드)으로 전체 분석 중입니다..."):
                st.session_state['parsed_cache'] = parse_pdf_titles_hybrid(
                    pdf_bytes, x_start_pct, x_end_pct, page_y_start, page_y_end, title_y_start, title_y_end, min_font_size
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
            st.error("조건에 맞는 사업명을 찾지 못했습니다. 슬라이더 값을 조금씩 조절해 보세요.")

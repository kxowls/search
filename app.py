import streamlit as st
import pandas as pd
import pdfplumber
import io
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# 페이지 설정
st.set_page_config(
    page_title="문서 키워드 검색 도구",
    page_icon="🔍",
    layout="wide"
)

def split_keywords(keyword_text):
    """키워드 문자열을 분리하여 리스트로 반환"""
    # 쉼표, 공백, 기호 등으로 분리
    keywords = re.split(r'[,，\s]+', keyword_text)
    # 빈 문자열 제거 및 공백 제거
    return [k.strip() for k in keywords if k.strip()]

def normalize_text(text):
    """텍스트 정규화: 공백 제거 및 소문자 변환"""
    if isinstance(text, str):
        return re.sub(r'\s+', '', text.lower())
    return str(text)

def highlight_keywords(text, keywords):
    """여러 키워드 하이라이팅"""
    if not isinstance(text, str):
        return str(text)
    
    highlighted_text = text
    for keyword in keywords:
        pattern = re.compile(f'({re.escape(keyword)})', re.IGNORECASE)
        highlighted_text = pattern.sub(r'<span style="background-color: yellow">\1</span>', highlighted_text)
    return highlighted_text

def process_pdf(file, keywords):
    results = []
    normalized_keywords = [normalize_text(k) for k in keywords]
    
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            for table_num, table in enumerate(tables, 1):
                for row_num, row in enumerate(table, 1):
                    # 각 셀의 정규화된 텍스트 확인
                    cell_texts = [str(cell) if cell else '' for cell in row]
                    normalized_cells = [normalize_text(text) for text in cell_texts]
                    
                    # 하나라도 키워드가 포함되어 있는지 확인
                    if any(any(kw in cell for kw in normalized_keywords) for cell in normalized_cells):
                        # 원본 행을 하이라이팅하여 저장
                        highlighted_row = [highlight_keywords(str(cell), keywords) if cell else '' for cell in row]
                        results.append({
                            '페이지': page_num,
                            '테이블': table_num,
                            '행': row_num,
                            '내용': highlighted_row
                        })
    return pd.DataFrame(results)

def process_excel(file, keywords):
    df = pd.read_excel(file)
    normalized_keywords = [normalize_text(k) for k in keywords]
    
    # 모든 셀을 문자열로 변환하고 정규화
    mask = df.astype(str).apply(lambda x: x.apply(
        lambda cell: any(kw in normalize_text(cell) for kw in normalized_keywords)
    ))
    filtered_df = df[mask.any(axis=1)].copy()
    
    # 검색된 행의 모든 셀에 하이라이팅 적용
    for col in filtered_df.columns:
        filtered_df[col] = filtered_df[col].apply(lambda x: highlight_keywords(str(x), keywords))
    
    return filtered_df

def main():
    st.title("📄 문서 키워드 검색 도구")
    st.write("PDF 또는 Excel 파일에서 키워드를 검색하고 결과를 확인하세요.")

    # 파일 업로드
    uploaded_file = st.file_uploader("파일을 업로드하세요", type=['pdf', 'xlsx', 'xls'])
    
    # 검색 옵션
    col1, col2 = st.columns([3, 1])
    with col1:
        keyword_text = st.text_input("검색할 키워드를 입력하세요 (쉼표로 구분)")
    with col2:
        st.write("")
        st.write("")
        st.write("※ 쉼표로 여러 키워드 검색 가능")

    if uploaded_file and keyword_text:
        try:
            # 키워드 분리
            keywords = split_keywords(keyword_text)
            
            if not keywords:
                st.warning("검색할 키워드를 입력해주세요.")
                return

            if uploaded_file.name.endswith('.pdf'):
                df = process_pdf(uploaded_file, keywords)
            else:
                df = process_excel(uploaded_file, keywords)

            if len(df) > 0:
                st.success(f"검색 결과: {len(df)}개의 항목을 찾았습니다.")
                st.markdown(
                    df.to_html(escape=False, index=False),
                    unsafe_allow_html=True
                )

                # Excel 다운로드 버튼
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                output.seek(0)
                st.download_button(
                    label="Excel로 다운로드",
                    data=output,
                    file_name="검색결과.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("검색 결과가 없습니다.")

        except Exception as e:
            st.error(f"오류가 발생했습니다: {str(e)}")

if __name__ == "__main__":
    main() 
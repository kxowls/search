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

def parse_query(query):
    query = query.replace('그리고', '&').replace('또는', '|').replace('제외', '-')
    query = query.replace('AND', '&').replace('OR', '|').replace('NOT', '-')
    query = query.replace('NEAR', '~').replace('WITHIN', '~')
    return query

def is_near(text, a, b, window=5):
    # 텍스트에서 a, b가 window 이내에 등장하는지 확인
    text = normalize_text(text)
    a = normalize_text(a)
    b = normalize_text(b)
    words = re.split(r'\s+', text)
    idx_a = [i for i, w in enumerate(words) if a in w]
    idx_b = [i for i, w in enumerate(words) if b in w]
    for i in idx_a:
        for j in idx_b:
            if abs(i - j) <= window:
                return True
    return False

def match_logic(cell, query):
    cell = normalize_text(str(cell))
    # NEAR, WITHIN
    if '~' in query:
        parts = [p.strip() for p in query.split('~')]
        if len(parts) == 2:
            return is_near(cell, parts[0], parts[1])
    # NOT
    if '-' in query:
        parts = [p.strip() for p in query.split('-')]
        must = parts[0]
        nots = parts[1:]
        if not all(match_logic(cell, must) for must in must.split('&')):
            return False
        for n in nots:
            if any(match_logic(cell, n) for n in n.split('|')):
                return False
        return True
    # AND
    if '&' in query:
        return all(match_logic(cell, q) for q in query.split('&'))
    # OR
    if '|' in query:
        return any(match_logic(cell, q) for q in query.split('|'))
    # 인용부호 exact match
    if query.startswith('"') and query.endswith('"'):
        return query[1:-1] in cell
    # 단일 키워드
    return query.strip() in cell

def process_pdf(file, query):
    results = []
    parsed_query = parse_query(query)
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            for table_num, table in enumerate(tables, 1):
                for row_num, row in enumerate(table, 1):
                    cell_texts = [str(cell) if cell else '' for cell in row]
                    if any(match_logic(cell, parsed_query) for cell in cell_texts):
                        results.append({
                            '페이지': page_num,
                            '테이블': table_num,
                            '행': row_num,
                            '내용': cell_texts
                        })
    return pd.DataFrame(results)

def process_excel(file, query):
    df = pd.read_excel(file)
    parsed_query = parse_query(query)
    mask = df.astype(str).apply(lambda x: x.apply(lambda cell: match_logic(cell, parsed_query)))
    return df[mask.any(axis=1)]

def main():
    st.title("📄 문서 키워드 검색 도구")
    st.write("PDF 또는 Excel 파일에서 키워드를 논리연산자와 함께 검색할 수 있습니다.")
    st.info("""
    **검색어 입력 예시**
    - `파이썬 & 데이터` : 두 키워드 모두 포함
    - `파이썬 | 자바` : 둘 중 하나라도 포함
    - `파이썬 -입문` : '파이썬'은 포함, '입문'은 제외
    - `파이썬 & (데이터 | 분석)` : '파이썬'과 '데이터' 또는 '분석'이 모두 포함
    - 연산자: `&`(AND), `|`(OR), `-`(NOT)
    """)
    uploaded_file = st.file_uploader("파일을 업로드하세요", type=['pdf', 'xlsx', 'xls'])
    query = st.text_input("검색할 키워드를 입력하세요 (논리연산자 사용 가능)")
    if uploaded_file and query:
        try:
            if uploaded_file.name.endswith('.pdf'):
                df = process_pdf(uploaded_file, query)
            else:
                df = process_excel(uploaded_file, query)
            if len(df) > 0:
                st.success(f"검색 결과: {len(df)}개의 항목을 찾았습니다.")
                st.markdown(df.to_html(escape=False, index=False), unsafe_allow_html=True)
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
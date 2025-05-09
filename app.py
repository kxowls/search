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
        # 공백을 하나로 통일하고 소문자로 변환
        return ' '.join(text.lower().split())
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
    """검색어 파싱: 논리 연산자 변환"""
    # NOT 연산자 변환 (! -> -)
    query = re.sub(r'!(\w+)', r'-\1', query)
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
    """검색 로직: 논리 연산자 처리"""
    cell = normalize_text(str(cell))
    
    # 괄호 처리
    if '(' in query and ')' in query:
        def replace_brackets(match):
            inner_query = match.group(1)
            return str(match_logic(cell, inner_query))
        query = re.sub(r'\((.*?)\)', replace_brackets, query)
    
    # NOT
    if '-' in query:
        parts = [p.strip() for p in query.split('-')]
        must = parts[0]
        nots = parts[1:]
        if not match_logic(cell, must):
            return False
        return not any(normalize_text(n) in cell for n in nots)
    
    # AND
    if '&' in query:
        parts = [p.strip() for p in query.split('&')]
        return all(normalize_text(part) in cell for part in parts)
    
    # OR
    if '|' in query:
        parts = [p.strip() for p in query.split('|')]
        return any(normalize_text(part) in cell for part in parts)
    
    # 구문 검색 (정확한 문구)
    if query.startswith('"') and query.endswith('"'):
        exact_phrase = normalize_text(query[1:-1])
        return exact_phrase in cell
    
    # 단일 키워드 (부분 문자열 매칭)
    query = normalize_text(query.strip())
    return query in cell

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
                        # 내용 컬럼을 문자열로 단순화
                        results.append({
                            '페이지': page_num,
                            '테이블': table_num,
                            '행': row_num,
                            '내용': ' | '.join(cell_texts)
                        })
    return pd.DataFrame(results)

def process_excel(file, query):
    """엑셀 파일 처리: 키워드가 포함된 행 전체 출력"""
    df = pd.read_excel(file)
    parsed_query = parse_query(query)
    
    # 각 행에 대해 검색 수행
    def search_row(row):
        # 각 셀의 값을 문자열로 변환하고 공백을 포함한 원본 텍스트로 검색
        row_text = ' '.join(str(cell).strip() for cell in row if pd.notna(cell))
        return match_logic(row_text, parsed_query)
    
    mask = df.apply(search_row, axis=1)
    
    # 검색된 행 전체 반환
    return df[mask]

def main():
    st.title("📄 문서 키워드 검색 도구")
    st.write("PDF 또는 Excel 파일에서 키워드를 논리연산자와 함께 검색할 수 있습니다.")
    
    # 검색 도움말
    st.info("""
    **검색 연산자 사용 가이드**
    
    | 연산자 | 의미 | 예시 | 설명 |
    |--------|------|------|------|
    | & (AND) | 모두 포함 | `교육 & 심리` | '교육'과 '심리' 모두 포함 |
    | \| (OR) | 하나라도 포함 | `교육 \| 심리` | '교육' 또는 '심리' 포함 |
    | ! (NOT) | 제외 | `교육 & !심리` | '교육'은 포함, '심리'는 제외 |
    | " " | 정확한 문구 | `"아동 발달"` | '아동 발달' 정확히 일치 |
    | ( ) | 그룹화 | `(교육 \| 심리) & 발달` | '교육' 또는 '심리'를 포함하면서 '발달' 포함 |
    
    **💡 검색 팁**
    - 공백은 무시됩니다 (예: '교육심리' = '교육 심리')
    - 대소문자를 구분하지 않습니다
    - 부분 문자열도 검색됩니다 (예: '혼자'로 '혼자공부하는파이썬' 검색 가능)
    - 엑셀 파일 검색 시 키워드가 포함된 행 전체가 출력됩니다
    - 여러 연산자를 조합하여 사용할 수 있습니다
    """)
    
    uploaded_file = st.file_uploader("파일을 업로드하세요", type=['pdf', 'xlsx', 'xls'])
    query = st.text_input("검색할 키워드를 입력하세요 (연산자: &, |, !, \"\", ())")
    if uploaded_file and query:
        try:
            if uploaded_file.name.endswith('.pdf'):
                df = process_pdf(uploaded_file, query)
            else:
                df = process_excel(uploaded_file, query)
            if len(df) > 0:
                st.success(f"검색 결과: {len(df)}개의 항목을 찾았습니다.")
                st.dataframe(df, use_container_width=True, hide_index=True)
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
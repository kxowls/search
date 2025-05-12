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

# 캐시 데코레이터 추가
@st.cache_data
def load_excel(file):
    """엑셀 파일을 로드하고 캐시합니다."""
    return pd.read_excel(file)

@st.cache_data
def get_columns(df):
    """데이터프레임의 컬럼 목록을 반환하고 캐시합니다."""
    return df.columns.tolist()

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
    # 특수 기호 변환
    query = query.replace('∣', '|')
    query = query.replace('＆', '&')
    query = query.replace('！', '!')
    
    # 공백 제거
    query = re.sub(r'\s+', '', query)
    
    # 연속된 NOT 연산자 처리
    while '!!' in query:
        query = query.replace('!!', '')
    
    # NOT 연산자 변환 (! -> -)
    query = re.sub(r'!(\w+)', r'-\1', query)
    
    return query

def is_near(text, a, b, window=5):
    """텍스트에서 a, b가 window 이내에 등장하는지 확인"""
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

def tokenize_query(query):
    """쿼리를 토큰으로 분리"""
    # 괄호, 연산자, 키워드를 분리
    pattern = r'([()&|!-])|([^()&|!-]+)'
    tokens = []
    for match in re.finditer(pattern, query):
        token = match.group(1) or match.group(2)
        if token:
            tokens.append(token)
    return tokens

def evaluate_expression(cell, tokens):
    """토큰화된 표현식을 평가"""
    cell = normalize_text(str(cell))
    stack = []
    operators = []
    
    for token in tokens:
        if token == '(':
            operators.append(token)
        elif token == ')':
            # 괄호 안의 표현식 평가
            while operators and operators[-1] != '(':
                op = operators.pop()
                if op == '&':
                    b = stack.pop()
                    a = stack.pop()
                    stack.append(a and b)
                elif op == '|':
                    b = stack.pop()
                    a = stack.pop()
                    stack.append(a or b)
            if operators and operators[-1] == '(':
                operators.pop()
        elif token in ['&', '|']:
            while operators and operators[-1] != '(' and operators[-1] in ['&', '|']:
                op = operators.pop()
                if op == '&':
                    b = stack.pop()
                    a = stack.pop()
                    stack.append(a and b)
                elif op == '|':
                    b = stack.pop()
                    a = stack.pop()
                    stack.append(a or b)
            operators.append(token)
        elif token.startswith('-'):
            # NOT 연산
            keyword = token[1:]
            stack.append(normalize_text(keyword) not in cell)
        elif token.startswith('"') and token.endswith('"'):
            # 정확한 문구 검색
            phrase = normalize_text(token[1:-1])
            stack.append(phrase in cell)
        else:
            # 일반 키워드 검색
            stack.append(normalize_text(token) in cell)
    
    # 남은 연산자 처리
    while operators:
        op = operators.pop()
        if op == '&':
            b = stack.pop()
            a = stack.pop()
            stack.append(a and b)
        elif op == '|':
            b = stack.pop()
            a = stack.pop()
            stack.append(a or b)
    
    return stack[0] if stack else False

def match_logic(cell, query):
    """검색 로직: 중첩된 논리 연산자 처리"""
    # 쿼리 파싱
    parsed_query = parse_query(query)
    
    # 토큰화
    tokens = tokenize_query(parsed_query)
    
    # 표현식 평가
    return evaluate_expression(cell, tokens)

def process_pdf(file, query):
    """PDF 파일 처리"""
    results = []
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            for table_num, table in enumerate(tables, 1):
                for row_num, row in enumerate(table, 1):
                    cell_texts = [str(cell) if cell else '' for cell in row]
                    if any(match_logic(cell, query) for cell in cell_texts):
                        results.append({
                            '페이지': page_num,
                            '테이블': table_num,
                            '행': row_num,
                            '내용': ' | '.join(cell_texts)
                        })
    return pd.DataFrame(results)

def process_excel(file, query, selected_columns=None):
    """엑셀 파일 처리: 선택된 컬럼에서만 키워드 검색"""
    df = pd.read_excel(file)
    
    # 선택된 컬럼이 없으면 모든 컬럼 사용
    if not selected_columns:
        selected_columns = df.columns.tolist()
    
    # 각 행에 대해 검색 수행
    def search_row(row):
        # 선택된 컬럼의 값만 문자열로 변환하여 검색
        row_text = ' '.join(str(row[col]).strip() for col in selected_columns if pd.notna(row[col]))
        return match_logic(row_text, query)
    
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
    | & (AND) | 모두 포함 | `파이썬 & 한빛` | '파이썬'과 '한빛' 모두 포함된 내용 검색 |
    | \| (OR) | 하나라도 포함 | `파이썬 \| 한빛` | '파이썬' 또는 '한빛' 중 하나라도 포함된 내용 검색 |
    | ! (NOT) | 제외 | `파이썬 & !한빛` | '파이썬'은 포함하되 '한빛'은 제외된 내용 검색 |
    | " " | 정확한 문구 | `"파이썬 프로그래밍"` | '파이썬 프로그래밍'이라는 정확한 문구 검색 |
    | ( ) | 그룹화 | `(파이썬 \| 한빛) & 프로그래밍` | '파이썬' 또는 '한빛'을 포함하면서 '프로그래밍'도 포함된 내용 검색 |
    
    **💡 검색 팁**
    - 공백은 무시됩니다 (예: '파이썬프로그래밍' = '파이썬 프로그래밍')
    - 대소문자를 구분하지 않습니다 (예: 'Python' = 'python')
    - 부분 문자열도 검색됩니다 (예: '파이썬'으로 '파이썬프로그래밍' 검색 가능)
    - 엑셀 파일 검색 시 키워드가 포함된 행 전체가 출력됩니다
    - 여러 연산자를 조합하여 사용할 수 있습니다
    - 복잡한 논리 연산 가능: `(기업가 & !한빛) | (기업가 & 창업)`
    
    **📝 실제 사용 예시**
    1. 파이썬 관련 모든 내용: `파이썬`
    2. 한빛미디어의 파이썬 책만: `파이썬 & 한빛`
    3. 파이썬이나 자바 관련 내용: `파이썬 | 자바`
    4. 파이썬은 포함하되 자바는 제외: `파이썬 & !자바`
    5. 정확한 책 제목 검색: `"혼자 공부하는 파이썬"`
    6. 복잡한 조건 검색: `(파이썬 | 한빛) & (기초 | 입문)`
    7. 기업가 관련 내용 중 한빛 제외 또는 창업 포함: `(기업가 & !한빛) | (기업가 & 창업)`
    """)
    
    uploaded_file = st.file_uploader("파일을 업로드하세요", type=['pdf', 'xlsx', 'xls'])
    query = st.text_input("검색할 키워드를 입력하세요 (연산자: &, |, !, \"\", ())")
    
    # 엑셀 파일인 경우 컬럼 선택 기능 추가
    selected_columns = None
    if uploaded_file and uploaded_file.name.endswith(('.xlsx', '.xls')):
        # 캐시된 함수를 사용하여 엑셀 파일 로드
        df = load_excel(uploaded_file)
        columns = get_columns(df)
        
        # 컬럼 선택 UI
        st.subheader("🔍 검색할 컬럼 선택")
        
        # 컬럼 선택을 위한 multiselect
        selected_columns = st.multiselect(
            "검색할 컬럼을 선택하세요 (여러 개 선택 가능)",
            options=columns,
            default=columns,
            key="column_selector"
        )
        
        # 선택된 컬럼이 없을 경우 경고
        if not selected_columns:
            st.warning("최소 하나 이상의 컬럼을 선택해주세요.")
        else:
            st.success(f"선택된 컬럼: {', '.join(selected_columns)}")
    
    # 검색 버튼 추가
    search_button = st.button("🔍 검색하기", type="primary")
    
    # 검색 버튼이 클릭되었을 때만 검색 실행
    if search_button and uploaded_file and query:
        try:
            if uploaded_file.name.endswith('.pdf'):
                df = process_pdf(uploaded_file, query)
            else:
                if not selected_columns:
                    st.warning("엑셀 파일의 경우 검색할 컬럼을 선택해주세요.")
                else:
                    df = process_excel(uploaded_file, query, selected_columns)
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
    elif search_button and not uploaded_file:
        st.warning("파일을 업로드해주세요.")
    elif search_button and not query:
        st.warning("검색어를 입력해주세요.")

if __name__ == "__main__":
    main() 
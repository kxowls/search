# 문서 키워드 검색 도구

PDF와 Excel 파일에서 키워드를 검색하고 결과를 시각화하는 웹 애플리케이션입니다.

## 주요 기능

- PDF 및 Excel 파일 업로드
- 키워드 기반 검색
- 검색 결과 표 형태로 표시
- 결과 Excel 파일로 다운로드

## 설치 방법

1. 필요한 패키지 설치:
```bash
pip install -r requirements.txt
```

2. 애플리케이션 실행:
```bash
streamlit run app.py
```

## 사용 방법

1. 웹 브라우저에서 애플리케이션 접속
2. PDF 또는 Excel 파일 업로드
3. 검색할 키워드 입력
4. 검색 결과 확인 및 Excel 파일로 다운로드

## 기술 스택

- Python 3.10+
- Streamlit
- pandas
- pdfplumber
- openpyxl 
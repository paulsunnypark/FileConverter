# MS Office to PDF Converter

## 개요

이 프로젝트는 MS Office 문서(.doc, .docx, .xls, .xlsx, .ppt, .pptx)를 PDF 파일로 변환하는 도구입니다. 사용자는 GUI(Tkinter 기반) 또는 API(FastAPI 기반)를 통해 파일을 변환할 수 있습니다.

## 기능

- **파일 선택**: GUI를 통해 단일 또는 복수 파일 선택 가능
- **저장 경로 지정**: 변환된 PDF 파일을 저장할 폴더 지정
- **문서 변환**: MS Office 문서를 PDF로 변환
- **변환 결과 표시**: 변환 성공/실패 목록 표시
- **운영 모드**: GUI 모드와 FastAPI 서버 모드 지원

## 설치 방법

1. **필요한 소프트웨어**:
   - Python 3.8 이상
   - Microsoft Office (Word, Excel, PowerPoint) 설치 필요

2. **의존성 설치**:
   ```bash
   pip install -r requirements.txt
   ```

## 사용법

### GUI 모드

1. GUI 실행:
   ```bash
   python gui.py
   ```
2. 파일 추가 버튼을 클릭하여 변환할 파일 선택
3. 출력 폴더 지정
4. 'PDF로 변환 시작' 버튼 클릭

### API 서버 모드

1. 서버 실행:
   ```bash
   uvicorn main:app --reload --host 127.0.0.1 --port 8000
   ```
2. 브라우저에서 `http://127.0.0.1:8000/docs` 접속하여 API 테스트
3. `/convert/` 엔드포인트로 파일 업로드 및 변환 요청

## 개발 과정

1. **환경 설정**: 필요한 라이브러리 설치 (`pywin32`, `fastapi`, `uvicorn` 등)
2. **핵심 변환 로직 구현**: `converter.py`에서 MS Office 문서를 PDF로 변환하는 기능 구현
3. **GUI 구현**: `gui.py`에서 Tkinter를 사용한 파일 선택 및 변환 인터페이스 구현
4. **API 서버 구현**: `main.py`에서 FastAPI를 사용한 웹 API 서버 구현
5. **오류 처리 및 개선**: 출력 폴더 생성, Excel 변환 오류 처리, 파일 선택 기본 설정 등 수정

## 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다.

## 기여

버그 제보, 기능 제안 등은 GitHub 이슈를 통해 가능합니다. 
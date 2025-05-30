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

## GitHub 등록 과정

1. **Git 초기화 및 로컬 저장소 생성**:
   - 터미널에서 프로젝트 디렉토리로 이동: `cd /d/projects/FileConverter`
   - Git 저장소 초기화: `git init`
   - 모든 파일 추가: `git add .`
   - 초기 커밋: `git commit -m "Initial commit for MS Office to PDF Converter"`

2. **GitHub에서 새 저장소 생성**:
   - GitHub 웹사이트에서 'New repository' 선택
   - 저장소 이름 입력 (예: `FileConverter`)
   - 공개(Public) 또는 비공개(Private) 설정 후 'Create repository' 클릭

3. **로컬 저장소를 GitHub에 연결 및 푸시**:
   - GitHub 저장소 URL 복사 (HTTPS 또는 SSH)
   - 로컬 저장소 연결: `git remote add origin https://github.com/사용자이름/FileConverter.git`
   - 커밋 푸시: `git push -u origin master`

4. **GitHub에서 확인**:
   - 저장소 페이지에서 파일 업로드 확인
   - `README.md` 파일이 메인 페이지에 표시되는지 확인

## 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다.

## 기여

버그 제보, 기능 제안 등은 GitHub 이슈를 통해 가능합니다. 

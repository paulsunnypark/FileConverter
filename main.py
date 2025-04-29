# main.py

# 구현 의도: FastAPI를 사용하여 웹 기반 문서 변환 API 제공
# 기능 요약: 파일 업로드 및 변환 요청 처리, 결과 반환

import os
import shutil
import logging
import uuid # 고유한 임시 디렉토리 생성용
from typing import List, Dict, Any # 타입 힌팅 강화

from fastapi import FastAPI, File, UploadFile, Form, HTTPException, BackgroundTasks
from fastapi.responses import JSONResponse

# converter 모듈 임포트 (같은 디렉토리에 있어야 함)
from converter import convert_to_pdf, SUPPORTED_EXTENSIONS

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- FastAPI 앱 초기화 ---
app = FastAPI(title="MS Office to PDF Converter API")

# --- 임시 파일 및 출력 디렉토리 설정 ---
# 💡 실제 운영 환경에서는 설정 파일이나 환경 변수 사용 고려
TEMP_UPLOAD_DIR = "temp_uploads"
DEFAULT_OUTPUT_DIR = "converted_pdfs"

# 서버 시작 시 임시 및 기본 출력 디렉토리 생성
os.makedirs(TEMP_UPLOAD_DIR, exist_ok=True)
os.makedirs(DEFAULT_OUTPUT_DIR, exist_ok=True)


# --- 백그라운드 작업 함수 ---
def cleanup_temp_dir(temp_dir: str):
    """임시 디렉토리와 그 안의 파일을 삭제합니다."""
    try:
        shutil.rmtree(temp_dir)
        logging.info(f"임시 디렉토리 삭제 완료: {temp_dir}")
    except OSError as e:
        logging.error(f"임시 디렉토리 삭제 실패: {temp_dir} - {e}")


# --- API 엔드포인트 ---
@app.post("/convert/", summary="Upload Office files and convert to PDF")
async def convert_files_endpoint(
    background_tasks: BackgroundTasks, # 백그라운드 작업 추가
    files: List[UploadFile] = File(..., description=f"변환할 Office 파일 목록 ({', '.join(SUPPORTED_EXTENSIONS.keys())})"),
    output_subdir: str = Form(None, description="결과 PDF를 저장할 하위 디렉토리명 (기본값: converted_pdfs)")
) -> JSONResponse:
    """
    Office 파일을 업로드 받아 PDF로 변환하고 결과를 반환합니다.

    - **files**: 변환할 MS Office 파일을 multipart/form-data 형식으로 전송합니다.
    - **output_subdir**: (선택 사항) PDF가 저장될 기본 출력 폴더 내의 하위 디렉토리 이름을 지정합니다.
                       지정하지 않으면 기본 출력 폴더에 저장됩니다.
    """
    results = []
    # 고유한 임시 작업 디렉토리 생성 (동시 요청 처리 시 충돌 방지)
    job_id = str(uuid.uuid4())
    temp_job_dir = os.path.join(TEMP_UPLOAD_DIR, job_id)
    os.makedirs(temp_job_dir, exist_ok=True)
    logging.info(f"임시 작업 디렉토리 생성: {temp_job_dir}")

    # 실제 출력 경로 결정
    if output_subdir:
        # 보안: 경로 조작 방지 (상위 디렉토리 이동 등 금지)
        # 간단하게 슬래시, 백슬래시, 점(..) 포함 여부 확인
        if any(c in output_subdir for c in ['/', '\\', '..']):
             # 임시 디렉토리 정리 예약
            background_tasks.add_task(cleanup_temp_dir, temp_job_dir)
            raise HTTPException(status_code=400, detail="출력 하위 디렉토리 이름에 유효하지 않은 문자가 포함되어 있습니다.")
        final_output_dir = os.path.join(DEFAULT_OUTPUT_DIR, output_subdir)
    else:
        final_output_dir = DEFAULT_OUTPUT_DIR

    # 필요한 경우 최종 출력 디렉토리 생성
    os.makedirs(final_output_dir, exist_ok=True)
    logging.info(f"최종 출력 디렉토리: {final_output_dir}")

    # 업로드된 파일 처리
    for file in files:
        original_filename = file.filename
        file_ext = os.path.splitext(original_filename)[1].lower()

        # 지원하는 확장자인지 확인
        if file_ext not in SUPPORTED_EXTENSIONS:
            results.append({
                "filename": original_filename,
                "success": False,
                "message": f"오류: 지원하지 않는 파일 형식 ({file_ext})"
            })
            logging.warning(f"지원하지 않는 파일 형식 업로드됨: {original_filename}")
            continue # 다음 파일로 넘어감

        # 파일을 임시 디렉토리에 저장
        temp_file_path = os.path.join(temp_job_dir, original_filename)
        try:
            with open(temp_file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            logging.info(f"임시 파일 저장 완료: {temp_file_path}")
        except Exception as e:
            results.append({
                "filename": original_filename,
                "success": False,
                "message": f"오류: 파일 저장 실패 - {e}"
            })
            logging.error(f"임시 파일 저장 실패: {original_filename} - {e}")
            continue # 다음 파일로
        finally:
            # UploadFile 객체의 file 포인터 닫기 (중요)
            await file.close()

        # PDF 변환 시도 (converter.py 사용)
        logging.info(f"PDF 변환 시작: {temp_file_path} -> {final_output_dir}")
        success, message = convert_to_pdf(temp_file_path, final_output_dir)

        # 결과 기록
        results.append({
            "filename": original_filename,
            "success": success,
            "message": message # 성공 시 PDF 경로, 실패 시 오류 메시지
        })

    # 변환 작업 완료 후 임시 디렉토리 정리 예약 (백그라운드에서 실행)
    # 💡 background_tasks.add_task는 응답 반환 후 비동기적으로 실행됨
    background_tasks.add_task(cleanup_temp_dir, temp_job_dir)

    return JSONResponse(content={"conversion_results": results})

@app.get("/", summary="API 기본 정보")
async def root():
    """API 서버가 실행 중인지 확인하는 기본 엔드포인트입니다."""
    return {"message": "MS Office to PDF Converter API가 실행 중입니다."}

@app.get("/favicon.ico", summary="Favicon 요청 처리")
async def favicon():
    """Favicon 요청에 대해 빈 응답을 반환하여 404 오류를 방지합니다."""
    return None

# --- 서버 실행 (uvicorn 사용) ---
# 터미널에서 실행: uvicorn main:app --reload
# 예: uvicorn main:app --host 0.0.0.0 --port 8000
# --reload 옵션은 개발 중 코드 변경 시 자동 재시작
if __name__ == "__main__":
    import uvicorn
    # host="0.0.0.0"으로 설정하면 외부에서도 접근 가능
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True) 
# converter.py

# 구현 의도: MS Office 문서를 PDF로 변환하는 핵심 로직 제공
# 기능 요약: 지정된 경로의 오피스 파일을 열어 PDF 형식으로 저장

import os
import sys
import comtypes.client
import logging

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 지원하는 파일 확장자 (Word, Excel, PowerPoint)
SUPPORTED_EXTENSIONS = {
    ".doc": "Word.Application",
    ".docx": "Word.Application",
    ".xls": "Excel.Application",
    ".xlsx": "Excel.Application",
    ".ppt": "PowerPoint.Application",
    ".pptx": "PowerPoint.Application",
}

# PDF 형식 코드 (Office 버전별 상이할 수 있음)
# Word: 17, Excel: 0, PowerPoint: 32
PDF_FORMAT_CODE = {
    "Word.Application": 17,
    "Excel.Application": 0, # xlTypePDF
    "PowerPoint.Application": 32 # ppSaveAsPDF
}

def convert_to_pdf(input_path, output_path):
  """
  MS Office 문서를 PDF로 변환합니다.

  Args:
      input_path (str): 변환할 원본 오피스 파일 경로.
      output_path (str): PDF 파일을 저장할 경로 (확장자 제외).

  Returns:
      tuple: (성공 여부 (bool), 결과 메시지 또는 에러 메시지 (str))
  """
  # 입력 파일 존재 여부 확인
  if not os.path.exists(input_path):
    return False, f"오류: 입력 파일 없음 - {input_path}"

  # 입력 파일 확장자 확인 및 지원 여부 검사
  file_name, file_ext = os.path.splitext(input_path)
  file_ext = file_ext.lower()

  if file_ext not in SUPPORTED_EXTENSIONS:
    return False, f"오류: 지원하지 않는 파일 형식 - {file_ext}"

  # 출력 디렉토리 존재 여부 확인 및 생성
  output_dir = output_path
  if output_dir == '':
    output_dir = os.path.dirname(input_path)  # 입력 파일 디렉토리를 기본으로 사용
  if not os.path.exists(output_dir):
    try:
      os.makedirs(output_dir)
      logging.info(f"출력 디렉토리 생성: {output_dir}")
    except OSError as e:
      return False, f"오류: 출력 디렉토리 생성 실패 - {e}"

  # PDF 파일 경로 생성 (원본 파일명 + .pdf)
  pdf_output_path = os.path.join(output_dir, os.path.basename(file_name) + ".pdf")

  app_name = SUPPORTED_EXTENSIONS[file_ext]
  pdf_format = PDF_FORMAT_CODE[app_name]

  app = None
  doc = None

  try:
    # COM 객체 생성 (MS Office 애플리케이션 실행)
    # 💡 comtypes.client.CreateObject는 백그라운드에서 Office 앱을 실행합니다.
    # 이미 실행 중인 인스턴스가 있다면 그것을 사용할 수도 있습니다 (GetActiveObject).
    app = comtypes.client.CreateObject(app_name)
    app.Visible = False # 백그라운드 실행

    logging.info(f"{app_name} 애플리케이션 시작: {input_path}")

    # 문서 열기
    # 💡 각 Office 앱마다 문서 열기 메서드와 인자가 다릅니다.
    if app_name == "Word.Application":
      doc = app.Documents.Open(os.path.abspath(input_path), ReadOnly=True)
      doc.SaveAs(os.path.abspath(pdf_output_path), FileFormat=pdf_format)
    elif app_name == "Excel.Application":
      # Workbooks.Open은 절대 경로 필요
      doc = app.Workbooks.Open(os.path.abspath(input_path), ReadOnly=True)
      # ExportAsFixedFormat 메서드 사용 (PDF 변환)
      # Quality=0 은 Standard 품질
      try:
        doc.ExportAsFixedFormat(0, os.path.abspath(pdf_output_path), Quality=0)
      except Exception as export_error:
        logging.error(f"Excel PDF 내보내기 오류: {input_path} - {export_error}")
        # 대체 방법: SaveAs 시도 (일부 Excel 버전에서 ExportAsFixedFormat이 실패할 수 있음)
        try:
          doc.SaveAs(os.path.abspath(pdf_output_path), FileFormat=57)  # 57은 PDF 형식
          logging.info(f"Excel SaveAs로 PDF 저장 성공: {pdf_output_path}")
        except Exception as saveas_error:
          logging.error(f"Excel SaveAs 실패: {input_path} - {saveas_error}")
          raise Exception(f"Excel PDF 변환 실패: ExportAsFixedFormat 및 SaveAs 모두 실패 - {export_error}; {saveas_error}")
    elif app_name == "PowerPoint.Application":
      doc = app.Presentations.Open(os.path.abspath(input_path), ReadOnly=True, WithWindow=False)
      doc.SaveAs(os.path.abspath(pdf_output_path), FileFormat=pdf_format)

    logging.info(f"PDF 변환 성공: {pdf_output_path}")
    return True, pdf_output_path

  except Exception as e:
    # 오류 발생 시 로그 기록 및 실패 반환
    # 💡 COM 객체 오류는 매우 다양하므로 포괄적인 예외 처리가 필요합니다.
    # 특정 오류 코드에 따른 세분화된 처리가 가능하면 더 좋습니다.
    logging.error(f"PDF 변환 실패: {input_path} - {e}")
    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    logging.error(f"오류 상세: {exc_type}, {fname}, {exc_tb.tb_lineno}")
    return False, f"오류: 변환 실패 ({input_path}) - {e}"

  finally:
    # 리소스 정리 (문서 닫기, 애플리케이션 종료)
    # 💡 finally 블록에서 확실하게 리소스를 해제해야 프로세스가 남지 않습니다.
    if doc:
      try:
        if app_name == "Word.Application":
            doc.Close(False) # 변경사항 저장 안함
        elif app_name == "Excel.Application":
            doc.Close(SaveChanges=False)
        elif app_name == "PowerPoint.Application":
            doc.Close()
        logging.info(f"문서 닫기 완료: {input_path}")
      except Exception as e:
        logging.warning(f"문서 닫기 중 오류: {input_path} - {e}")
    if app:
      try:
        # Quit 메서드는 애플리케이션별로 다를 수 있음
        app.Quit()
        logging.info(f"{app_name} 애플리케이션 종료")
      except Exception as e:
        logging.warning(f"{app_name} 종료 중 오류: {e}")

    # COM 객체 참조 해제 (메모리 누수 방지)
    # comtypes 사용 시 명시적 해제가 중요할 수 있습니다.
    app = None
    doc = None
    # comtypes.CoUninitialize() # 스레드 단위 초기화/해제 필요시 고려

# --- 테스트 코드 ---
if __name__ == "__main__":
  # 테스트용 입력 파일 경로 (실제 파일 경로로 수정 필요)
  test_input_doc = r"D:\path\to\your\test.docx"
  test_input_xls = r"D:\path\to\your\test.xlsx"
  test_input_ppt = r"D:\path\to\your\test.pptx"
  # 테스트용 출력 디렉토리 경로
  test_output_dir = r"D:\path\to\your\output"

  if not os.path.exists(test_output_dir):
      os.makedirs(test_output_dir)

  logging.info("--- Word 변환 테스트 시작 ---")
  success_doc, message_doc = convert_to_pdf(test_input_doc, test_output_dir)
  if success_doc:
    logging.info(f"Word 변환 성공: {message_doc}")
  else:
    logging.error(f"Word 변환 실패: {message_doc}")

  logging.info("--- Excel 변환 테스트 시작 ---")
  success_xls, message_xls = convert_to_pdf(test_input_xls, test_output_dir)
  if success_xls:
    logging.info(f"Excel 변환 성공: {message_xls}")
  else:
    logging.error(f"Excel 변환 실패: {message_xls}")

  logging.info("--- PowerPoint 변환 테스트 시작 ---")
  success_ppt, message_ppt = convert_to_pdf(test_input_ppt, test_output_dir)
  if success_ppt:
    logging.info(f"PowerPoint 변환 성공: {message_ppt}")
  else:
    logging.error(f"PowerPoint 변환 실패: {message_ppt}")
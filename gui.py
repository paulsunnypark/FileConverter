# gui.py

# 구현 의도: 사용자 친화적인 GUI를 통해 문서 변환 기능을 제공
# 기능 요약: 파일/폴더 선택, 변환 실행, 결과 표시

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import os
import threading
import logging
from converter import convert_to_pdf, SUPPORTED_EXTENSIONS

# 로깅 설정 (GUI에서는 파일 핸들러 추가 등 고려 가능)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class PdfConverterApp:
  """
  Tkinter 기반 PDF 변환 GUI 애플리케이션 클래스
  """
  def __init__(self, master):
    self.master = master
    master.title("MS Office to PDF Converter")
    master.geometry("700x600") # 창 크기 조정

    # --- 변수 초기화 ---
    self.input_files = []
    self.output_dir = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Documents", "ConvertedPDFs")) # 기본 출력 경로

    # --- 위젯 생성 ---
    self._create_widgets()

  def _create_widgets(self):
    """GUI 위젯들을 생성하고 배치합니다."""

    # --- 프레임 생성 ---
    input_frame = ttk.LabelFrame(self.master, text="입력 파일 선택", padding=(10, 5))
    input_frame.pack(padx=10, pady=5, fill="x")

    output_frame = ttk.LabelFrame(self.master, text="출력 폴더 지정", padding=(10, 5))
    output_frame.pack(padx=10, pady=5, fill="x")

    action_frame = ttk.Frame(self.master, padding=(10, 5))
    action_frame.pack(padx=10, pady=10, fill="x")

    result_frame = ttk.LabelFrame(self.master, text="변환 결과", padding=(10, 5))
    result_frame.pack(padx=10, pady=5, fill="both", expand=True)

    # --- 입력 파일 선택 ---
    self.file_listbox = tk.Listbox(input_frame, selectmode=tk.EXTENDED, height=8) # 높이 조정
    self.file_listbox.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))

    scrollbar = ttk.Scrollbar(input_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
    scrollbar.pack(side=tk.RIGHT, fill="y")
    self.file_listbox.config(yscrollcommand=scrollbar.set)

    file_button_frame = ttk.Frame(input_frame)
    file_button_frame.pack(side=tk.LEFT, padx=(5, 0))

    add_button = ttk.Button(file_button_frame, text="파일 추가", command=self.select_files)
    add_button.pack(pady=2, fill="x")

    remove_button = ttk.Button(file_button_frame, text="선택 제거", command=self.remove_selected_files)
    remove_button.pack(pady=2, fill="x")

    clear_button = ttk.Button(file_button_frame, text="목록 비우기", command=self.clear_file_list)
    clear_button.pack(pady=2, fill="x")


    # --- 출력 폴더 지정 ---
    output_entry = ttk.Entry(output_frame, textvariable=self.output_dir, width=60) # 너비 조정
    output_entry.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))

    browse_button = ttk.Button(output_frame, text="폴더 찾기", command=self.select_output_dir)
    browse_button.pack(side=tk.LEFT)

    # --- 변환 실행 ---
    self.progress_bar = ttk.Progressbar(action_frame, orient="horizontal", length=300, mode="determinate")
    self.progress_bar.pack(side=tk.LEFT, padx=(0, 10), expand=True, fill="x")

    self.convert_button = ttk.Button(action_frame, text="PDF로 변환 시작", command=self.start_conversion_thread)
    self.convert_button.pack(side=tk.LEFT)

    # --- 결과 표시 ---
    self.result_text = scrolledtext.ScrolledText(result_frame, height=15, state=tk.DISABLED) # 읽기 전용
    self.result_text.pack(fill="both", expand=True)

  # --- 이벤트 핸들러 및 로직 함수 ---

  def select_files(self):
    """파일 선택 대화상자를 열어 변환할 파일을 목록에 추가합니다."""
    # 💡 지원하는 확장자 목록을 생성하여 filedialog에 적용합니다.
    supported_types = []
    supported_types.append(("모든 지원 파일", " ".join([f"*{ext}" for ext in SUPPORTED_EXTENSIONS.keys()])))
    for ext in SUPPORTED_EXTENSIONS.keys():
        # description = ext.split('.')[1].upper() + " 파일"
        supported_types.append((f"{ext.split('.')[1].upper()} 파일", f"*{ext}"))
    supported_types.append(("모든 파일", "*.*"))


    # 💡 askopenfilenames는 여러 파일 경로를 튜플로 반환합니다.
    selected_files = filedialog.askopenfilenames(
      title="변환할 파일 선택",
      filetypes=supported_types
    )
    if selected_files:
      for file_path in selected_files:
        if file_path not in self.input_files:
          self.input_files.append(file_path)
          self.file_listbox.insert(tk.END, os.path.basename(file_path)) # 리스트박스에는 파일명만 표시
      logging.info(f"{len(selected_files)}개 파일 추가됨.")

  def remove_selected_files(self):
      """리스트박스에서 선택된 항목들을 제거합니다."""
      selected_indices = self.file_listbox.curselection()
      if not selected_indices:
          messagebox.showwarning("선택 없음", "제거할 파일을 목록에서 선택해주세요.")
          return

      # 💡 뒤에서부터 삭제해야 인덱스 오류 방지
      for i in reversed(selected_indices):
          file_name_to_remove = self.file_listbox.get(i)
          # 실제 input_files 리스트에서도 해당 경로를 찾아 제거해야 함
          # 경로 기반으로 찾는 것이 더 정확
          original_path_to_remove = None
          for path in self.input_files:
              if os.path.basename(path) == file_name_to_remove:
                  # 동일 파일명이 여러 디렉토리에 있을 수 있으므로,
                  # 실제로는 더 정확한 매칭 로직이 필요할 수 있음 (예: 전체 경로 저장/비교)
                  # 여기서는 단순화를 위해 첫 번째 매칭 사용
                  original_path_to_remove = path
                  break
          if original_path_to_remove:
              self.input_files.remove(original_path_to_remove)
          self.file_listbox.delete(i)
      logging.info(f"{len(selected_indices)}개 파일 제거됨.")


  def clear_file_list(self):
      """파일 목록 전체를 비웁니다."""
      self.input_files.clear()
      self.file_listbox.delete(0, tk.END)
      logging.info("파일 목록 비워짐.")


  def select_output_dir(self):
    """출력 폴더 선택 대화상자를 열어 저장 경로를 설정합니다."""
    # 💡 askdirectory는 선택된 폴더 경로를 문자열로 반환합니다.
    directory = filedialog.askdirectory(title="PDF 저장 폴더 선택")
    if directory:
      self.output_dir.set(directory)
      logging.info(f"출력 폴더 변경: {directory}")

  def update_result_text(self, message):
      """결과 텍스트 영역을 안전하게 업데이트합니다."""
      self.result_text.config(state=tk.NORMAL) # 쓰기 가능 상태로 변경
      self.result_text.insert(tk.END, message + "\n")
      self.result_text.see(tk.END) # 스크롤을 맨 아래로 이동
      self.result_text.config(state=tk.DISABLED) # 다시 읽기 전용으로

  def set_progress(self, value):
      """진행률 표시줄 값을 업데이트합니다."""
      self.progress_bar['value'] = value


  def start_conversion_thread(self):
      """변환 작업을 별도 스레드에서 시작합니다."""
      if not self.input_files:
          messagebox.showwarning("입력 없음", "변환할 파일을 먼저 추가해주세요.")
          return

      output_dir = self.output_dir.get()
      if not output_dir:
          messagebox.showwarning("출력 없음", "PDF를 저장할 폴더를 지정해주세요.")
          return

      # 출력 폴더가 존재하지 않으면 생성
      if not os.path.exists(output_dir):
          try:
              os.makedirs(output_dir)
              logging.info(f"출력 폴더 생성: {output_dir}")
              self.update_result_text(f"출력 폴더 생성: {output_dir}")
          except OSError as e:
              messagebox.showerror("폴더 생성 오류", f"출력 폴더를 생성할 수 없습니다: {e}")
              return

      # GUI 비활성화 (중복 실행 방지)
      self.convert_button.config(state=tk.DISABLED)
      # 결과 텍스트 초기화
      self.result_text.config(state=tk.NORMAL)
      self.result_text.delete('1.0', tk.END)
      self.result_text.config(state=tk.DISABLED)
      self.update_result_text("--- 변환 시작 ---")

      # 💡 변환 작업은 시간이 오래 걸릴 수 있으므로 별도 스레드에서 실행
      # self를 넘겨주어 스레드에서 GUI 요소에 접근 가능하게 함 (주의 필요)
      conversion_thread = threading.Thread(target=self._run_conversion, args=(self.input_files.copy(), output_dir), daemon=True)
      conversion_thread.start()


  def _run_conversion(self, files_to_convert, output_directory):
      """실제 파일 변환 로직을 실행하는 스레드 함수."""
      total_files = len(files_to_convert)
      success_count = 0
      fail_count = 0
      self.set_progress(0)

      for i, input_path in enumerate(files_to_convert):
          base_name = os.path.basename(input_path)
          self.update_result_text(f"[{i+1}/{total_files}] '{base_name}' 변환 중...")
          logging.info(f"변환 시도: {input_path}")

          # 💡 converter 모듈의 함수 호출
          success, message = convert_to_pdf(input_path, output_directory)

          if success:
              success_count += 1
              self.update_result_text(f"  -> 성공: {message}")
              logging.info(f"변환 성공: {input_path} -> {message}")
          else:
              fail_count += 1
              self.update_result_text(f"  -> 실패: {message}")
              logging.error(f"변환 실패: {input_path} - {message}")

          # 진행률 업데이트
          progress_value = (i + 1) / total_files * 100
          self.set_progress(progress_value)

      # 변환 완료 후 메시지 표시 및 GUI 활성화
      completion_message = f"--- 변환 완료 --- 총 {total_files}개 중 {success_count}개 성공, {fail_count}개 실패"
      self.update_result_text(completion_message)
      logging.info(completion_message)
      messagebox.showinfo("변환 완료", completion_message)

      # 💡 GUI 요소 업데이트는 메인 스레드에서 수행하는 것이 가장 안전합니다.
      # Tkinter는 기본적으로 스레드 안전하지 않으므로, after() 메서드나 Queue 사용 고려
      # 여기서는 단순화를 위해 직접 호출하지만, 복잡한 앱에서는 문제 발생 가능성 있음
      self.master.after(0, self._enable_convert_button) # 메인 스레드에서 버튼 활성화

  def _enable_convert_button(self):
      """변환 버튼을 다시 활성화합니다."""
      self.convert_button.config(state=tk.NORMAL)


# --- 애플리케이션 실행 ---
if __name__ == "__main__":
  root = tk.Tk()
  app = PdfConverterApp(root)
  root.mainloop() 
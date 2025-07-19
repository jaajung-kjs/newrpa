import subprocess
import time
import pygetwindow as gw
import pyautogui
import os
from pywinauto import Desktop
from datetime import datetime
from tkinter import messagebox

# docxrpa.py에서 필요한 함수 임포트
from docxrpa import create_base_report_document, insert_image_to_document_table

year = datetime.now().year
month = datetime.now().month
day = datetime.now().day

def open_cmd_and_run_systeminfo():
    subprocess.Popen('start cmd', shell=True)
    time.sleep(2)
    # CMD 창을 열고 systeminfo 명령어 실행 후 창을 최대화하여 유지
    subprocess.Popen('start /max cmd /k "systeminfo | findstr /I /C:"OS 이름" /C:"OS 버전" /C:"원래 설치 날짜""', shell=True)

def open_screensaver_settings():
    # 화면 보호기 설정 창 열기
    subprocess.Popen('control desk.cpl,,@screensaver', shell=True)

def open_mac_address():
    subprocess.Popen('start /max cmd /k "ipconfig /all | more"', shell=True)

def run_and_capture_app(gui_instance=None):
    # CMD로 애플리케이션 실행
    subprocess.Popen(r'C:\Program Files\AhnLab\V3Lite40\v3lite4.exe', shell=True)
    
    # 잠시 대기하여 애플리케이션 창이 열리도록 함
    time.sleep(1)
    
    # 열린 창 목록 가져오기
    windows = gw.getWindowsWithTitle('AhnLab V3 Lite')

    if windows:
        window = windows[0]
        left, top, right, bottom = window.left, window.top, window.right, window.bottom
        width, height = right - left, bottom - top
        if gui_instance: gui_instance.hide_overlay()
        screenshot = pyautogui.screenshot(region=(left, top, width, height))
        if gui_instance: gui_instance.show_overlay()
        return screenshot # 이미지 객체를 반환하도록 수정

    else:
        print("애플리케이션 창을 찾을 수 없습니다.")
        return None

# 특정 창 제목을 기준으로 창을 닫는 함수
def close_window_by_title(target_title):
    try:
        # Desktop 객체를 생성하여 현재 열려 있는 모든 창을 검색
        desktop = Desktop(backend="uia")
        window = desktop.window(title=target_title)

        if window.exists():
            window.close()
            print(f"Closed window: {target_title}")
        else:
            print(f"Window not found: {target_title}")

    except Exception as e:
        print(f"Error: {e}")

def v3log():
    # CMD로 애플리케이션 실행
    subprocess.Popen(r'C:\Program Files\AhnLab\V3Lite40\v3lite4.exe', shell=True)

    # 잠시 대기하여 애플리케이션 창이 열리도록 함
    time.sleep(1)

    # Tab 키 9번 누르기
    for _ in range(9):
        pyautogui.press('tab')
        time.sleep(0.2)  # 키 입력 사이에 약간의 지연 추가

    pyautogui.press('space')
    time.sleep(0.2)

    pyautogui.press('right')
    pyautogui.press('right')

    pyautogui.press('tab')

    pyautogui.write(str(year-1), interval=0.1)
    pyautogui.press('right')
    pyautogui.write(str(month), interval=0.1)
    pyautogui.press('right')
    pyautogui.write(str(day), interval=0.1)

    for _ in range(4):
        pyautogui.press('tab')
        time.sleep(0.2)

    pyautogui.press('space')

def appwiz():
    # 제어판의 "프로그램 추가/제거" 패널 열기
    os.system('control appwiz.cpl')

    # 잠시 대기하여 제어판이 열리도록 함
    time.sleep(3)

    # Alt + Space, X 키 입력으로 창 최대화
    pyautogui.hotkey('alt', 'space')
    time.sleep(0.5)
    pyautogui.press('x')

    # 잠시 대기하여 창 최대화 적용
    time.sleep(1)

    # Tab 키 8번 누르기
    for _ in range(8):
        pyautogui.press('tab')
        time.sleep(0.2)  # 키 입력 사이에 약간의 지연 추가

    # 오른쪽 방향키 1번 누르기
    pyautogui.press('right')

    time.sleep(0.1)

    # 아래쪽 방향키 5번 누르기
    for _ in range(5):
        pyautogui.press('down')
        time.sleep(0.1)

    time.sleep(0.1)

    pyautogui.press('enter')

def v3_check():
    # CMD로 애플리케이션 실행
    subprocess.Popen(r'C:\Program Files\AhnLab\V3Lite40\v3lite4.exe', shell=True)

    # 잠시 대기
    time.sleep(1)

    # Tab 키 10번 누르기
    for _ in range(10):
        pyautogui.press('tab')
        time.sleep(0.2)  # 키 입력 사이에 약간의 지연 추가

    pyautogui.press('space')


def capture_cmd_window(gui_instance=None):
    # 일정 시간 대기 후 CMD 창 캡쳐
    time.sleep(5)  # CMD 창이 열리고 systeminfo 명령어가 실행될 시간을 대기
    if gui_instance: gui_instance.hide_overlay()
    screenshot = pyautogui.screenshot()
    if gui_instance: gui_instance.show_overlay()
    return screenshot

def v3_check_capture(gui_instance=None):

    time.sleep(5)  
     # 열린 창 목록 가져오기
    windows = gw.getWindowsWithTitle('AhnLab V3 Lite')

    if windows:
        window = windows[0]
        left, top, right, bottom = window.left, window.top, window.right, window.bottom
        width, height = right - left, bottom - top
        if gui_instance: gui_instance.hide_overlay()
        screenshot = pyautogui.screenshot(region=(left, top, width, height))
        if gui_instance: gui_instance.show_overlay()
        return screenshot
    
def v3_log_capture(gui_instance=None):

    time.sleep(5)  
     # 열린 창 목록 가져오기
    windows = gw.getWindowsWithTitle('AhnLab V3 Lite')

    if windows:
        window = windows[0]
        left, top, right, bottom = window.left, window.top, window.right, window.bottom
        width, height = right - left, bottom - top
        if gui_instance: gui_instance.hide_overlay()
        screenshot = pyautogui.screenshot(region=(left, top, width, height))
        if gui_instance: gui_instance.show_overlay()
        return screenshot

def capture_screensaver_settings_window(gui_instance=None):
    time.sleep(2)  # 화면 보호기 설정 창이 열릴 시간을 대기
    windows = gw.getWindowsWithTitle('화면 보호기 설정')  # 창의 제목에 맞게 조정
    if windows:
        window = windows[0]
        left, top, right, bottom = window.left, window.top, window.right, window.bottom
        width, height = right - left, bottom - top
        if gui_instance: gui_instance.hide_overlay()
        screenshot = pyautogui.screenshot(region=(left, top, width, height))
        if gui_instance: gui_instance.show_overlay()
        return screenshot
    else:
        print("화면 보호기 설정 창을 찾을 수 없습니다.")
        return None

def show_desktop():
    # 잠시 대기하여 사용자가 작업을 완료할 수 있도록 함
    time.sleep(1)
    
    # 윈도우 키 누르기
    pyautogui.keyDown('win')
    
    # D 키 누르기
    pyautogui.press('d')
    
    # 윈도우 키 떼기
    pyautogui.keyUp('win')


import threading
from tkinter import Tk, Button, Label, Toplevel, messagebox

def run_rpa_tasks(update_progress_callback, gui_instance):
    # 새 문서 생성
    document = create_base_report_document()
    output_filename = f"노트북_점검_보고서_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"

    tasks = [
        {"name": "Systeminfo 캡쳐", "action": lambda: (open_cmd_and_run_systeminfo(), capture_cmd_window(gui_instance)), "doc_row": 3, "doc_col": 0},
        {"name": "화면 보호기 설정 캡쳐", "action": lambda: (open_screensaver_settings(), capture_screensaver_settings_window(gui_instance)), "doc_row": 7, "doc_col": 0},
        {"name": "MAC 주소 캡쳐", "action": lambda: (open_mac_address(), capture_cmd_window(gui_instance)), "doc_row": 3, "doc_col": 1},
        {"name": "V3 Lite 실행 및 캡쳐", "action": lambda: run_and_capture_app(gui_instance), "doc_row": 9, "doc_col": 1},
        {"name": "프로그램 추가/제거 캡쳐", "action": lambda: (appwiz(), capture_cmd_window(gui_instance)), "doc_row": 9, "doc_col": 0},
        {"name": "AhnLab V3 Lite 창 닫기 (1차)", "action": lambda: close_window_by_title("AhnLab V3 Lite")},
        {"name": "V3 검사 실행 및 캡쳐", "action": lambda: (v3_check(), v3_check_capture(gui_instance)), "doc_row": 13, "doc_col": 0},
        {"name": "AhnLab V3 Lite 창 닫기 (2차)", "action": lambda: close_window_by_title("AhnLab V3 Lite")},
        {"name": "V3 로그 캡쳐", "action": lambda: (v3log(), v3_log_capture(gui_instance)), "doc_row": 17, "doc_col": 0},
        {"name": "바탕화면 캡쳐", "action": lambda: (show_desktop(), capture_cmd_window(gui_instance)), "doc_row": 7, "doc_col": 1}
    ]

    total_tasks = len(tasks)
    gui_instance.task_results = [] # Initialize task results

    for i, task in enumerate(tasks):
        step_num = i + 1
        task_name = task["name"]
        message = f"({step_num}/{total_tasks}) {task_name} 진행 중..."
        update_progress_callback(message)
        
        result = {"name": task_name, "status": "실패", "error": None}

        try:
            action_output = task["action"]()

            image_object = None
            if isinstance(action_output, tuple):
                # 튜플의 두 번째 요소가 이미지 객체라고 가정
                image_object = action_output[1]
            else:
                # 직접 이미지 객체를 반환하는 경우
                image_object = action_output

            if image_object and task.get("doc_row") is not None and task.get("doc_col") is not None:
                print(f"{task_name} 스크린샷 캡처 완료.")
                insert_image_to_document_table(document, image_object, task["doc_row"], task["doc_col"])
            
            result["status"] = "성공"

        except Exception as e:
            result["error"] = str(e)
            print(f"Error during {task_name}: {e}")
        finally:
            gui_instance.task_results.append(result)
            time.sleep(1) # Small delay for visual feedback
            
    # 모든 작업 완료 후 문서 저장
    document.save(output_filename)
    print(f"최종 보고서 저장: {output_filename}")

    update_progress_callback(f"모든 작업 완료! 보고서가 '{output_filename}'(으)로 저장되었습니다.")
    # The final messagebox will be handled by RPA_GUI after the thread finishes

class RPA_GUI:
    def __init__(self, master):
        self.master = master
        master.title("RPA 자동 점검")

        self.start_button = Button(master, text="RPA 시작", command=self.start_rpa)
        self.start_button.pack(pady=20)

        self.overlay = None
        self.progress_label = None
        self.task_results = [] # To store results

    def start_rpa(self):
        self.start_button.config(state="disabled")  # 버튼 비활성화
        self.show_overlay()
        
        # RPA 작업을 별도의 스레드에서 실행
        rpa_thread = threading.Thread(target=run_rpa_tasks, args=(self.update_progress, self)) # document 인자 추가
        rpa_thread.start()
        
        # 스레드 완료를 주기적으로 확인
        self.check_rpa_thread(rpa_thread)

    def show_overlay(self):
        self.overlay = Toplevel(self.master)
        self.overlay.wm_overrideredirect(True)  # 윈도우 테두리 제거
        self.overlay.wm_attributes("-topmost", True)  # 항상 위에
        self.overlay.wm_attributes("-alpha", 0.7)  # 투명도 설정

        # 화면 중앙에 위치
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        overlay_width = 800
        overlay_height = 100
        x = (screen_width // 2) - (overlay_width // 2)
        y = (screen_height // 2) - (overlay_height // 2)
        self.overlay.geometry(f"{overlay_width}x{overlay_height}+{x}+{y}")

        self.progress_label = Label(self.overlay, text="다음 작업 시작 중...", font=("Helvetica", 16), bg="lightgray")
        self.progress_label.pack(expand=True, fill="both")

    def update_progress(self, message):
        if self.progress_label:
            self.master.after(0, lambda: self.progress_label.config(text=message + " (조작하지 마세요)"))

    def check_rpa_thread(self, rpa_thread):
        if rpa_thread.is_alive():
            self.master.after(100, self.check_rpa_thread, rpa_thread)
        else:
            self.hide_overlay()
            self.start_button.config(state="normal")
            self.show_results_popup() # Show results after thread finishes

    def show_results_popup(self):
        success_count = sum(1 for r in self.task_results if r["status"] == "성공")
        fail_count = len(self.task_results) - success_count
        
        summary_message = f"""RPA 자동 점검 완료!

총 {len(self.task_results)}개 작업 중:
  성공: {success_count}개
  실패: {fail_count}개

"""
        
        if fail_count > 0:
            summary_message += """실패한 작업 목록:
"""
            for r in self.task_results:
                if r["status"] == "실패":
                    summary_message += f"""- {r['name']}: {r['error'] if r['error'] else '알 수 없는 오류'}
"""
        
        messagebox.showinfo("RPA 결과", summary_message)
        self.master.destroy() # Close the main window after showing results

    def hide_overlay(self):
        if self.overlay:
            self.overlay.destroy()
            self.overlay = None
            self.progress_label = None

    

if __name__ == "__main__":
    root = Tk()
    gui = RPA_GUI(root)
    root.mainloop()
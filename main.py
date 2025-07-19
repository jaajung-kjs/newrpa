import subprocess
import time
import pygetwindow as gw
import pyautogui
from datetime import datetime
from tkinter import Tk, Button, Label, Toplevel, messagebox
import threading

# docxrpa.py에서 필요한 함수 임포트
from docxrpa import create_base_report_document, insert_image_to_document_table

def capture_systeminfo_and_close(gui_instance=None):
    """Systeminfo 실행, 캡쳐, 창 닫기를 한 번에 처리"""
    # CMD 창 열고 systeminfo 실행
    subprocess.Popen('C:\\Windows\\System32\\cmd.exe /c start /max C:\\Windows\\System32\\cmd.exe /k systeminfo', shell=False)
    time.sleep(5)  # systeminfo 로딩 대기
    
    # 스크린샷 캡쳐
    if gui_instance: gui_instance.hide_overlay()
    screenshot = pyautogui.screenshot()
    if gui_instance: gui_instance.show_overlay()
    
    # CMD 창 닫기
    windows = gw.getWindowsWithTitle('C:\\Windows\\System32\\cmd.exe')
    for window in windows:
        try:
            window.close()
        except:
            pass
    
    print("Systeminfo 캡쳐 완료")
    return screenshot

def capture_mac_address_and_close(gui_instance=None):
    """MAC 주소 확인, 캡쳐, 창 닫기를 한 번에 처리"""
    subprocess.Popen('C:\\Windows\\System32\\cmd.exe /c start /max C:\\Windows\\System32\\cmd.exe /k "ipconfig /all | more"', shell=False)
    time.sleep(3)
    
    if gui_instance: gui_instance.hide_overlay()
    screenshot = pyautogui.screenshot()
    if gui_instance: gui_instance.show_overlay()
    
    windows = gw.getWindowsWithTitle('C:\\Windows\\System32\\cmd.exe')
    for window in windows:
        try:
            window.close()
        except:
            pass
    
    print("MAC 주소 캡쳐 완료")
    return screenshot

def capture_screensaver_and_close(gui_instance=None):
    """화면 보호기 설정 열기, 캡쳐, 창 닫기를 한 번에 처리"""
    subprocess.Popen(['C:\\Windows\\System32\\cmd.exe', '/c', 'control', 'desk.cpl,,@screensaver'])
    time.sleep(2)
    
    # 한글 또는 영문 창 찾기
    windows = gw.getWindowsWithTitle('화면 보호기 설정') or gw.getWindowsWithTitle('Screen Saver Settings')
    
    if windows:
        window = windows[0]
        window.activate()
        time.sleep(0.5)
        
        left, top, right, bottom = window.left, window.top, window.right, window.bottom
        width, height = right - left, bottom - top
        
        if gui_instance: gui_instance.hide_overlay()
        screenshot = pyautogui.screenshot(region=(left, top, width, height))
        if gui_instance: gui_instance.show_overlay()
        
        window.close()
        print("화면 보호기 설정 캡쳐 완료")
        return screenshot
    else:
        print("화면 보호기 설정 창을 찾을 수 없습니다.")
        return None

def v3_lite_capture_and_close(gui_instance=None, mode="main"):
    """V3 Lite 실행, 특정 작업 수행, 캡쳐, 창 닫기를 한 번에 처리"""
    # V3 실행
    subprocess.Popen(['C:\\Windows\\System32\\cmd.exe', '/c', r'C:\Program Files\AhnLab\V3Lite40\v3lite4.exe'])
    time.sleep(3)
    
    # 모드별 작업 수행
    if mode == "check":
        # 검사 실행
        for _ in range(10):
            pyautogui.press('tab')
            time.sleep(0.1)
        pyautogui.press('space')
        time.sleep(5)  # 검사 시작 대기
        
    elif mode == "log":
        # 로그 보기
        year = datetime.now().year
        month = datetime.now().month
        day = datetime.now().day
        
        for _ in range(9):
            pyautogui.press('tab')
            time.sleep(0.1)
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
            time.sleep(0.1)
        pyautogui.press('space')
        time.sleep(3)
    
    # 창 찾기 및 캡쳐
    windows = gw.getWindowsWithTitle('AhnLab V3 Lite')
    if windows:
        window = windows[0]
        window.activate()
        time.sleep(0.5)
        
        left, top, right, bottom = window.left, window.top, window.right, window.bottom
        width, height = right - left, bottom - top
        
        if gui_instance: gui_instance.hide_overlay()
        screenshot = pyautogui.screenshot(region=(left, top, width, height))
        if gui_instance: gui_instance.show_overlay()
        
        # V3 창이 닫힐 때까지 Alt+F4 반복
        print("V3 종료 시도 중...")
        close_attempts = 0
        max_attempts = 10
        
        while close_attempts < max_attempts:
            # V3 창이 있는지 확인
            v3_windows = gw.getWindowsWithTitle('AhnLab V3 Lite')
            if not v3_windows:
                print("V3 창이 완전히 닫혔습니다.")
                break
            
            # V3 창을 정확히 활성화하고 Alt+F4
            try:
                v3_window = v3_windows[0]
                v3_window.activate()
                time.sleep(0.3)  # 창이 활성화될 시간 대기
                
                # 활성 창이 V3인지 다시 확인
                active_window = gw.getActiveWindow()
                if active_window and 'AhnLab V3 Lite' in active_window.title:
                    pyautogui.hotkey('alt', 'f4')
                    time.sleep(0.5)
                    
                    # 종료 확인 대화상자가 있을 수 있으므로 Enter
                    pyautogui.press('enter')
                    time.sleep(0.5)
                else:
                    print("V3 창 활성화 실패, 재시도...")
            except Exception as e:
                print(f"V3 창 닫기 오류: {e}")
            
            close_attempts += 1
        
        if close_attempts >= max_attempts:
            print(f"V3 창 닫기 시도 {max_attempts}회 초과")
        
        print(f"V3 Lite {mode} 캡쳐 완료 및 창 닫기")
        return screenshot
    else:
        print("V3 Lite 창을 찾을 수 없습니다.")
        return None

def capture_desktop(gui_instance=None):
    """바탕화면 보기 및 캡쳐"""
    pyautogui.hotkey('win', 'd')
    time.sleep(1)
    
    if gui_instance: gui_instance.hide_overlay()
    screenshot = pyautogui.screenshot()
    if gui_instance: gui_instance.show_overlay()
    
    print("바탕화면 캡쳐 완료")
    return screenshot

def capture_programs_list(gui_instance=None):
    """프로그램 추가/제거 열기, 캡쳐"""
    subprocess.Popen(['C:\\Windows\\System32\\cmd.exe', '/c', 'control', 'appwiz.cpl'])
    time.sleep(3)
    
    # 창 최대화
    pyautogui.hotkey('alt', 'space')
    time.sleep(0.3)
    pyautogui.press('x')
    time.sleep(1)
    
    # 설치 날짜순 정렬
    for _ in range(8):
        pyautogui.press('tab')
        time.sleep(0.1)
    pyautogui.press('right')
    time.sleep(0.1)
    
    for _ in range(5):
        pyautogui.press('down')
        time.sleep(0.1)
    pyautogui.press('enter')
    time.sleep(2)
    
    if gui_instance: gui_instance.hide_overlay()
    screenshot = pyautogui.screenshot()
    if gui_instance: gui_instance.show_overlay()
    
    print("프로그램 추가/제거 캡쳐 완료")
    return screenshot

def run_rpa_tasks(update_progress_callback, gui_instance):
    """RPA 작업 실행"""
    captured_images = {}
    output_filename = f"노트북_점검_보고서_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    
    # 작업 목록 - 통합된 함수들 사용
    tasks = [
        {"name": "Systeminfo 캡쳐", "func": capture_systeminfo_and_close, "row": 3, "col": 0},
        {"name": "MAC 주소 캡쳐", "func": capture_mac_address_and_close, "row": 3, "col": 1},
        {"name": "화면 보호기 설정 캡쳐", "func": capture_screensaver_and_close, "row": 7, "col": 0},
        {"name": "V3 Lite 실행 및 캡쳐", "func": lambda g: v3_lite_capture_and_close(g, "main"), "row": 9, "col": 1},
        {"name": "V3 검사 실행 및 캡쳐", "func": lambda g: v3_lite_capture_and_close(g, "check"), "row": 13, "col": 0},
        {"name": "V3 로그 캡쳐", "func": lambda g: v3_lite_capture_and_close(g, "log"), "row": 17, "col": 0},
        {"name": "바탕화면 캡쳐", "func": capture_desktop, "row": 7, "col": 1},
        {"name": "프로그램 추가/제거 캡쳐", "func": capture_programs_list, "row": 9, "col": 0}
    ]
    
    total_tasks = len(tasks)
    gui_instance.task_results = []
    
    # 각 작업 실행
    for i, task in enumerate(tasks):
        step_num = i + 1
        task_name = task["name"]
        update_progress_callback(f"({step_num}/{total_tasks}) {task_name} 진행 중...")
        
        result = {"name": task_name, "status": "실패", "error": None}
        
        try:
            screenshot = task["func"](gui_instance)
            if screenshot:
                captured_images[(task["row"], task["col"])] = screenshot
                result["status"] = "성공"
        except Exception as e:
            result["error"] = str(e)
            print(f"Error during {task_name}: {e}")
        
        gui_instance.task_results.append(result)
        time.sleep(0.5)
    
    # 문서 생성 및 저장
    update_progress_callback("문서 생성 중...")
    print("문서 생성 시작...")
    start_time = time.time()
    
    document = create_base_report_document()
    for (row, col), image in captured_images.items():
        insert_image_to_document_table(document, image, row, col)
    
    document.save(output_filename)
    print(f"문서 생성 완료. 소요 시간: {time.time() - start_time:.2f}초")
    
    update_progress_callback(f"모든 작업 완료! 보고서가 '{output_filename}'(으)로 저장되었습니다.")

class RPA_GUI:
    def __init__(self, master):
        self.master = master
        master.title("RPA 자동 점검")
        
        self.start_button = Button(master, text="RPA 시작", command=self.start_rpa)
        self.start_button.pack(pady=20)
        
        self.overlay = None
        self.progress_label = None
        self.task_results = []
    
    def start_rpa(self):
        self.start_button.config(state="disabled")
        self.show_overlay()
        
        rpa_thread = threading.Thread(target=run_rpa_tasks, args=(self.update_progress, self))
        rpa_thread.start()
        
        self.check_rpa_thread(rpa_thread)
    
    def show_overlay(self):
        self.overlay = Toplevel(self.master)
        self.overlay.wm_overrideredirect(True)
        self.overlay.wm_attributes("-topmost", True)
        self.overlay.wm_attributes("-alpha", 0.7)
        
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        overlay_width = 800
        overlay_height = 100
        x = (screen_width // 2) - (overlay_width // 2)
        y = (screen_height // 2) - (overlay_height // 2)
        self.overlay.geometry(f"{overlay_width}x{overlay_height}+{x}+{y}")
        
        self.progress_label = Label(self.overlay, text="작업 시작 중...", font=("Helvetica", 16), bg="lightgray")
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
            self.show_results_popup()
    
    def show_results_popup(self):
        success_count = sum(1 for r in self.task_results if r["status"] == "성공")
        fail_count = len(self.task_results) - success_count
        
        summary_message = f"""RPA 자동 점검 완료!

총 {len(self.task_results)}개 작업 중:
  성공: {success_count}개
  실패: {fail_count}개
"""
        
        if fail_count > 0:
            summary_message += "\n실패한 작업 목록:\n"
            for r in self.task_results:
                if r["status"] == "실패":
                    summary_message += f"- {r['name']}: {r['error'] if r['error'] else '알 수 없는 오류'}\n"
        
        messagebox.showinfo("RPA 결과", summary_message)
        self.master.destroy()
    
    def hide_overlay(self):
        if self.overlay:
            self.overlay.destroy()
            self.overlay = None
            self.progress_label = None

if __name__ == "__main__":
    root = Tk()
    gui = RPA_GUI(root)
    root.mainloop()
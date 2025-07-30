# -*- coding: utf-8 -*-
import subprocess
import time
import pygetwindow as gw
import pyautogui
from datetime import datetime
from tkinter import messagebox, StringVar
import customtkinter as ctk
from PIL import Image, ImageTk
import threading
import os
import sys
import locale

# 한글 로케일 설정
try:
    locale.setlocale(locale.LC_ALL, 'ko_KR.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_ALL, 'Korean_Korea.949')
    except locale.Error:
        pass  # 로케일 설정 실패시 기본값 사용

# docxrpa.py에서 필요한 함수는 나중에 임포트 (지연 로딩)

def get_system_paths():
    """시스템 경로들을 동적으로 찾기"""
    # Windows 시스템 디렉토리 (보통 C:\Windows\System32)
    system32_dir = os.environ.get('SYSTEMROOT', 'C:\\Windows') + '\\System32'
    cmd_path = os.path.join(system32_dir, 'cmd.exe')
    
    # V3 경로 찾기 (여러 가능한 위치 시도)
    v3_possible_paths = [
        r'C:\Program Files\AhnLab\V3IS90\V3UI.exe',
        r'C:\Program Files (x86)\AhnLab\V3IS90\V3UI.exe',
        r'C:\Program Files\AhnLab\V3IS80\V3UI.exe',
        r'C:\Program Files (x86)\AhnLab\V3IS80\V3UI.exe',
        r'C:\Program Files\AhnLab\V3Lite40\v3lite4.exe',
        r'C:\Program Files (x86)\AhnLab\V3Lite40\v3lite4.exe',
        r'C:\Program Files\AhnLab\V3Lite\v3lite4.exe',
        r'C:\Program Files (x86)\AhnLab\V3Lite\v3lite4.exe'
    ]
    
    v3_path = None
    for path in v3_possible_paths:
        if os.path.exists(path):
            v3_path = path
            break
    
    return {
        'cmd': cmd_path,
        'v3': v3_path
    }

# 시스템 경로 초기화
SYSTEM_PATHS = get_system_paths()
CMD_PATH = SYSTEM_PATHS['cmd']

def capture_systeminfo_and_close(gui_instance=None):
    """Systeminfo 실행, 캡쳐, 창 닫기를 한 번에 처리"""
    # CMD 창 열고 systeminfo 실행
    cmd_command = f'{CMD_PATH} /c start /max {CMD_PATH} /k systeminfo'
    subprocess.Popen(cmd_command, shell=False)
    time.sleep(10)  # systeminfo 로딩 대기 (10초)
    
    # 스크린샷 캡쳐
    if gui_instance: gui_instance.hide_overlay()
    screenshot = pyautogui.screenshot()
    if gui_instance: gui_instance.show_overlay()
    
    # CMD 창 닫기
    windows = gw.getWindowsWithTitle(CMD_PATH)
    for window in windows:
        try:
            window.close()
        except:
            pass
    
    print("Systeminfo 캡쳐 완료")
    return screenshot

def capture_mac_address_and_close(gui_instance=None):
    """MAC 주소 확인, 캡쳐, 창 닫기를 한 번에 처리"""
    cmd_command = f'{CMD_PATH} /c start /max {CMD_PATH} /k "ipconfig /all | more"'
    subprocess.Popen(cmd_command, shell=False)
    time.sleep(5)
    
    if gui_instance: gui_instance.hide_overlay()
    screenshot = pyautogui.screenshot()
    if gui_instance: gui_instance.show_overlay()
    
    windows = gw.getWindowsWithTitle(CMD_PATH)
    for window in windows:
        try:
            window.close()
        except:
            pass
    
    print("MAC 주소 캡쳐 완료")
    return screenshot

def capture_screensaver_and_close(gui_instance=None):
    """화면 보호기 설정 열기, 캡쳐, 창 닫기를 한 번에 처리"""
    subprocess.Popen([CMD_PATH, '/c', 'control', 'desk.cpl,,@screensaver'])
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
    v3_path = SYSTEM_PATHS['v3']
    if not v3_path:
        print("V3를 찾을 수 없습니다. 다음 경로들을 확인하세요:")
        print("- C:\\Program Files\\AhnLab\\V3IS90\\V3UI.exe")
        print("- C:\\Program Files (x86)\\AhnLab\\V3IS90\\V3UI.exe")
        print("- C:\\Program Files\\AhnLab\\V3IS80\\V3UI.exe")
        print("- C:\\Program Files\\AhnLab\\V3Lite40\\v3lite4.exe")
        return None
    
    subprocess.Popen([CMD_PATH, '/c', v3_path])
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
    
    # 창 찾기 및 캡쳐 (V3IS90, V3 Lite 등 다양한 버전 지원)
    windows = gw.getWindowsWithTitle('AhnLab V3')
    if not windows:
        windows = gw.getWindowsWithTitle('V3')
    if not windows:
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
        
        # V3 창이 닫힐 때까지 Alt+F4 반복 (개선된 버전)
        print("V3 종료 시도 중...")
        close_attempts = 0
        max_attempts = 8
        
        while close_attempts < max_attempts:
            # V3 창이 있는지 확인
            v3_windows = gw.getWindowsWithTitle('AhnLab V3')
            if not v3_windows:
                v3_windows = gw.getWindowsWithTitle('V3')
            if not v3_windows:
                v3_windows = gw.getWindowsWithTitle('AhnLab V3 Lite')
            if not v3_windows:
                print("V3 창이 완전히 닫혔습니다.")
                break
            
            # V3 창을 정확히 활성화하고 Alt+F4
            try:
                v3_window = v3_windows[0]
                print(f"V3 창 닫기 시도 {close_attempts + 1}/{max_attempts}")
                
                # 창 활성화 전에 잠시 대기
                time.sleep(0.2)
                v3_window.activate()
                time.sleep(0.5)  # 창 활성화 대기 시간 증가
                
                # 활성 창이 V3인지 다시 확인
                active_window = gw.getActiveWindow()
                if active_window and ('AhnLab V3' in active_window.title or 'V3' in active_window.title):
                    print("V3 창이 활성화됨, Alt+F4 전송")
                    pyautogui.hotkey('alt', 'f4')
                    time.sleep(0.8)  # Alt+F4 후 대기 시간 증가
                    
                    # 종료 확인 대화상자가 있을 수 있으므로 Enter (조건부)
                    remaining_windows = gw.getWindowsWithTitle('AhnLab V3')
                    if not remaining_windows:
                        remaining_windows = gw.getWindowsWithTitle('V3')
                    if not remaining_windows:
                        remaining_windows = gw.getWindowsWithTitle('AhnLab V3 Lite')
                    if remaining_windows:  # 아직 창이 있으면 Enter 눌러서 확인
                        pyautogui.press('enter')
                        time.sleep(0.5)
                else:
                    print("V3 창 활성화 실패, 재시도...")
                    time.sleep(0.3)
            except Exception as e:
                print(f"V3 창 닫기 오류: {e}")
                time.sleep(0.5)
            
            close_attempts += 1
        
        # 최종 확인 및 강제 정리
        final_v3_windows = gw.getWindowsWithTitle('AhnLab V3')
        if not final_v3_windows:
            final_v3_windows = gw.getWindowsWithTitle('V3')
        if not final_v3_windows:
            final_v3_windows = gw.getWindowsWithTitle('AhnLab V3 Lite')
        if final_v3_windows:
            print(f"V3 창 닫기 시도 {max_attempts}회 후에도 창이 남아있음")
        else:
            print("V3 창 닫기 완료")
        
        # 키보드 상태 정리 (혹시 키가 눌린 상태로 남아있을 수 있음)
        pyautogui.press('esc')
        time.sleep(0.2)
        
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
    subprocess.Popen([CMD_PATH, '/c', 'control', 'appwiz.cpl'])
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

def capture_shared_folders_and_close(gui_instance=None):
    """공유폴더 확인, 캡쳐, 창 닫기를 한 번에 처리"""
    cmd_command = f'{CMD_PATH} /c start /max {CMD_PATH} /k "net share"'
    subprocess.Popen(cmd_command, shell=False)
    time.sleep(5)
    
    if gui_instance: gui_instance.hide_overlay()
    screenshot = pyautogui.screenshot()
    if gui_instance: gui_instance.show_overlay()
    
    # CMD 창 닫기
    windows = gw.getWindowsWithTitle(CMD_PATH)
    for window in windows:
        try:
            window.close()
        except:
            pass
    
    print("공유폴더 삭제 확인 캡쳐 완료")
    return screenshot

def capture_system_tray_icons(gui_instance=None):
    """시스템 트레이 숨겨진 아이콘 표시 및 캡쳐 (오른쪽 하단 영역만)"""
    # Win + B로 시스템 트레이로 포커스 이동
    pyautogui.hotkey('win', 'b')
    time.sleep(0.5)
    
    # Enter로 숨겨진 아이콘 펼치기
    pyautogui.press('enter')
    time.sleep(0.5)
    
    # 화면 해상도 확인 후 오른쪽 하단 영역만 캡쳐
    screen_width, screen_height = pyautogui.size()
    
    # 오른쪽 하단 500x500 영역 캡쳐 (시스템 트레이 영역)
    tray_width = 500
    tray_height = 500
    left = screen_width - tray_width
    top = screen_height - tray_height
    
    print(f"시스템 트레이 영역 캡쳐: {tray_width}x{tray_height} at ({left}, {top})")
    
    if gui_instance: gui_instance.hide_overlay()
    screenshot = pyautogui.screenshot(region=(left, top, tray_width, tray_height))
    if gui_instance: gui_instance.show_overlay()
    
    # ESC로 닫기
    pyautogui.press('esc')
    time.sleep(0.3)
    
    print("필수 보안 프로그램 설치 여부 캡쳐 완료")
    return screenshot

def run_rpa_tasks(update_progress_callback, gui_instance):
    """RPA 작업 실행"""
    captured_images = {}
    output_filename = f"노트북_점검_보고서_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    
    # 작업 목록 - 통합된 함수들 사용
    tasks = [
        {"name": "공유폴더 삭제 확인", "func": capture_shared_folders_and_close, "row": 1, "col": 0},
        {"name": "Systeminfo 캡쳐", "func": capture_systeminfo_and_close, "row": 3, "col": 0},
        {"name": "MAC 주소 캡쳐", "func": capture_mac_address_and_close, "row": 3, "col": 1},
        {"name": "화면 보호기 설정 캡쳐", "func": capture_screensaver_and_close, "row": 7, "col": 0},
        {"name": "필수 보안 프로그램 설치 여부", "func": capture_system_tray_icons, "row": 11, "col": 1},
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
        print(f"=== 작업 {step_num}/{total_tasks} 시작: {task_name} ===")
        task_start_time = time.time()
        update_progress_callback(f"{task_name} 진행 중...", step_num - 1, total_tasks + 1)
        
        result = {"name": task_name, "status": "실패", "error": None}
        
        try:
            screenshot = task["func"](gui_instance)
            if screenshot:
                captured_images[(task["row"], task["col"])] = screenshot
                result["status"] = "성공"
        except Exception as e:
            result["error"] = str(e)
            print(f"Error during {task_name}: {e}")
        
        task_end_time = time.time()
        print(f"{task_name} 완료 (소요시간: {task_end_time - task_start_time:.2f}초)")
        gui_instance.task_results.append(result)
        
        if i < len(tasks) - 1:  # 마지막 작업이 아니면 대기
            time.sleep(2.0)  # 작업간 텀을 2초로 설정
    
    # 문서 생성 및 저장 (지연 로딩으로 임포트)
    update_progress_callback("문서 생성 중...", total_tasks, total_tasks + 1)
    print("문서 생성 시작...")
    start_time = time.time()
    
    # 문서 생성이 필요할 때만 docxrpa 모듈 임포트
    from docxrpa import create_base_report_document, insert_image_to_document_table
    
    document = create_base_report_document()
    for (row, col), image in captured_images.items():
        insert_image_to_document_table(document, image, row, col)
    
    document.save(output_filename)
    print(f"문서 생성 완료. 소요 시간: {time.time() - start_time:.2f}초")
    
    update_progress_callback(f"모든 작업 완료! 보고서 저장됨", total_tasks + 1, total_tasks + 1)

class RPA_GUI:
    def __init__(self, master):
        self.master = master
        self.setup_window()
        self.setup_theme()
        
        self.overlay = None
        self.progress_label = None
        self.task_results = []
        self.progress_var = StringVar()
        self.progress_var.set("준비 완료")
        
        self.create_widgets()
        
    def setup_window(self):
        """윈도우 기본 설정"""
        self.master.title("노트북 보안점검 자동화 도구")
        self.master.geometry("350x420")
        self.master.resizable(False, False)
        
        # 윈도우를 화면 중앙에 배치
        self.master.update_idletasks()
        width = 350
        height = 420
        x = (self.master.winfo_screenwidth() // 2) - (width // 2)
        y = (self.master.winfo_screenheight() // 2) - (height // 2)
        self.master.geometry(f'{width}x{height}+{x}+{y}')
        
    def setup_theme(self):
        """CustomTkinter 테마 설정"""
        # Windows 11 스타일 색상 설정
        ctk.set_appearance_mode("light")  # "dark" 또는 "light"
        ctk.set_default_color_theme("blue")  # Windows 11 Blue
        
        # 커스텀 색상 정의
        self.colors = {
            'primary': '#0078D4',      # Windows 11 Blue
            'success': '#107C10',      # Green
            'warning': '#FFB900',      # Yellow
            'danger': '#D83B01',       # Red
            'light': '#F3F3F3',        # Light gray
            'dark': '#323130',         # Dark gray
            'card_bg': '#FFFFFF',      # Card background
            'text_primary': '#323130',  # Primary text
            'text_secondary': '#605E5C' # Secondary text
        }
        
    def create_widgets(self):
        """위젯 생성 및 배치"""
        # 메인 프레임 (스크롤 없이)
        self.main_frame = ctk.CTkFrame(self.master, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 헤더 (간소화)
        title_label = ctk.CTkLabel(self.main_frame, 
                                  text="🛡️ 노트북 보안점검 도구", 
                                  font=ctk.CTkFont(family="Segoe UI Variable", size=20, weight="bold"),
                                  text_color=self.colors['primary'])
        title_label.pack(pady=(0, 10))
        
        # 점검 항목 (간소화)
        check_title = ctk.CTkLabel(self.main_frame, 
                                  text="📋 점검 항목",
                                  font=ctk.CTkFont(family="Segoe UI Variable", size=16, weight="bold"),
                                  text_color=self.colors['text_primary'])
        check_title.pack(anchor="w", pady=(0, 5))
        
        # 점검 항목들을 간단한 텍스트로 표시
        items_text = "• 공유폴더/시스템정보  • MAC주소/화면보호기\n• 안티바이러스 상태  • 기타 보안 점검"
        items_label = ctk.CTkLabel(self.main_frame,
                                  text=items_text,
                                  font=ctk.CTkFont(size=12),
                                  text_color=self.colors['text_secondary'],
                                  justify="left")
        items_label.pack(anchor="w", pady=(0, 8))
        
        
        # 진행 상태 (간소화)
        progress_title = ctk.CTkLabel(self.main_frame, 
                                     text="📊 진행 상태",
                                     font=ctk.CTkFont(family="Segoe UI Variable", size=16, weight="bold"),
                                     text_color=self.colors['text_primary'])
        progress_title.pack(anchor="w", pady=(10, 5))
        
        self.progress_display = ctk.CTkLabel(self.main_frame, 
                                            textvariable=self.progress_var,
                                            font=ctk.CTkFont(size=13),
                                            text_color=self.colors['success'])
        self.progress_display.pack(anchor="w", pady=(0, 5))
        
        # 프로그레스 바
        self.progress_bar = ctk.CTkProgressBar(self.main_frame, 
                                              width=320,
                                              height=15,
                                              corner_radius=8,
                                              progress_color=self.colors['primary'])
        self.progress_bar.pack(pady=(0, 10))
        self.progress_bar.set(0)
        
        # 시작 버튼
        self.start_button = ctk.CTkButton(self.main_frame, 
                                         text="🚀 보안점검 시작",
                                         font=ctk.CTkFont(size=16, weight="bold"),
                                         command=self.start_rpa,
                                         height=40,
                                         corner_radius=10,
                                         fg_color=self.colors['primary'],
                                         hover_color=self.colors['dark'])
        self.start_button.pack(fill="x", pady=(0, 10))
        
        # 경고 (간소화)
        warning_text = "⚠️ 점검 중에는 컴퓨터를 사용하지 마세요."
        warning_label = ctk.CTkLabel(self.main_frame, 
                                    text=warning_text,
                                    font=ctk.CTkFont(size=12),
                                    text_color=self.colors['warning'],
                                    wraplength=320)
        warning_label.pack(pady=(0, 5))
        
        # 푸터
        footer_label = ctk.CTkLabel(self.main_frame,
                                   text="Version 2.0",
                                   font=ctk.CTkFont(size=11),
                                   text_color=self.colors['text_secondary'])
        footer_label.pack(pady=(3, 0))
    
    def start_rpa(self):
        """RPA 작업 시작"""
        # CustomTkinter 스타일 확인 대화상자
        dialog = ctk.CTkToplevel(self.master)
        dialog.title("보안점검 시작")
        dialog.geometry("400x280")
        dialog.resizable(False, False)
        
        # 중앙 배치
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - 200
        y = (dialog.winfo_screenheight() // 2) - 140
        dialog.geometry(f"400x280+{x}+{y}")
        dialog.transient(self.master)
        dialog.grab_set()
        
        # 다이얼로그 내용
        frame = ctk.CTkFrame(dialog)
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        title = ctk.CTkLabel(frame, text="🛡️ 보안점검을 시작하시겠습니까?",
                            font=ctk.CTkFont(size=16, weight="bold"))
        title.pack(pady=(0, 20))
        
        info_text = "⚠️ 점검 중에는 컴퓨터를 조작하지 마세요\n⏱️ 예상 소요시간: 약 3-5분\n📄 완료 후 Word 문서가 자동 생성됩니다"
        info_label = ctk.CTkLabel(frame, text=info_text,
                                 font=ctk.CTkFont(size=13),
                                 justify="left")
        info_label.pack(pady=(0, 30))
        
        # 버튼 프레임
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(fill="x")
        
        def on_yes():
            dialog.destroy()
            self._start_rpa_process()
            
        def on_no():
            dialog.destroy()
        
        yes_btn = ctk.CTkButton(btn_frame, text="시작", 
                               command=on_yes,
                               width=100, height=35,
                               fg_color=self.colors['primary'])
        yes_btn.pack(side="left", expand=True, padx=(0, 10))
        
        no_btn = ctk.CTkButton(btn_frame, text="취소", 
                              command=on_no,
                              width=100, height=35,
                              fg_color=self.colors['dark'])
        no_btn.pack(side="right", expand=True, padx=(10, 0))
        
    def _start_rpa_process(self):
        """실제 RPA 프로세스 시작"""
        # UI 업데이트
        self.start_button.configure(state="disabled", text="점검 진행 중...")
        self.progress_var.set("작업 초기화 중...")
        self.progress_bar.set(0)
        self.show_overlay()
        
        # 작업 스레드 시작
        rpa_thread = threading.Thread(target=run_rpa_tasks, args=(self.update_progress_enhanced, self))
        rpa_thread.daemon = True
        rpa_thread.start()
        
        self.check_rpa_thread(rpa_thread)
    
    def update_progress_enhanced(self, message, current_task=0, total_tasks=10):
        """개선된 진행률 업데이트"""
        if self.progress_label:
            progress_percentage = (current_task / total_tasks)
            self.master.after(0, lambda: self.progress_var.set(f"{message} ({current_task}/{total_tasks})"))
            self.master.after(0, lambda: self.progress_bar.set(progress_percentage))
            self.master.after(0, lambda: self.progress_label.configure(text=f"{message}\n조작하지 마세요"))
            
            # 오버레이 프로그레스 바도 업데이트
            if hasattr(self, 'overlay_progress'):
                self.master.after(0, lambda: self.overlay_progress.set(progress_percentage))
            
            # 체크박스 업데이트 (작업 완료 시)
            if current_task > 0 and current_task <= len(self.check_vars):
                self.master.after(0, lambda: self.check_vars[current_task-1].set(True))
    
    def show_overlay(self):
        """개선된 오버레이 창"""
        self.overlay = ctk.CTkToplevel(self.master)
        self.overlay.wm_overrideredirect(True)
        self.overlay.wm_attributes("-topmost", True)
        self.overlay.wm_attributes("-alpha", 0.95)
        
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        overlay_width = 500
        overlay_height = 180
        x = (screen_width // 2) - (overlay_width // 2)
        y = (screen_height // 2) - (overlay_height // 2)
        self.overlay.geometry(f"{overlay_width}x{overlay_height}+{x}+{y}")
        
        # 메인 카드 프레임 (둥근 모서리와 그림자 효과)
        overlay_card = ctk.CTkFrame(self.overlay, 
                                   corner_radius=20,
                                   fg_color=self.colors['card_bg'],
                                   border_width=2,
                                   border_color=self.colors['primary'])
        overlay_card.pack(expand=True, fill='both', padx=10, pady=10)
        
        # 내부 프레임
        inner_frame = ctk.CTkFrame(overlay_card, fg_color="transparent")
        inner_frame.pack(expand=True, fill='both', padx=30, pady=25)
        
        # 애니메이션 로딩 아이콘과 제목
        title_frame = ctk.CTkFrame(inner_frame, fg_color="transparent")
        title_frame.pack()
        
        title_label = ctk.CTkLabel(title_frame, 
                                  text="🛡️ 보안점검 진행 중", 
                                  font=ctk.CTkFont(size=20, weight="bold"),
                                  text_color=self.colors['primary'])
        title_label.pack()
        
        # 진행 상태 라벨
        self.progress_label = ctk.CTkLabel(inner_frame, 
                                          text="작업 준비 중...\n조작하지 마세요", 
                                          font=ctk.CTkFont(size=14),
                                          text_color=self.colors['text_secondary'])
        self.progress_label.pack(pady=(15, 10))
        
        # 프로그레스 바 (오버레이용)
        self.overlay_progress = ctk.CTkProgressBar(inner_frame, 
                                                  width=400,
                                                  height=15,
                                                  corner_radius=8,
                                                  progress_color=self.colors['primary'])
        self.overlay_progress.pack()
        self.overlay_progress.set(0)
    
    def update_progress(self, message):
        if self.progress_label:
            self.master.after(0, lambda: self.progress_label.config(text=message + " (조작하지 마세요)"))
    
    def check_rpa_thread(self, rpa_thread):
        if rpa_thread.is_alive():
            self.master.after(100, self.check_rpa_thread, rpa_thread)
        else:
            self.hide_overlay()
            self.start_button.configure(state="normal", text="🚀 보안점검 시작")
            self.progress_var.set("작업 완료!")
            self.progress_bar.set(1.0)
            self.show_results_popup()
    
    def show_results_popup(self):
        success_count = sum(1 for r in self.task_results if r["status"] == "성공")
        fail_count = len(self.task_results) - success_count
        total_count = len(self.task_results)
        
        # 성공률 계산
        success_rate = (success_count / total_count * 100) if total_count > 0 else 0
        
        # 결과 다이얼로그
        dialog = ctk.CTkToplevel(self.master)
        dialog.title("점검 결과")
        dialog.geometry("450x400")
        dialog.resizable(False, False)
        
        # 중앙 배치
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - 225
        y = (dialog.winfo_screenheight() // 2) - 200
        dialog.geometry(f"450x400+{x}+{y}")
        dialog.transient(self.master)
        dialog.grab_set()
        
        # 메인 프레임
        main_frame = ctk.CTkFrame(dialog)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 결과 아이콘과 제목
        if fail_count == 0:
            icon = "✅"
            title_text = "보안점검 완료"
            title_color = self.colors['success']
        else:
            icon = "⚠️"
            title_text = "보안점검 완료 (일부 실패)"
            title_color = self.colors['warning']
        
        title_label = ctk.CTkLabel(main_frame, 
                                  text=f"{icon} {title_text}",
                                  font=ctk.CTkFont(size=20, weight="bold"),
                                  text_color=title_color)
        title_label.pack(pady=(0, 20))
        
        # 결과 카드
        result_card = ctk.CTkFrame(main_frame, corner_radius=10, fg_color=self.colors['light'])
        result_card.pack(fill="x", pady=(0, 20))
        
        # 통계 표시
        stats_frame = ctk.CTkFrame(result_card, fg_color="transparent")
        stats_frame.pack(fill="x", padx=20, pady=15)
        
        stats = [
            ("전체 작업", total_count, self.colors['text_primary']),
            ("성공", success_count, self.colors['success']),
            ("실패", fail_count, self.colors['danger'] if fail_count > 0 else self.colors['text_secondary']),
            ("성공률", f"{success_rate:.0f}%", self.colors['primary'])
        ]
        
        for i, (label, value, color) in enumerate(stats):
            stat_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
            stat_frame.grid(row=i//2, column=i%2, padx=10, pady=5, sticky="ew")
            stats_frame.grid_columnconfigure(i%2, weight=1)
            
            ctk.CTkLabel(stat_frame, text=f"{label}:", 
                        font=ctk.CTkFont(size=13)).pack(side="left")
            ctk.CTkLabel(stat_frame, text=str(value), 
                        font=ctk.CTkFont(size=13, weight="bold"),
                        text_color=color).pack(side="right")
        
        # 파일 정보
        file_info = ctk.CTkLabel(main_frame,
                                text="📄 보고서가 자동으로 생성되었습니다\n💾 파일 위치: 프로그램 실행 폴더",
                                font=ctk.CTkFont(size=12),
                                text_color=self.colors['text_secondary'])
        file_info.pack(pady=(0, 20))
        
        # 버튼 프레임
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")
        
        def on_exit():
            dialog.destroy()
            self.master.destroy()
            
        def on_continue():
            dialog.destroy()
            self.progress_var.set("새로운 점검을 위해 준비 완료")
            self.progress_bar.set(0)
            # 체크박스 초기화
            for var in self.check_vars:
                var.set(False)
        
        continue_btn = ctk.CTkButton(btn_frame, text="계속", 
                                    command=on_continue,
                                    width=100, height=35,
                                    fg_color=self.colors['primary'])
        continue_btn.pack(side="left", expand=True, padx=(0, 10))
        
        exit_btn = ctk.CTkButton(btn_frame, text="종료", 
                                command=on_exit,
                                width=100, height=35,
                                fg_color=self.colors['danger'])
        exit_btn.pack(side="right", expand=True, padx=(10, 0))
    
    def hide_overlay(self):
        if self.overlay:
            self.overlay.destroy()
            self.overlay = None
            self.progress_label = None

if __name__ == "__main__":
    # CustomTkinter 설정
    ctk.set_appearance_mode("light")  # "dark" 또는 "light"
    ctk.set_default_color_theme("blue")  # Windows 11 스타일
    
    root = ctk.CTk()
    gui = RPA_GUI(root)
    root.mainloop()
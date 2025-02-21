#!/usr/bin/env python3
import time
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from pynput import keyboard, mouse

DRAG_THRESHOLD = 5  # 드래그 시작 전 최소 이동 픽셀
RECORD_WAIT_THRESHOLD = 0.1  # 이벤트 사이 최소 대기시간 (초)

# 색상 및 폰트 설정
BG_COLOR = "#2C2F33"
FRAME_BG = BG_COLOR
LABEL_BG = BG_COLOR
LABEL_FG = "white"
ENTRY_BG = "#23272A"
ENTRY_FG = "white"
LISTBOX_BG = "#23272A"
LISTBOX_FG = "white"
BUTTON_BG = "#7289DA"
BUTTON_FG = "white"
BUTTON_ACTIVE_BG = "#5b6eae"
FONT = ("Helvetica", 12)

class ManualMacroGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BLOUplanet's Macro - 키보드/마우스 매크로 프로그램")
        self.geometry("800x900")
        self.resizable(False, False)
        self.configure(bg=BG_COLOR)
        
        # 매크로 명령 관련 변수들
        self.commands = []  # 매크로 명령들을 저장하는 리스트
        self.macro_running = False
        self.drag_original_index = None  # 드래그 시작 시 항목 인덱스
        self.dragged_command = None      # 드래그 시작 시 선택한 명령 객체
        self.ghost = None                # 드래그 중 표시할 반투명 ghost
        self.drop_index = None           # 현재 예상 드롭 위치
        self.drop_indicator = None       # 드래그 중에만 보일 drop indicator (빨간 선)
        self._drag_start_x = None
        self._drag_start_y = None
        
        # 동작 기록 관련 변수
        self.action_recording = False
        self.recorded_commands = []      # 동작 기록으로 생성된 명령들
        self.last_record_time = 0
        self.action_keyboard_listener = None
        self.action_mouse_listener = None

        # 단축키 관련 변수 (기존 매크로 단축키: f2/f3, 동작 기록 단축키: f4/f5)
        self.action_start_hotkey_var = tk.StringVar(value="f4")
        self.action_stop_hotkey_var = tk.StringVar(value="f5")
        
        # 배경 클릭 시 포커스 해제
        self.bind_all("<Button-1>", self.clear_focus, add="+")
        
        # --- 명령 목록 영역 ---
        self.frame_list = tk.Frame(self, bg=FRAME_BG)
        self.frame_list.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        
        self.listbox = tk.Listbox(self.frame_list, width=80, height=10, bg=LISTBOX_BG, fg=LISTBOX_FG,
                                  font=FONT, selectbackground=BUTTON_BG, selectforeground="white", relief=tk.FLAT)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.listbox.bind("<Double-Button-1>", self.on_listbox_double_click)
        self.listbox.bind("<ButtonPress-1>", self.on_start_drag)
        self.listbox.bind("<B1-Motion>", self.on_drag_motion)
        self.listbox.bind("<ButtonRelease-1>", self.on_drag_stop)
        
        self.scrollbar = tk.Scrollbar(self.frame_list, orient=tk.VERTICAL)
        self.scrollbar.config(command=self.listbox.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=self.scrollbar.set)
        
        # --- 명령 추가용 에디터 영역 ---
        self.frame_editor = tk.Frame(self, bg=FRAME_BG)
        self.frame_editor.pack(padx=10, pady=5, fill=tk.X)
        tk.Label(self.frame_editor, text="명령 종류:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .grid(row=0, column=0, padx=5, pady=5)
        self.command_type_var = tk.StringVar(value="Key Tap")
        self.option_menu = tk.OptionMenu(self.frame_editor, self.command_type_var,
                                         "Key Tap", "Wait", "Mouse Click", "Key Hold", "Mouse Hold", "Mouse Scroll",
                                         command=self.update_param_fields)
        self.option_menu.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
        self.option_menu["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
        self.option_menu.grid(row=0, column=1, padx=5, pady=5)
        self.frame_params = tk.Frame(self.frame_editor, bg=FRAME_BG)
        self.frame_params.grid(row=0, column=2, padx=5, pady=5)
        self.param_entries = {}
        self.update_param_fields("Key Tap")
        self.button_add = tk.Button(self.frame_editor, text="명령 추가", command=self.add_command,
                                    bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                    activebackground=BUTTON_ACTIVE_BG)
        self.button_add.grid(row=0, column=3, padx=5, pady=5)
        
        # --- 제어 버튼 영역 ---
        self.frame_controls = tk.Frame(self, bg=FRAME_BG)
        self.frame_controls.pack(padx=10, pady=5, fill=tk.X)
        # 상단: 선택 명령 삭제, 매크로 실행, 매크로 중지
        self.frame_controls_top = tk.Frame(self.frame_controls, bg=FRAME_BG)
        self.frame_controls_top.pack(fill=tk.X)
        self.button_remove = tk.Button(self.frame_controls_top, text="선택 명령 삭제", command=self.remove_command,
                                       bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                       activebackground=BUTTON_ACTIVE_BG)
        self.button_remove.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_play = tk.Button(self.frame_controls_top, text="매크로 실행", command=self.play_macro,
                                     bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                     activebackground=BUTTON_ACTIVE_BG)
        self.button_play.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_stop = tk.Button(self.frame_controls_top, text="매크로 중지", command=self.stop_macro, state=tk.DISABLED,
                                     bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                     activebackground=BUTTON_ACTIVE_BG)
        self.button_stop.pack(side=tk.LEFT, padx=5, pady=5)
        # 하단: 매크로 저장, 매크로 불러오기, 반복 횟수
        self.frame_controls_bottom = tk.Frame(self.frame_controls, bg=FRAME_BG)
        self.frame_controls_bottom.pack(fill=tk.X, pady=(5,0))
        self.button_save = tk.Button(self.frame_controls_bottom, text="매크로 저장", command=self.save_macro,
                                     bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                     activebackground=BUTTON_ACTIVE_BG)
        self.button_save.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_load = tk.Button(self.frame_controls_bottom, text="매크로 불러오기", command=self.load_macro,
                                     bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                     activebackground=BUTTON_ACTIVE_BG)
        self.button_load.pack(side=tk.LEFT, padx=5, pady=5)
        tk.Label(self.frame_controls_bottom, text="반복 횟수 (0:무한):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .pack(side=tk.LEFT, padx=5, pady=5)
        self.entry_loop = tk.Entry(self.frame_controls_bottom, width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_loop.insert(0, "1")
        self.entry_loop.pack(side=tk.LEFT, padx=5, pady=5)
        
        # --- 기존 단축키 설정 영역 ---
        self.frame_hotkeys = tk.Frame(self, bg=FRAME_BG)
        self.frame_hotkeys.pack(padx=10, pady=5, fill=tk.X)
        tk.Label(self.frame_hotkeys, text="매크로 실행 단축키:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .pack(side=tk.LEFT, padx=5, pady=5)
        self.start_hotkey_var = tk.StringVar(value="f2")
        self.entry_start_hotkey = tk.Entry(self.frame_hotkeys, textvariable=self.start_hotkey_var,
                                           width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_start_hotkey.pack(side=tk.LEFT, padx=5, pady=5)
        tk.Label(self.frame_hotkeys, text="매크로 중지 단축키:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .pack(side=tk.LEFT, padx=5, pady=5)
        self.stop_hotkey_var = tk.StringVar(value="f3")
        self.entry_stop_hotkey = tk.Entry(self.frame_hotkeys, textvariable=self.stop_hotkey_var,
                                          width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_stop_hotkey.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_apply_hotkeys = tk.Button(self.frame_hotkeys, text="단축키 적용", command=self.apply_hotkeys,
                                              bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                              activebackground=BUTTON_ACTIVE_BG)
        self.button_apply_hotkeys.pack(side=tk.LEFT, padx=5, pady=5)
        
        # --- 동작 기록 영역 ---
        self.frame_action_record = tk.Frame(self, bg=FRAME_BG)
        self.frame_action_record.pack(padx=10, pady=5, fill=tk.X)
        tk.Label(self.frame_action_record, text="동작 기록 단축키 (시작/종료):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .pack(side=tk.LEFT, padx=5, pady=5)
        self.entry_action_start_hotkey = tk.Entry(self.frame_action_record, textvariable=self.action_start_hotkey_var,
                                                  width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_action_start_hotkey.pack(side=tk.LEFT, padx=5, pady=5)
        self.entry_action_stop_hotkey = tk.Entry(self.frame_action_record, textvariable=self.action_stop_hotkey_var,
                                                 width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_action_stop_hotkey.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_toggle_recording = tk.Button(self.frame_action_record, text="동작 기록 시작", command=self.toggle_action_recording,
                                                 bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                                 activebackground=BUTTON_ACTIVE_BG)
        self.button_toggle_recording.pack(side=tk.LEFT, padx=5, pady=5)
        
        # --- 로그 출력 영역 ---
        self.text_log = tk.Text(self, height=10, width=90, state=tk.NORMAL, bg=LISTBOX_BG, fg=LISTBOX_FG,
                                font=FONT, relief=tk.FLAT)
        self.text_log.pack(padx=10, pady=5)
        
        self.keyboard_controller = keyboard.Controller()
        self.mouse_controller = mouse.Controller()
        
        self.hotkey_listener = None
        self.after(100, self.start_hotkey_listener)
        
    # 포커스 해제
    def clear_focus(self, event):
        if not isinstance(event.widget, (tk.Entry, tk.Text)):
            self.focus_set()
            
    def log(self, message):
        self.text_log.insert(tk.END, message + "\n")
        self.text_log.see(tk.END)
        print(message)
        
    def update_param_fields(self, command_type):
        for widget in self.frame_params.winfo_children():
            widget.destroy()
        self.param_entries.clear()
        if command_type == "Key Tap":
            tk.Label(self.frame_params, text="키:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=2)
            entry_key = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_key.grid(row=0, column=1, padx=5, pady=2)
            def on_key_press(event):
                entry_key.delete(0, tk.END)
                entry_key.insert(0, event.keysym)
                return "break"
            entry_key.bind("<Key>", on_key_press)
            self.param_entries["key"] = entry_key
            tk.Label(self.frame_params, text="반복 횟수:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=2)
            entry_repeat = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_repeat.insert(0, "1")
            entry_repeat.grid(row=1, column=1, padx=5, pady=2)
            self.param_entries["repeat"] = entry_repeat
        elif command_type == "Wait":
            tk.Label(self.frame_params, text="대기 시간(초):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=2)
            entry_duration = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.grid(row=0, column=1, padx=5, pady=2)
            self.param_entries["duration"] = entry_duration
        elif command_type == "Mouse Click":
            tk.Label(self.frame_params, text="X:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=2)
            entry_x = tk.Entry(self.frame_params, width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_x.grid(row=0, column=1, padx=5, pady=2)
            self.param_entries["x"] = entry_x
            tk.Label(self.frame_params, text="Y:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=2, padx=5, pady=2)
            entry_y = tk.Entry(self.frame_params, width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_y.grid(row=0, column=3, padx=5, pady=2)
            self.param_entries["y"] = entry_y
            tk.Label(self.frame_params, text="버튼:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=4, padx=5, pady=2)
            self.mouse_button_var = tk.StringVar(value="left")
            option_button = tk.OptionMenu(self.frame_params, self.mouse_button_var, "left", "right", "middle")
            option_button.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
            option_button["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
            option_button.grid(row=0, column=5, padx=5, pady=2)
            self.param_entries["button"] = self.mouse_button_var
            # "마우스 위치 기록" 버튼은 새 행에 배치
            self.button_record_mouse = tk.Button(self.frame_params, text="마우스 위치 기록", command=self.record_mouse_position,
                                                 bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                                 activebackground=BUTTON_ACTIVE_BG)
            self.button_record_mouse.grid(row=1, column=0, columnspan=6, padx=5, pady=2, sticky="w")
        elif command_type == "Key Hold":
            tk.Label(self.frame_params, text="키:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=2)
            entry_key = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_key.grid(row=0, column=1, padx=5, pady=2)
            def on_key_press(event):
                entry_key.delete(0, tk.END)
                entry_key.insert(0, event.keysym)
                return "break"
            entry_key.bind("<Key>", on_key_press)
            self.param_entries["key"] = entry_key
            tk.Label(self.frame_params, text="누름 시간(초):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=2)
            entry_duration = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.insert(0, "1")
            entry_duration.grid(row=1, column=1, padx=5, pady=2)
            self.param_entries["duration"] = entry_duration
        elif command_type == "Mouse Hold":
            # 첫 행: X, Y, 버튼
            tk.Label(self.frame_params, text="X:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=2)
            entry_x = tk.Entry(self.frame_params, width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_x.grid(row=0, column=1, padx=5, pady=2)
            self.param_entries["x"] = entry_x
            tk.Label(self.frame_params, text="Y:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=2, padx=5, pady=2)
            entry_y = tk.Entry(self.frame_params, width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_y.grid(row=0, column=3, padx=5, pady=2)
            self.param_entries["y"] = entry_y
            tk.Label(self.frame_params, text="버튼:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=4, padx=5, pady=2)
            self.mouse_button_var = tk.StringVar(value="left")
            option_button = tk.OptionMenu(self.frame_params, self.mouse_button_var, "left", "right", "middle")
            option_button.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
            option_button["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
            option_button.grid(row=0, column=5, padx=5, pady=2)
            self.param_entries["button"] = self.mouse_button_var
            # 두번째 행: 누름 시간과 마우스 위치 기록 버튼
            tk.Label(self.frame_params, text="누름 시간(초):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=2)
            entry_duration = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.insert(0, "1")
            entry_duration.grid(row=1, column=1, padx=5, pady=2)
            self.param_entries["duration"] = entry_duration
            self.button_record_mouse = tk.Button(self.frame_params, text="마우스 위치 기록", command=self.record_mouse_position,
                                                 bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                                 activebackground=BUTTON_ACTIVE_BG)
            self.button_record_mouse.grid(row=1, column=2, columnspan=4, padx=5, pady=2, sticky="w")
        elif command_type == "Mouse Scroll":
            tk.Label(self.frame_params, text="수평 스크롤:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=2)
            entry_dx = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_dx.insert(0, "0")
            entry_dx.grid(row=0, column=1, padx=5, pady=2)
            self.param_entries["dx"] = entry_dx
            tk.Label(self.frame_params, text="수직 스크롤:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=2, padx=5, pady=2)
            entry_dy = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_dy.insert(0, "0")
            entry_dy.grid(row=0, column=3, padx=5, pady=2)
            self.param_entries["dy"] = entry_dy

    def record_mouse_position(self):
        self.button_record_mouse.config(text="좌클릭해서 마우스 위치 기록 완료", state=tk.DISABLED)
        self.log("마우스 위치 기록 대기 중... 좌클릭하면 기록이 완료됩니다.")
        def on_click(x, y, button, pressed):
            if button == mouse.Button.left and pressed:
                self.after(0, lambda: self.set_mouse_position(x, y))
                return False
        listener = mouse.Listener(on_click=on_click)
        listener.start()
        
    def set_mouse_position(self, x, y):
        if "x" in self.param_entries:
            self.param_entries["x"].delete(0, tk.END)
            self.param_entries["x"].insert(0, str(int(x)))
        if "y" in self.param_entries:
            self.param_entries["y"].delete(0, tk.END)
            self.param_entries["y"].insert(0, str(int(y)))
        self.log(f"마우스 위치 기록됨: ({int(x)}, {int(y)})")
        self.button_record_mouse.config(text="마우스 위치 기록", state=tk.NORMAL)
        
    def add_command(self):
        command_type = self.command_type_var.get()
        if command_type == "Key Tap":
            key = self.param_entries["key"].get().strip()
            if key == "":
                messagebox.showerror("오류", "키 값을 입력하세요.")
                return
            try:
                repeat = int(self.param_entries["repeat"].get().strip())
            except ValueError:
                messagebox.showerror("오류", "유효한 반복 횟수를 입력하세요.")
                return
            cmd = {"command": "key_tap", "key": key, "repeat": repeat}
            display_text = f"키 탭: {key} x {repeat}회"
        elif command_type == "Wait":
            try:
                duration = float(self.param_entries["duration"].get().strip())
            except ValueError:
                messagebox.showerror("오류", "유효한 대기 시간을 입력하세요.")
                return
            cmd = {"command": "wait", "duration": duration}
            display_text = f"대기: {duration}초"
        elif command_type == "Mouse Click":
            x_str = self.param_entries["x"].get().strip()
            y_str = self.param_entries["y"].get().strip()
            if not x_str or not y_str:
                messagebox.showerror("오류", "X, Y 값을 입력하세요.")
                return
            try:
                x = int(x_str)
                y = int(y_str)
            except ValueError:
                messagebox.showerror("오류", "유효한 정수 X, Y 값을 입력하세요.")
                return
            button = self.param_entries["button"].get()
            cmd = {"command": "mouse_click", "x": x, "y": y, "button": button}
            display_text = f"마우스 클릭: ({x}, {y}), 버튼: {button}"
        elif command_type == "Key Hold":
            key = self.param_entries["key"].get().strip()
            if key == "":
                messagebox.showerror("오류", "키 값을 입력하세요.")
                return
            try:
                duration = float(self.param_entries["duration"].get().strip())
            except ValueError:
                messagebox.showerror("오류", "유효한 누름 시간을 입력하세요.")
                return
            cmd = {"command": "key_hold", "key": key, "duration": duration}
            display_text = f"키 누름: {key} (누름 시간: {duration}초)"
        elif command_type == "Mouse Hold":
            x_str = self.param_entries["x"].get().strip()
            y_str = self.param_entries["y"].get().strip()
            if not x_str or not y_str:
                messagebox.showerror("오류", "X, Y 값을 입력하세요.")
                return
            try:
                x = int(x_str)
                y = int(y_str)
            except ValueError:
                messagebox.showerror("오류", "유효한 정수 X, Y 값을 입력하세요.")
                return
            button = self.param_entries["button"].get()
            try:
                duration = float(self.param_entries["duration"].get().strip())
            except ValueError:
                messagebox.showerror("오류", "유효한 누름 시간을 입력하세요.")
                return
            cmd = {"command": "mouse_hold", "x": x, "y": y, "button": button, "duration": duration}
            display_text = f"마우스 누름: ({x}, {y}), 버튼: {button} (누름 시간: {duration}초)"
        elif command_type == "Mouse Scroll":
            try:
                dx = int(self.param_entries["dx"].get().strip())
                dy = int(self.param_entries["dy"].get().strip())
            except ValueError:
                messagebox.showerror("오류", "유효한 스크롤 값을 입력하세요.")
                return
            cmd = {"command": "mouse_scroll", "dx": dx, "dy": dy}
            display_text = f"마우스 스크롤: 수평 {dx}, 수직 {dy}"
        else:
            return
        
        selected = self.listbox.curselection()
        if selected:
            index = selected[0] + 1
            self.commands.insert(index, cmd)
            self.listbox.insert(index, display_text)
        else:
            self.commands.append(cmd)
            self.listbox.insert(tk.END, display_text)
        self.log("명령 추가됨: " + display_text)
        
    def remove_command(self):
        selected = self.listbox.curselection()
        if not selected:
            messagebox.showerror("오류", "삭제할 명령을 선택하세요.")
            return
        index = selected[0]
        self.listbox.delete(index)
        removed = self.commands.pop(index)
        self.log("명령 삭제됨: " + str(removed))
        
    def play_macro(self):
        if not self.commands:
            messagebox.showinfo("정보", "실행할 명령이 없습니다.")
            return
        try:
            loop_count = int(self.entry_loop.get().strip())
        except ValueError:
            messagebox.showerror("오류", "유효한 반복 횟수를 입력하세요.")
            return
        self.log("매크로 실행 시작.")
        self.macro_running = True
        self.button_stop.config(state=tk.NORMAL)
        thread = threading.Thread(target=self.execute_macro, args=(loop_count,))
        thread.daemon = True
        thread.start()
        
    def execute_macro(self, loop_count):
        iteration = 0
        while self.macro_running and (loop_count == 0 or iteration < loop_count):
            self.log(f"반복 {iteration+1} 시작.")
            for cmd in self.commands:
                if not self.macro_running:
                    break
                if cmd["command"] == "key_tap":
                    key = cmd["key"]
                    repeat = cmd.get("repeat", 1)
                    for _ in range(repeat):
                        if not self.macro_running:
                            break
                        try:
                            key_obj = getattr(keyboard.Key, key.lower())
                        except AttributeError:
                            key_obj = key
                        self.keyboard_controller.press(key_obj)
                        self.keyboard_controller.release(key_obj)
                        self.log(f"키 탭 실행: {key}")
                        time.sleep(0.05)
                elif cmd["command"] == "key_hold":
                    key = cmd["key"]
                    duration = cmd["duration"]
                    try:
                        key_obj = getattr(keyboard.Key, key.lower())
                    except AttributeError:
                        key_obj = key
                    self.keyboard_controller.press(key_obj)
                    self.log(f"키 누름 시작: {key}")
                    time.sleep(duration)
                    self.keyboard_controller.release(key_obj)
                    self.log(f"키 누름 종료: {key}")
                elif cmd["command"] == "wait":
                    duration = cmd["duration"]
                    self.log(f"대기 시작: {duration}초")
                    time.sleep(duration)
                    self.log("대기 종료")
                elif cmd["command"] == "mouse_click":
                    x = cmd["x"]
                    y = cmd["y"]
                    button_str = cmd["button"]
                    if button_str == "left":
                        btn = mouse.Button.left
                    elif button_str == "right":
                        btn = mouse.Button.right
                    elif button_str == "middle":
                        btn = mouse.Button.middle
                    else:
                        btn = mouse.Button.left
                    self.mouse_controller.position = (x, y)
                    self.mouse_controller.click(btn)
                    self.log(f"마우스 클릭 실행: ({x}, {y}), 버튼: {button_str}")
                elif cmd["command"] == "mouse_hold":
                    x = cmd["x"]
                    y = cmd["y"]
                    button_str = cmd["button"]
                    if button_str == "left":
                        btn = mouse.Button.left
                    elif button_str == "right":
                        btn = mouse.Button.right
                    elif button_str == "middle":
                        btn = mouse.Button.middle
                    else:
                        btn = mouse.Button.left
                    self.mouse_controller.position = (x, y)
                    self.mouse_controller.press(btn)
                    self.log(f"마우스 누름 시작: ({x}, {y}), 버튼: {button_str}")
                    time.sleep(cmd["duration"])
                    self.mouse_controller.release(btn)
                    self.log(f"마우스 누름 종료: ({x}, {y}), 버튼: {button_str}")
                elif cmd["command"] == "mouse_scroll":
                    dx = cmd["dx"]
                    dy = cmd["dy"]
                    self.mouse_controller.scroll(dx, dy)
                    self.log(f"마우스 스크롤: 수평 {dx}, 수직 {dy}")
                time.sleep(0.1)
            iteration += 1
            self.log(f"반복 {iteration} 완료.")
        self.log("매크로 실행 완료.")
        self.macro_running = False
        self.button_stop.config(state=tk.DISABLED)
        
    def stop_macro(self):
        self.macro_running = False
        self.log("매크로 실행 중지 요청됨.")
        
    def save_macro(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", 
                                                 filetypes=[("JSON files", "*.json")])
        if not file_path:
            return
        try:
            with open(file_path, "w") as f:
                json.dump(self.commands, f, indent=4)
            self.log("매크로 저장됨: " + file_path)
        except Exception as e:
            messagebox.showerror("오류", "매크로 저장 실패: " + str(e))
        
    def load_macro(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not file_path:
            return
        try:
            with open(file_path, "r") as f:
                self.commands = json.load(f)
            self.listbox.delete(0, tk.END)
            for cmd in self.commands:
                disp = self.get_display_text(cmd)
                self.listbox.insert(tk.END, disp)
                self.log("불러온 명령: " + disp)
            self.log("매크로 불러오기 완료: " + file_path)
        except Exception as e:
            messagebox.showerror("오류", "매크로 불러오기 실패: " + str(e))
            
    def get_display_text(self, cmd):
        try:
            if cmd.get("command") == "key_tap":
                return f"키 탭: {cmd.get('key', '')} x {cmd.get('repeat', 1)}회"
            elif cmd.get("command") == "key_hold":
                return f"키 누름: {cmd.get('key', '')} (누름 시간: {cmd.get('duration', 0)}초)"
            elif cmd.get("command") == "wait":
                return f"대기: {cmd.get('duration', 0)}초"
            elif cmd.get("command") == "mouse_click":
                return f"마우스 클릭: ({cmd.get('x', 0)}, {cmd.get('y', 0)}), 버튼: {cmd.get('button', '')}"
            elif cmd.get("command") == "mouse_hold":
                return f"마우스 누름: ({cmd.get('x', 0)}, {cmd.get('y', 0)}), 버튼: {cmd.get('button', '')} (누름 시간: {cmd.get('duration', 0)}초)"
            elif cmd.get("command") == "mouse_scroll":
                return f"마우스 스크롤: 수평 {cmd.get('dx',0)}, 수직 {cmd.get('dy',0)}"
            else:
                return str(cmd)
        except Exception as e:
            self.log("명령 표시 오류: " + str(e))
            return str(cmd)
        
    def format_hotkey(self, key_str):
        key_str = key_str.strip().lower()
        if not key_str.startswith("<") and not key_str.endswith(">"):
            key_str = f"<{key_str}>"
        return key_str
        
    def start_hotkey_listener(self):
        mapping = {
            self.format_hotkey(self.start_hotkey_var.get()): self.on_hotkey_start,
            self.format_hotkey(self.stop_hotkey_var.get()): self.on_hotkey_stop,
            self.format_hotkey(self.action_start_hotkey_var.get()): self.start_action_recording,
            self.format_hotkey(self.action_stop_hotkey_var.get()): self.stop_action_recording
        }
        self.hotkey_listener = keyboard.GlobalHotKeys(mapping)
        self.hotkey_thread = threading.Thread(target=self.hotkey_listener.run, daemon=True)
        self.hotkey_thread.start()
        self.log("Hotkey listener started with 매크로: start={} stop={}; 동작 기록: start={} stop={}".format(
            self.format_hotkey(self.start_hotkey_var.get()),
            self.format_hotkey(self.stop_hotkey_var.get()),
            self.format_hotkey(self.action_start_hotkey_var.get()),
            self.format_hotkey(self.action_stop_hotkey_var.get())
        ))
        
    def stop_hotkey_listener(self):
        if self.hotkey_listener:
            self.hotkey_listener.stop()
            self.hotkey_listener = None
            self.log("Hotkey listener stopped.")
        
    def apply_hotkeys(self):
        self.stop_hotkey_listener()
        self.start_hotkey_listener()
        self.log("단축키가 적용되었습니다.")
        
    def on_hotkey_start(self):
        self.log("단축키로 매크로 실행 요청됨.")
        if not self.macro_running:
            self.play_macro()
        
    def on_hotkey_stop(self):
        self.log("단축키로 매크로 중지 요청됨.")
        self.stop_macro()
        
    def start_action_recording(self):
        if self.action_recording:
            return
        self.action_recording = True
        self.recorded_commands = []
        self.last_record_time = time.time()
        self.log("동작 기록 시작됨.")
        self.action_keyboard_listener = keyboard.Listener(on_release=self.action_on_key_release)
        self.action_mouse_listener = mouse.Listener(on_click=self.action_on_mouse_click)
        self.action_keyboard_listener.start()
        self.action_mouse_listener.start()
        self.button_toggle_recording.config(text="동작 기록 종료")
        
    def stop_action_recording(self):
        if not self.action_recording:
            return
        self.action_recording = False
        if self.action_keyboard_listener:
            self.action_keyboard_listener.stop()
            self.action_keyboard_listener = None
        if self.action_mouse_listener:
            self.action_mouse_listener.stop()
            self.action_mouse_listener = None
        self.log("동작 기록 종료됨. 기록된 동작을 명령 목록에 추가합니다.")
        for cmd in self.recorded_commands:
            self.commands.append(cmd)
            self.listbox.insert(tk.END, self.get_display_text(cmd))
        self.button_toggle_recording.config(text="동작 기록 시작")
        
    def action_on_key_release(self, key):
        if not self.action_recording:
            return
        now = time.time()
        dt = now - self.last_record_time
        if dt > RECORD_WAIT_THRESHOLD:
            wait_cmd = {"command": "wait", "duration": round(dt, 2)}
            self.recorded_commands.append(wait_cmd)
            self.log("기록된 대기: {}초".format(round(dt, 2)))
        start_key = getattr(keyboard.Key, self.action_start_hotkey_var.get().strip().lower(), None)
        stop_key  = getattr(keyboard.Key, self.action_stop_hotkey_var.get().strip().lower(), None)
        if key == start_key or key == stop_key:
            self.last_record_time = now
            return
        try:
            k = key.char
        except AttributeError:
            k = str(key)
        key_cmd = {"command": "key_tap", "key": k, "repeat": 1}
        self.recorded_commands.append(key_cmd)
        self.log("기록된 키 탭: {}".format(k))
        self.last_record_time = now
        
    def action_on_mouse_click(self, x, y, button, pressed):
        if not self.action_recording:
            return
        if button != mouse.Button.left or not pressed:
            return
        now = time.time()
        dt = now - self.last_record_time
        if dt > RECORD_WAIT_THRESHOLD:
            wait_cmd = {"command": "wait", "duration": round(dt, 2)}
            self.recorded_commands.append(wait_cmd)
            self.log("기록된 대기: {}초".format(round(dt, 2)))
        mouse_cmd = {"command": "mouse_click", "x": int(x), "y": int(y), "button": "left"}
        self.recorded_commands.append(mouse_cmd)
        self.log("기록된 마우스 클릭: ({}, {})".format(int(x), int(y)))
        self.last_record_time = now
        
    def toggle_action_recording(self):
        if self.action_recording:
            self.stop_action_recording()
        else:
            self.start_action_recording()
        
    def on_listbox_double_click(self, event):
        selection = self.listbox.curselection()
        if selection:
            self.edit_command(selection[0])
            
    def edit_command(self, index):
        cmd = self.commands[index]
        edit_win = tk.Toplevel(self)
        edit_win.title("명령 수정")
        edit_win.configure(bg=BG_COLOR)
        edit_win.wait_visibility()
        edit_win.grab_set()
        def save_changes():
            if cmd["command"] == "key_tap":
                new_key = entry_key.get().strip()
                try:
                    new_repeat = int(entry_repeat.get().strip())
                except ValueError:
                    messagebox.showerror("오류", "유효한 반복 횟수를 입력하세요.", parent=edit_win)
                    return
                if new_key == "":
                    messagebox.showerror("오류", "키 값을 입력하세요.", parent=edit_win)
                    return
                cmd["key"] = new_key
                cmd["repeat"] = new_repeat
            elif cmd["command"] == "key_hold":
                new_key = entry_key.get().strip()
                try:
                    new_duration = float(entry_duration.get().strip())
                except ValueError:
                    messagebox.showerror("오류", "유효한 누름 시간을 입력하세요.", parent=edit_win)
                    return
                if new_key == "":
                    messagebox.showerror("오류", "키 값을 입력하세요.", parent=edit_win)
                    return
                cmd["key"] = new_key
                cmd["duration"] = new_duration
            elif cmd["command"] == "wait":
                try:
                    new_duration = float(entry_duration.get().strip())
                except ValueError:
                    messagebox.showerror("오류", "유효한 대기 시간을 입력하세요.", parent=edit_win)
                    return
                cmd["duration"] = new_duration
            elif cmd["command"] == "mouse_click":
                try:
                    new_x = int(entry_x.get().strip())
                    new_y = int(entry_y.get().strip())
                except ValueError:
                    messagebox.showerror("오류", "유효한 정수 X, Y 값을 입력하세요.", parent=edit_win)
                    return
                new_button = var_button.get()
                cmd["x"] = new_x
                cmd["y"] = new_y
                cmd["button"] = new_button
            elif cmd["command"] == "mouse_hold":
                try:
                    new_x = int(entry_x.get().strip())
                    new_y = int(entry_y.get().strip())
                except ValueError:
                    messagebox.showerror("오류", "유효한 정수 X, Y 값을 입력하세요.", parent=edit_win)
                    return
                try:
                    new_duration = float(entry_duration.get().strip())
                except ValueError:
                    messagebox.showerror("오류", "유효한 누름 시간을 입력하세요.", parent=edit_win)
                    return
                new_button = var_button.get()
                cmd["x"] = new_x
                cmd["y"] = new_y
                cmd["button"] = new_button
                cmd["duration"] = new_duration
            elif cmd["command"] == "mouse_scroll":
                try:
                    new_dx = int(entry_dx.get().strip())
                    new_dy = int(entry_dy.get().strip())
                except ValueError:
                    messagebox.showerror("오류", "유효한 스크롤 값을 입력하세요.", parent=edit_win)
                    return
                cmd["dx"] = new_dx
                cmd["dy"] = new_dy
            self.listbox.delete(index)
            self.listbox.insert(index, self.get_display_text(cmd))
            self.log("명령 수정됨: " + self.get_display_text(cmd))
            edit_win.destroy()
        if cmd["command"] == "key_tap":
            tk.Label(edit_win, text="키:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=5)
            entry_key = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_key.insert(0, cmd["key"])
            entry_key.grid(row=0, column=1, padx=5, pady=5)
            def on_key_press(event):
                entry_key.delete(0, tk.END)
                entry_key.insert(0, event.keysym)
                return "break"
            entry_key.bind("<Key>", on_key_press)
            tk.Label(edit_win, text="반복 횟수:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=5)
            entry_repeat = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_repeat.insert(0, str(cmd.get("repeat", 1)))
            entry_repeat.grid(row=1, column=1, padx=5, pady=5)
        elif cmd["command"] == "key_hold":
            tk.Label(edit_win, text="키:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=5)
            entry_key = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_key.insert(0, cmd["key"])
            entry_key.grid(row=0, column=1, padx=5, pady=5)
            def on_key_press(event):
                entry_key.delete(0, tk.END)
                entry_key.insert(0, event.keysym)
                return "break"
            entry_key.bind("<Key>", on_key_press)
            tk.Label(edit_win, text="누름 시간(초):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=5)
            entry_duration = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.insert(0, str(cmd.get("duration", 1)))
            entry_duration.grid(row=1, column=1, padx=5, pady=5)
        elif cmd["command"] == "wait":
            tk.Label(edit_win, text="대기 시간(초):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=5)
            entry_duration = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.insert(0, str(cmd["duration"]))
            entry_duration.grid(row=0, column=1, padx=5, pady=5)
        elif cmd["command"] == "mouse_click":
            tk.Label(edit_win, text="X:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=5)
            entry_x = tk.Entry(edit_win, width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_x.insert(0, str(cmd["x"]))
            entry_x.grid(row=0, column=1, padx=5, pady=5)
            tk.Label(edit_win, text="Y:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=2, padx=5, pady=5)
            entry_y = tk.Entry(edit_win, width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_y.insert(0, str(cmd["y"]))
            entry_y.grid(row=0, column=3, padx=5, pady=5)
            tk.Label(edit_win, text="버튼:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=4, padx=5, pady=5)
            var_button = tk.StringVar(value=cmd["button"])
            option_button = tk.OptionMenu(edit_win, var_button, "left", "right", "middle")
            option_button.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
            option_button["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
            option_button.grid(row=0, column=5, padx=5, pady=5)
            def record_mouse_edit():
                record_button_edit.config(text="좌클릭해서 마우스 위치 기록 완료", state=tk.DISABLED)
                def on_click(x, y, button, pressed):
                    if button == mouse.Button.left and pressed:
                        entry_x.delete(0, tk.END)
                        entry_x.insert(0, str(int(x)))
                        entry_y.delete(0, tk.END)
                        entry_y.insert(0, str(int(y)))
                        edit_win.after(0, lambda: self.log(f"수정창에서 마우스 위치 기록됨: ({int(x)}, {int(y)})"))
                        record_button_edit.config(text="마우스 위치 기록", state=tk.NORMAL)
                        return False
                listener = mouse.Listener(on_click=on_click)
                listener.start()
            record_button_edit = tk.Button(edit_win, text="마우스 위치 기록", command=record_mouse_edit,
                                           bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                           activebackground=BUTTON_ACTIVE_BG)
            record_button_edit.grid(row=1, column=0, columnspan=6, padx=5, pady=5, sticky="w")
        elif cmd["command"] == "mouse_hold":
            # 첫 행: X, Y, 버튼
            tk.Label(edit_win, text="X:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=5)
            entry_x = tk.Entry(edit_win, width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_x.insert(0, str(cmd["x"]))
            entry_x.grid(row=0, column=1, padx=5, pady=5)
            tk.Label(edit_win, text="Y:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=2, padx=5, pady=5)
            entry_y = tk.Entry(edit_win, width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_y.insert(0, str(cmd["y"]))
            entry_y.grid(row=0, column=3, padx=5, pady=5)
            tk.Label(edit_win, text="버튼:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=4, padx=5, pady=5)
            var_button = tk.StringVar(value=cmd["button"])
            option_button = tk.OptionMenu(edit_win, var_button, "left", "right", "middle")
            option_button.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
            option_button["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
            option_button.grid(row=0, column=5, padx=5, pady=5)
            tk.Label(edit_win, text="누름 시간(초):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=5)
            entry_duration = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.insert(0, str(cmd.get("duration", 1)))
            entry_duration.grid(row=1, column=1, padx=5, pady=5)
            def record_mouse_edit():
                record_button_edit.config(text="좌클릭해서 마우스 위치 기록 완료", state=tk.DISABLED)
                def on_click(x, y, button, pressed):
                    if button == mouse.Button.left and pressed:
                        entry_x.delete(0, tk.END)
                        entry_x.insert(0, str(int(x)))
                        entry_y.delete(0, tk.END)
                        entry_y.insert(0, str(int(y)))
                        edit_win.after(0, lambda: self.log(f"수정창에서 마우스 위치 기록됨: ({int(x)}, {int(y)})"))
                        record_button_edit.config(text="마우스 위치 기록", state=tk.NORMAL)
                        return False
                listener = mouse.Listener(on_click=on_click)
                listener.start()
            record_button_edit = tk.Button(edit_win, text="마우스 위치 기록", command=record_mouse_edit,
                                           bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                           activebackground=BUTTON_ACTIVE_BG)
            record_button_edit.grid(row=1, column=2, columnspan=4, padx=5, pady=5, sticky="w")
        elif cmd["command"] == "mouse_scroll":
            tk.Label(edit_win, text="수평 스크롤:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=5)
            entry_dx = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_dx.insert(0, str(cmd["dx"]))
            entry_dx.grid(row=0, column=1, padx=5, pady=5)
            tk.Label(edit_win, text="수직 스크롤:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=2, padx=5, pady=5)
            entry_dy = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_dy.insert(0, str(cmd["dy"]))
            entry_dy.grid(row=0, column=3, padx=5, pady=5)
        tk.Button(edit_win, text="저장", command=save_changes,
                  bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT, activebackground=BUTTON_ACTIVE_BG)\
            .grid(row=10, column=0, padx=5, pady=10)
        tk.Button(edit_win, text="취소", command=edit_win.destroy,
                  bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT, activebackground=BUTTON_ACTIVE_BG)\
            .grid(row=10, column=1, padx=5, pady=10)
        
    def on_start_drag(self, event):
        index = self.listbox.nearest(event.y)
        if index < 0 or index >= len(self.commands):
            return
        self.drag_original_index = index
        self.dragged_command = self.commands[index]
        self._drag_start_x = event.x
        self._drag_start_y = event.y
        self.ghost = None
        self.drop_index = index
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(index)
        self.listbox.activate(index)
        
    def on_drag_motion(self, event):
        if self.ghost is None:
            dx = event.x - self._drag_start_x
            dy = event.y - self._drag_start_y
            if (dx**2 + dy**2)**0.5 < DRAG_THRESHOLD:
                return
            self.ghost = tk.Toplevel(self)
            self.ghost.overrideredirect(True)
            self.ghost.attributes("-alpha", 0.5)
            self.ghost.configure(bg=BG_COLOR)
            label = tk.Label(self.ghost, text=self.listbox.get(self.drag_original_index),
                              bg="lightgrey", borderwidth=2, relief="solid", font=FONT)
            label.pack()
            self.update_idletasks()
            x = self.listbox.winfo_rootx() + event.x
            y = self.listbox.winfo_rooty() + event.y
            self.ghost.geometry(f"+{x}+{y}")
            self.drop_index = self.drag_original_index
            self.draw_drop_indicator(self.drag_original_index)
        else:
            x = self.listbox.winfo_rootx() + event.x
            y = self.listbox.winfo_rooty() + event.y
            self.ghost.geometry(f"+{x}+{y}")
            new_index = self.listbox.nearest(event.y)
            bbox = self.listbox.bbox(new_index)
            if bbox and event.y > bbox[1] + bbox[3] // 2:
                new_index += 1
            if new_index != self.drop_index:
                self.drop_index = new_index
                self.draw_drop_indicator(new_index)
        if self.drag_original_index is not None:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(self.drag_original_index)
            self.listbox.activate(self.drag_original_index)
        
    def on_drag_stop(self, event):
        if self.ghost:
            self.ghost.destroy()
            self.ghost = None
        if self.drop_indicator:
            self.drop_indicator.destroy()
            self.drop_indicator = None
        if self.drop_index is None:
            self.drag_original_index = None
            return
        if self.drag_original_index is not None and self.drag_original_index < len(self.commands):
            if self.drop_index != self.drag_original_index and self.drop_index != self.drag_original_index + 1:
                cmd = self.commands.pop(self.drag_original_index)
                if self.drop_index > self.drag_original_index:
                    self.drop_index -= 1
                self.commands.insert(self.drop_index, cmd)
                self.listbox.delete(0, tk.END)
                for c in self.commands:
                    self.listbox.insert(tk.END, self.get_display_text(c))
                self.log(f"명령이 {self.drag_original_index}에서 {self.drop_index}로 이동됨.")
                new_index = self.drop_index
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(new_index)
                self.listbox.activate(new_index)
        self.drag_original_index = None
        self.drop_index = None
        
    def draw_drop_indicator(self, index):
        if self.drop_indicator:
            self.drop_indicator.destroy()
            self.drop_indicator = None
        if index < len(self.commands):
            bbox = self.listbox.bbox(index)
        elif len(self.commands) > 0:
            bbox = self.listbox.bbox(len(self.commands) - 1)
            if bbox:
                bbox = (bbox[0], bbox[1] + bbox[3], bbox[2], bbox[3])
        else:
            bbox = (0, 0, 0, 0)
        y = bbox[1] if bbox else 0
        self.drop_indicator = tk.Canvas(self.frame_list, width=self.listbox.winfo_width(),
                                         height=2, highlightthickness=0, bd=0, bg="red")
        self.drop_indicator.place(x=0, y=y)

def main():
    app = ManualMacroGUI()
    app.mainloop()

if __name__ == "__main__":
    main()

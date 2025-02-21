#!/usr/bin/env python3
import time
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from pynput import keyboard, mouse

DRAG_THRESHOLD = 5  # 拖拽开始前的最小移动像素
RECORD_WAIT_THRESHOLD = 0.1  # 事件之间的最小等待时间（秒）

# 色彩及字体设置
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
        self.title("BLOUplanet's Macro - 键盘/鼠标宏程序")
        self.geometry("800x900")
        self.resizable(False, False)
        self.configure(bg=BG_COLOR)
        
        # 宏命令相关变量
        self.commands = []  # 保存宏命令的列表
        self.macro_running = False
        self.drag_original_index = None  # 拖拽开始时的项目索引
        self.dragged_command = None      # 拖拽开始时选中的命令对象
        self.ghost = None                # 拖拽时显示的半透明影像
        self.drop_index = None           # 当前预期的放置位置
        self.drop_indicator = None       # 仅在拖拽时显示的放置指示器（红色横线）
        self._drag_start_x = None
        self._drag_start_y = None
        
        # 动作记录相关变量
        self.action_recording = False
        self.recorded_commands = []      # 通过动作记录生成的命令
        self.last_record_time = 0
        self.action_keyboard_listener = None
        self.action_mouse_listener = None

        # 快捷键相关变量 (现有宏快捷键: f2/f3, 动作记录快捷键: f4/f5)
        self.action_start_hotkey_var = tk.StringVar(value="f4")
        self.action_stop_hotkey_var = tk.StringVar(value="f5")
        
        # 点击背景时取消焦点
        self.bind_all("<Button-1>", self.clear_focus, add="+")
        
        # --- 命令列表区域 ---
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
        
        # --- 命令添加编辑区域 ---
        self.frame_editor = tk.Frame(self, bg=FRAME_BG)
        self.frame_editor.pack(padx=10, pady=5, fill=tk.X)
        tk.Label(self.frame_editor, text="命令类型:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .grid(row=0, column=0, padx=5, pady=5)
        self.command_type_var = tk.StringVar(value="键敲击")
        self.option_menu = tk.OptionMenu(self.frame_editor, self.command_type_var,
                                         "键敲击", "等待", "鼠标点击", "键长按", "鼠标长按", "鼠标滚动",
                                         command=self.update_param_fields)
        self.option_menu.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
        self.option_menu["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
        self.option_menu.grid(row=0, column=1, padx=5, pady=5)
        self.frame_params = tk.Frame(self.frame_editor, bg=FRAME_BG)
        self.frame_params.grid(row=0, column=2, padx=5, pady=5)
        self.param_entries = {}
        self.update_param_fields("键敲击")
        self.button_add = tk.Button(self.frame_editor, text="添加命令", command=self.add_command,
                                    bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                    activebackground=BUTTON_ACTIVE_BG)
        self.button_add.grid(row=0, column=3, padx=5, pady=5)
        
        # --- 控制按钮区域 ---
        self.frame_controls = tk.Frame(self, bg=FRAME_BG)
        self.frame_controls.pack(padx=10, pady=5, fill=tk.X)
        # 上部: 删除所选命令、运行宏、停止宏
        self.frame_controls_top = tk.Frame(self.frame_controls, bg=FRAME_BG)
        self.frame_controls_top.pack(fill=tk.X)
        self.button_remove = tk.Button(self.frame_controls_top, text="删除所选命令", command=self.remove_command,
                                       bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                       activebackground=BUTTON_ACTIVE_BG)
        self.button_remove.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_play = tk.Button(self.frame_controls_top, text="运行宏", command=self.play_macro,
                                     bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                     activebackground=BUTTON_ACTIVE_BG)
        self.button_play.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_stop = tk.Button(self.frame_controls_top, text="停止宏", command=self.stop_macro, state=tk.DISABLED,
                                     bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                     activebackground=BUTTON_ACTIVE_BG)
        self.button_stop.pack(side=tk.LEFT, padx=5, pady=5)
        # 下部: 保存宏、加载宏、重复次数
        self.frame_controls_bottom = tk.Frame(self.frame_controls, bg=FRAME_BG)
        self.frame_controls_bottom.pack(fill=tk.X, pady=(5,0))
        self.button_save = tk.Button(self.frame_controls_bottom, text="保存宏", command=self.save_macro,
                                     bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                     activebackground=BUTTON_ACTIVE_BG)
        self.button_save.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_load = tk.Button(self.frame_controls_bottom, text="加载宏", command=self.load_macro,
                                     bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                     activebackground=BUTTON_ACTIVE_BG)
        self.button_load.pack(side=tk.LEFT, padx=5, pady=5)
        tk.Label(self.frame_controls_bottom, text="重复次数 (0:无限):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .pack(side=tk.LEFT, padx=5, pady=5)
        self.entry_loop = tk.Entry(self.frame_controls_bottom, width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_loop.insert(0, "1")
        self.entry_loop.pack(side=tk.LEFT, padx=5, pady=5)
        
        # --- 快捷键设置区域 ---
        self.frame_hotkeys = tk.Frame(self, bg=FRAME_BG)
        self.frame_hotkeys.pack(padx=10, pady=5, fill=tk.X)
        tk.Label(self.frame_hotkeys, text="宏运行快捷键:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .pack(side=tk.LEFT, padx=5, pady=5)
        self.start_hotkey_var = tk.StringVar(value="f2")
        self.entry_start_hotkey = tk.Entry(self.frame_hotkeys, textvariable=self.start_hotkey_var,
                                           width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_start_hotkey.pack(side=tk.LEFT, padx=5, pady=5)
        tk.Label(self.frame_hotkeys, text="宏停止快捷键:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .pack(side=tk.LEFT, padx=5, pady=5)
        self.stop_hotkey_var = tk.StringVar(value="f3")
        self.entry_stop_hotkey = tk.Entry(self.frame_hotkeys, textvariable=self.stop_hotkey_var,
                                          width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_stop_hotkey.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_apply_hotkeys = tk.Button(self.frame_hotkeys, text="应用快捷键", command=self.apply_hotkeys,
                                              bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                              activebackground=BUTTON_ACTIVE_BG)
        self.button_apply_hotkeys.pack(side=tk.LEFT, padx=5, pady=5)
        
        # --- 动作记录区域 ---
        self.frame_action_record = tk.Frame(self, bg=FRAME_BG)
        self.frame_action_record.pack(padx=10, pady=5, fill=tk.X)
        tk.Label(self.frame_action_record, text="动作记录快捷键 (开始/结束):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .pack(side=tk.LEFT, padx=5, pady=5)
        self.entry_action_start_hotkey = tk.Entry(self.frame_action_record, textvariable=self.action_start_hotkey_var,
                                                  width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_action_start_hotkey.pack(side=tk.LEFT, padx=5, pady=5)
        self.entry_action_stop_hotkey = tk.Entry(self.frame_action_record, textvariable=self.action_stop_hotkey_var,
                                                 width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_action_stop_hotkey.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_toggle_recording = tk.Button(self.frame_action_record, text="开始记录动作", command=self.toggle_action_recording,
                                                 bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                                 activebackground=BUTTON_ACTIVE_BG)
        self.button_toggle_recording.pack(side=tk.LEFT, padx=5, pady=5)
        
        # --- 日志输出区域 ---
        self.text_log = tk.Text(self, height=10, width=90, state=tk.NORMAL, bg=LISTBOX_BG, fg=LISTBOX_FG,
                                font=FONT, relief=tk.FLAT)
        self.text_log.pack(padx=10, pady=5)
        
        self.keyboard_controller = keyboard.Controller()
        self.mouse_controller = mouse.Controller()
        
        self.hotkey_listener = None
        self.after(100, self.start_hotkey_listener)
        
    # 取消焦点
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
        if command_type == "键敲击":
            tk.Label(self.frame_params, text="键:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=2)
            entry_key = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_key.grid(row=0, column=1, padx=5, pady=2)
            def on_key_press(event):
                entry_key.delete(0, tk.END)
                entry_key.insert(0, event.keysym)
                return "break"
            entry_key.bind("<Key>", on_key_press)
            self.param_entries["key"] = entry_key
            tk.Label(self.frame_params, text="重复次数:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=2)
            entry_repeat = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_repeat.insert(0, "1")
            entry_repeat.grid(row=1, column=1, padx=5, pady=2)
            self.param_entries["repeat"] = entry_repeat
        elif command_type == "等待":
            tk.Label(self.frame_params, text="等待时间（秒）:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=2)
            entry_duration = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.grid(row=0, column=1, padx=5, pady=2)
            self.param_entries["duration"] = entry_duration
        elif command_type == "鼠标点击":
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
            tk.Label(self.frame_params, text="按钮:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=4, padx=5, pady=2)
            self.mouse_button_var = tk.StringVar(value="left")
            option_button = tk.OptionMenu(self.frame_params, self.mouse_button_var, "left", "right", "middle")
            option_button.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
            option_button["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
            option_button.grid(row=0, column=5, padx=5, pady=2)
            self.param_entries["button"] = self.mouse_button_var
            # “记录鼠标位置”按钮放在新的一行
            self.button_record_mouse = tk.Button(self.frame_params, text="记录鼠标位置", command=self.record_mouse_position,
                                                 bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                                 activebackground=BUTTON_ACTIVE_BG)
            self.button_record_mouse.grid(row=1, column=0, columnspan=6, padx=5, pady=2, sticky="w")
        elif command_type == "键长按":
            tk.Label(self.frame_params, text="键:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=2)
            entry_key = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_key.grid(row=0, column=1, padx=5, pady=2)
            def on_key_press(event):
                entry_key.delete(0, tk.END)
                entry_key.insert(0, event.keysym)
                return "break"
            entry_key.bind("<Key>", on_key_press)
            self.param_entries["key"] = entry_key
            tk.Label(self.frame_params, text="按住时间（秒）:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=2)
            entry_duration = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.insert(0, "1")
            entry_duration.grid(row=1, column=1, padx=5, pady=2)
            self.param_entries["duration"] = entry_duration
        elif command_type == "鼠标长按":
            # 第一行: X, Y, 按钮
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
            tk.Label(self.frame_params, text="按钮:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=4, padx=5, pady=2)
            self.mouse_button_var = tk.StringVar(value="left")
            option_button = tk.OptionMenu(self.frame_params, self.mouse_button_var, "left", "right", "middle")
            option_button.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
            option_button["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
            option_button.grid(row=0, column=5, padx=5, pady=2)
            self.param_entries["button"] = self.mouse_button_var
            # 第二行: 按住时间及记录鼠标位置按钮
            tk.Label(self.frame_params, text="按住时间（秒）:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=2)
            entry_duration = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.insert(0, "1")
            entry_duration.grid(row=1, column=1, padx=5, pady=2)
            self.param_entries["duration"] = entry_duration
            self.button_record_mouse = tk.Button(self.frame_params, text="记录鼠标位置", command=self.record_mouse_position,
                                                 bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                                 activebackground=BUTTON_ACTIVE_BG)
            self.button_record_mouse.grid(row=1, column=2, columnspan=4, padx=5, pady=2, sticky="w")
        elif command_type == "鼠标滚动":
            tk.Label(self.frame_params, text="水平滚动:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=2)
            entry_dx = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_dx.insert(0, "0")
            entry_dx.grid(row=0, column=1, padx=5, pady=2)
            self.param_entries["dx"] = entry_dx
            tk.Label(self.frame_params, text="垂直滚动:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=2, padx=5, pady=2)
            entry_dy = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_dy.insert(0, "0")
            entry_dy.grid(row=0, column=3, padx=5, pady=2)
            self.param_entries["dy"] = entry_dy

    def record_mouse_position(self):
        self.button_record_mouse.config(text="单击左键记录鼠标位置完成", state=tk.DISABLED)
        self.log("等待记录鼠标位置... 单击左键完成记录.")
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
        self.log(f"记录到鼠标位置: ({int(x)}, {int(y)})")
        self.button_record_mouse.config(text="记录鼠标位置", state=tk.NORMAL)
        
    def add_command(self):
        command_type = self.command_type_var.get()
        if command_type == "键敲击":
            key = self.param_entries["key"].get().strip()
            if key == "":
                messagebox.showerror("错误", "请输入键值.")
                return
            try:
                repeat = int(self.param_entries["repeat"].get().strip())
            except ValueError:
                messagebox.showerror("错误", "请输入有效的重复次数.")
                return
            cmd = {"command": "key_tap", "key": key, "repeat": repeat}
            display_text = f"键敲击: {key} x {repeat}次"
        elif command_type == "等待":
            try:
                duration = float(self.param_entries["duration"].get().strip())
            except ValueError:
                messagebox.showerror("错误", "请输入有效的等待时间.")
                return
            cmd = {"command": "wait", "duration": duration}
            display_text = f"等待: {duration}秒"
        elif command_type == "鼠标点击":
            x_str = self.param_entries["x"].get().strip()
            y_str = self.param_entries["y"].get().strip()
            if not x_str or not y_str:
                messagebox.showerror("错误", "请输入X, Y值.")
                return
            try:
                x = int(x_str)
                y = int(y_str)
            except ValueError:
                messagebox.showerror("错误", "请输入有效的整数X, Y值.")
                return
            button = self.param_entries["button"].get()
            cmd = {"command": "mouse_click", "x": x, "y": y, "button": button}
            display_text = f"鼠标点击: ({x}, {y}), 按钮: {button}"
        elif command_type == "键长按":
            key = self.param_entries["key"].get().strip()
            if key == "":
                messagebox.showerror("错误", "请输入键值.")
                return
            try:
                duration = float(self.param_entries["duration"].get().strip())
            except ValueError:
                messagebox.showerror("错误", "请输入有效的按住时间.")
                return
            cmd = {"command": "key_hold", "key": key, "duration": duration}
            display_text = f"键长按: {key} (按住时间: {duration}秒)"
        elif command_type == "鼠标长按":
            x_str = self.param_entries["x"].get().strip()
            y_str = self.param_entries["y"].get().strip()
            if not x_str or not y_str:
                messagebox.showerror("错误", "请输入X, Y值.")
                return
            try:
                x = int(x_str)
                y = int(y_str)
            except ValueError:
                messagebox.showerror("错误", "请输入有效的整数X, Y值.")
                return
            button = self.param_entries["button"].get()
            try:
                duration = float(self.param_entries["duration"].get().strip())
            except ValueError:
                messagebox.showerror("错误", "请输入有效的按住时间.")
                return
            cmd = {"command": "mouse_hold", "x": x, "y": y, "button": button, "duration": duration}
            display_text = f"鼠标长按: ({x}, {y}), 按钮: {button} (按住时间: {duration}秒)"
        elif command_type == "鼠标滚动":
            try:
                dx = int(self.param_entries["dx"].get().strip())
                dy = int(self.param_entries["dy"].get().strip())
            except ValueError:
                messagebox.showerror("错误", "请输入有效的滚动值.")
                return
            cmd = {"command": "mouse_scroll", "dx": dx, "dy": dy}
            display_text = f"鼠标滚动: 水平 {dx}, 垂直 {dy}"
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
        self.log("命令已添加: " + display_text)
        
    def remove_command(self):
        selected = self.listbox.curselection()
        if not selected:
            messagebox.showerror("错误", "请选择要删除的命令.")
            return
        index = selected[0]
        self.listbox.delete(index)
        removed = self.commands.pop(index)
        self.log("命令已删除: " + str(removed))
        
    def play_macro(self):
        if not self.commands:
            messagebox.showinfo("信息", "没有要执行的命令.")
            return
        try:
            loop_count = int(self.entry_loop.get().strip())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的重复次数.")
            return
        self.log("宏执行开始.")
        self.macro_running = True
        self.button_stop.config(state=tk.NORMAL)
        thread = threading.Thread(target=self.execute_macro, args=(loop_count,))
        thread.daemon = True
        thread.start()
        
    def execute_macro(self, loop_count):
        iteration = 0
        while self.macro_running and (loop_count == 0 or iteration < loop_count):
            self.log(f"第 {iteration+1} 次循环开始.")
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
                        self.log(f"执行键敲击: {key}")
                        time.sleep(0.05)
                elif cmd["command"] == "key_hold":
                    key = cmd["key"]
                    duration = cmd["duration"]
                    try:
                        key_obj = getattr(keyboard.Key, key.lower())
                    except AttributeError:
                        key_obj = key
                    self.keyboard_controller.press(key_obj)
                    self.log(f"开始键长按: {key}")
                    time.sleep(duration)
                    self.keyboard_controller.release(key_obj)
                    self.log(f"结束键长按: {key}")
                elif cmd["command"] == "wait":
                    duration = cmd["duration"]
                    self.log(f"开始等待: {duration}秒")
                    time.sleep(duration)
                    self.log("等待结束")
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
                    self.log(f"执行鼠标点击: ({x}, {y}), 按钮: {button_str}")
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
                    self.log(f"开始鼠标长按: ({x}, {y}), 按钮: {button_str}")
                    time.sleep(cmd["duration"])
                    self.mouse_controller.release(btn)
                    self.log(f"结束鼠标长按: ({x}, {y}), 按钮: {button_str}")
                elif cmd["command"] == "mouse_scroll":
                    dx = cmd["dx"]
                    dy = cmd["dy"]
                    self.mouse_controller.scroll(dx, dy)
                    self.log(f"执行鼠标滚动: 水平 {dx}, 垂直 {dy}")
                time.sleep(0.1)
            iteration += 1
            self.log(f"第 {iteration} 次循环完成.")
        self.log("宏执行完成.")
        self.macro_running = False
        self.button_stop.config(state=tk.DISABLED)
        
    def stop_macro(self):
        self.macro_running = False
        self.log("请求停止宏执行.")
        
    def save_macro(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", 
                                                 filetypes=[("JSON files", "*.json")])
        if not file_path:
            return
        try:
            with open(file_path, "w") as f:
                json.dump(self.commands, f, indent=4)
            self.log("宏已保存: " + file_path)
        except Exception as e:
            messagebox.showerror("错误", "宏保存失败: " + str(e))
        
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
                self.log("加载命令: " + disp)
            self.log("宏加载完成: " + file_path)
        except Exception as e:
            messagebox.showerror("错误", "宏加载失败: " + str(e))
            
    def get_display_text(self, cmd):
        try:
            if cmd.get("command") == "key_tap":
                return f"键敲击: {cmd.get('key', '')} x {cmd.get('repeat', 1)}次"
            elif cmd.get("command") == "key_hold":
                return f"键长按: {cmd.get('key', '')} (按住时间: {cmd.get('duration', 0)}秒)"
            elif cmd.get("command") == "wait":
                return f"等待: {cmd.get('duration', 0)}秒"
            elif cmd.get("command") == "mouse_click":
                return f"鼠标点击: ({cmd.get('x', 0)}, {cmd.get('y', 0)}), 按钮: {cmd.get('button', '')}"
            elif cmd.get("command") == "mouse_hold":
                return f"鼠标长按: ({cmd.get('x', 0)}, {cmd.get('y', 0)}), 按钮: {cmd.get('button', '')} (按住时间: {cmd.get('duration', 0)}秒)"
            elif cmd.get("command") == "mouse_scroll":
                return f"鼠标滚动: 水平 {cmd.get('dx',0)}, 垂直 {cmd.get('dy',0)}"
            else:
                return str(cmd)
        except Exception as e:
            self.log("命令显示错误: " + str(e))
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
        self.log("快捷键监听器已启动, 宏: 开始=" + self.format_hotkey(self.start_hotkey_var.get()) +
                 " 停止=" + self.format_hotkey(self.stop_hotkey_var.get()) +
                 "; 动作记录: 开始=" + self.format_hotkey(self.action_start_hotkey_var.get()) +
                 " 结束=" + self.format_hotkey(self.action_stop_hotkey_var.get()))
        
    def stop_hotkey_listener(self):
        if self.hotkey_listener:
            self.hotkey_listener.stop()
            self.hotkey_listener = None
            self.log("快捷键监听器已停止.")
        
    def apply_hotkeys(self):
        self.stop_hotkey_listener()
        self.start_hotkey_listener()
        self.log("快捷键已应用.")
        
    def on_hotkey_start(self):
        self.log("通过快捷键请求宏执行.")
        if not self.macro_running:
            self.play_macro()
        
    def on_hotkey_stop(self):
        self.log("通过快捷键请求宏停止.")
        self.stop_macro()
        
    def start_action_recording(self):
        if self.action_recording:
            return
        self.action_recording = True
        self.recorded_commands = []
        self.last_record_time = time.time()
        self.log("动作记录已开始.")
        self.action_keyboard_listener = keyboard.Listener(on_release=self.action_on_key_release)
        self.action_mouse_listener = mouse.Listener(on_click=self.action_on_mouse_click)
        self.action_keyboard_listener.start()
        self.action_mouse_listener.start()
        self.button_toggle_recording.config(text="停止记录动作")
        
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
        self.log("动作记录已结束. 将记录的动作添加到命令列表.")
        for cmd in self.recorded_commands:
            self.commands.append(cmd)
            self.listbox.insert(tk.END, self.get_display_text(cmd))
        self.button_toggle_recording.config(text="开始记录动作")
        
    def action_on_key_release(self, key):
        if not self.action_recording:
            return
        now = time.time()
        dt = now - self.last_record_time
        if dt > RECORD_WAIT_THRESHOLD:
            wait_cmd = {"command": "wait", "duration": round(dt, 2)}
            self.recorded_commands.append(wait_cmd)
            self.log("记录等待: {}秒".format(round(dt, 2)))
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
        self.log("记录键敲击: {}".format(k))
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
            self.log("记录等待: {}秒".format(round(dt, 2)))
        mouse_cmd = {"command": "mouse_click", "x": int(x), "y": int(y), "button": "left"}
        self.recorded_commands.append(mouse_cmd)
        self.log("记录鼠标点击: ({}, {})".format(int(x), int(y)))
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
        edit_win.title("修改命令")
        edit_win.configure(bg=BG_COLOR)
        edit_win.wait_visibility()
        edit_win.grab_set()
        def save_changes():
            if cmd["command"] == "key_tap":
                new_key = entry_key.get().strip()
                try:
                    new_repeat = int(entry_repeat.get().strip())
                except ValueError:
                    messagebox.showerror("错误", "请输入有效的重复次数.", parent=edit_win)
                    return
                if new_key == "":
                    messagebox.showerror("错误", "请输入键值.", parent=edit_win)
                    return
                cmd["key"] = new_key
                cmd["repeat"] = new_repeat
            elif cmd["command"] == "key_hold":
                new_key = entry_key.get().strip()
                try:
                    new_duration = float(entry_duration.get().strip())
                except ValueError:
                    messagebox.showerror("错误", "请输入有效的按住时间.", parent=edit_win)
                    return
                if new_key == "":
                    messagebox.showerror("错误", "请输入键值.", parent=edit_win)
                    return
                cmd["key"] = new_key
                cmd["duration"] = new_duration
            elif cmd["command"] == "wait":
                try:
                    new_duration = float(entry_duration.get().strip())
                except ValueError:
                    messagebox.showerror("错误", "请输入有效的等待时间.", parent=edit_win)
                    return
                cmd["duration"] = new_duration
            elif cmd["command"] == "mouse_click":
                try:
                    new_x = int(entry_x.get().strip())
                    new_y = int(entry_y.get().strip())
                except ValueError:
                    messagebox.showerror("错误", "请输入有效的整数X, Y值.", parent=edit_win)
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
                    messagebox.showerror("错误", "请输入有效的整数X, Y值.", parent=edit_win)
                    return
                try:
                    new_duration = float(entry_duration.get().strip())
                except ValueError:
                    messagebox.showerror("错误", "请输入有效的按住时间.", parent=edit_win)
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
                    messagebox.showerror("错误", "请输入有效的滚动值.", parent=edit_win)
                    return
                cmd["dx"] = new_dx
                cmd["dy"] = new_dy
            self.listbox.delete(index)
            self.listbox.insert(index, self.get_display_text(cmd))
            self.log("命令已修改: " + self.get_display_text(cmd))
            edit_win.destroy()
        if cmd["command"] == "key_tap":
            tk.Label(edit_win, text="键:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=5)
            entry_key = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_key.insert(0, cmd["key"])
            entry_key.grid(row=0, column=1, padx=5, pady=5)
            def on_key_press(event):
                entry_key.delete(0, tk.END)
                entry_key.insert(0, event.keysym)
                return "break"
            entry_key.bind("<Key>", on_key_press)
            tk.Label(edit_win, text="重复次数:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=5)
            entry_repeat = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_repeat.insert(0, str(cmd.get("repeat", 1)))
            entry_repeat.grid(row=1, column=1, padx=5, pady=5)
        elif cmd["command"] == "key_hold":
            tk.Label(edit_win, text="键:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=5)
            entry_key = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_key.insert(0, cmd["key"])
            entry_key.grid(row=0, column=1, padx=5, pady=5)
            def on_key_press(event):
                entry_key.delete(0, tk.END)
                entry_key.insert(0, event.keysym)
                return "break"
            entry_key.bind("<Key>", on_key_press)
            tk.Label(edit_win, text="按住时间（秒）:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=5)
            entry_duration = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.insert(0, str(cmd.get("duration", 1)))
            entry_duration.grid(row=1, column=1, padx=5, pady=5)
        elif cmd["command"] == "wait":
            tk.Label(edit_win, text="等待时间（秒）:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
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
            tk.Label(edit_win, text="按钮:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=4, padx=5, pady=5)
            var_button = tk.StringVar(value=cmd["button"])
            option_button = tk.OptionMenu(edit_win, var_button, "left", "right", "middle")
            option_button.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
            option_button["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
            option_button.grid(row=0, column=5, padx=5, pady=5)
            def record_mouse_edit():
                record_button_edit.config(text="单击左键记录鼠标位置完成", state=tk.DISABLED)
                def on_click(x, y, button, pressed):
                    if button == mouse.Button.left and pressed:
                        entry_x.delete(0, tk.END)
                        entry_x.insert(0, str(int(x)))
                        entry_y.delete(0, tk.END)
                        entry_y.insert(0, str(int(y)))
                        edit_win.after(0, lambda: self.log(f"在编辑窗口中记录鼠标位置: ({int(x)}, {int(y)})"))
                        record_button_edit.config(text="记录鼠标位置", state=tk.NORMAL)
                        return False
                listener = mouse.Listener(on_click=on_click)
                listener.start()
            record_button_edit = tk.Button(edit_win, text="记录鼠标位置", command=record_mouse_edit,
                                           bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                           activebackground=BUTTON_ACTIVE_BG)
            record_button_edit.grid(row=1, column=0, columnspan=6, padx=5, pady=5, sticky="w")
        elif cmd["command"] == "mouse_hold":
            # 第一行: X, Y, 按钮
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
            tk.Label(edit_win, text="按钮:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=4, padx=5, pady=5)
            var_button = tk.StringVar(value=cmd["button"])
            option_button = tk.OptionMenu(edit_win, var_button, "left", "right", "middle")
            option_button.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
            option_button["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
            option_button.grid(row=0, column=5, padx=5, pady=5)
            tk.Label(edit_win, text="按住时间（秒）:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=5)
            entry_duration = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.insert(0, str(cmd.get("duration", 1)))
            entry_duration.grid(row=1, column=1, padx=5, pady=5)
            def record_mouse_edit():
                record_button_edit.config(text="单击左键记录鼠标位置完成", state=tk.DISABLED)
                def on_click(x, y, button, pressed):
                    if button == mouse.Button.left and pressed:
                        entry_x.delete(0, tk.END)
                        entry_x.insert(0, str(int(x)))
                        entry_y.delete(0, tk.END)
                        entry_y.insert(0, str(int(y)))
                        edit_win.after(0, lambda: self.log(f"在编辑窗口中记录鼠标位置: ({int(x)}, {int(y)})"))
                        record_button_edit.config(text="记录鼠标位置", state=tk.NORMAL)
                        return False
                listener = mouse.Listener(on_click=on_click)
                listener.start()
            record_button_edit = tk.Button(edit_win, text="记录鼠标位置", command=record_mouse_edit,
                                           bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                           activebackground=BUTTON_ACTIVE_BG)
            record_button_edit.grid(row=1, column=2, columnspan=4, padx=5, pady=5, sticky="w")
        elif cmd["command"] == "mouse_scroll":
            tk.Label(edit_win, text="水平滚动:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=5)
            entry_dx = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_dx.insert(0, str(cmd["dx"]))
            entry_dx.grid(row=0, column=1, padx=5, pady=5)
            tk.Label(edit_win, text="垂直滚动:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=2, padx=5, pady=5)
            entry_dy = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_dy.insert(0, str(cmd["dy"]))
            entry_dy.grid(row=0, column=3, padx=5, pady=5)
        tk.Button(edit_win, text="保存", command=save_changes,
                  bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT, activebackground=BUTTON_ACTIVE_BG)\
            .grid(row=10, column=0, padx=5, pady=10)
        tk.Button(edit_win, text="取消", command=edit_win.destroy,
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
                self.log(f"命令从 {self.drag_original_index} 移动到 {self.drop_index}.")
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

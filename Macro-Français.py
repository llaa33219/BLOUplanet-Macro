#!/usr/bin/env python3
import time
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from pynput import keyboard, mouse

DRAG_THRESHOLD = 5  # Seuil de déplacement minimum avant de commencer le glisser-déposer
RECORD_WAIT_THRESHOLD = 0.1  # Temps d'attente minimum entre les événements (sec)

# Couleurs et police
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
        self.title("BLOUplanet's Macro - Programme de macro pour clavier/souris")
        self.geometry("800x900")
        self.resizable(False, False)
        self.configure(bg=BG_COLOR)
        
        # Variables liées aux commandes de macro
        self.commands = []  # Liste pour stocker les commandes de macro
        self.macro_running = False
        self.drag_original_index = None  # Index de l'élément au début du glisser
        self.dragged_command = None      # Commande sélectionnée lors du début du glisser
        self.ghost = None                # Fenêtre fantôme semi-transparente pendant le glisser
        self.drop_index = None           # Position d'insertion prévue lors du dépôt
        self.drop_indicator = None       # Indicateur de dépôt (ligne rouge) visible uniquement pendant le glisser
        self._drag_start_x = None
        self._drag_start_y = None
        
        # Variables liées à l'enregistrement des actions
        self.action_recording = False
        self.recorded_commands = []      # Commandes générées par l'enregistrement des actions
        self.last_record_time = 0
        self.action_keyboard_listener = None
        self.action_mouse_listener = None

        # Variables pour les raccourcis clavier (raccourcis de macro existants : f2/f3, enregistrement des actions : f4/f5)
        self.action_start_hotkey_var = tk.StringVar(value="f4")
        self.action_stop_hotkey_var = tk.StringVar(value="f5")
        
        # Retrait du focus lors du clic sur le fond
        self.bind_all("<Button-1>", self.clear_focus, add="+")
        
        # --- Zone de la liste des commandes ---
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
        
        # --- Zone de l'éditeur pour ajouter une commande ---
        self.frame_editor = tk.Frame(self, bg=FRAME_BG)
        self.frame_editor.pack(padx=10, pady=5, fill=tk.X)
        tk.Label(self.frame_editor, text="Type de commande:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .grid(row=0, column=0, padx=5, pady=5)
        self.command_type_var = tk.StringVar(value="Appui de touche")
        self.option_menu = tk.OptionMenu(self.frame_editor, self.command_type_var,
                                         "Appui de touche", "Attente", "Clic de souris", "Maintien de touche", "Maintien de clic", "Défilement de souris",
                                         command=self.update_param_fields)
        self.option_menu.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
        self.option_menu["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
        self.option_menu.grid(row=0, column=1, padx=5, pady=5)
        self.frame_params = tk.Frame(self.frame_editor, bg=FRAME_BG)
        self.frame_params.grid(row=0, column=2, padx=5, pady=5)
        self.param_entries = {}
        self.update_param_fields("Appui de touche")
        self.button_add = tk.Button(self.frame_editor, text="Ajouter commande", command=self.add_command,
                                    bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                    activebackground=BUTTON_ACTIVE_BG)
        self.button_add.grid(row=0, column=3, padx=5, pady=5)
        
        # --- Zone des boutons de contrôle ---
        self.frame_controls = tk.Frame(self, bg=FRAME_BG)
        self.frame_controls.pack(padx=10, pady=5, fill=tk.X)
        # Partie supérieure : Supprimer commande sélectionnée, Exécuter macro, Arrêter macro
        self.frame_controls_top = tk.Frame(self.frame_controls, bg=FRAME_BG)
        self.frame_controls_top.pack(fill=tk.X)
        self.button_remove = tk.Button(self.frame_controls_top, text="Supprimer commande sélectionnée", command=self.remove_command,
                                       bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                       activebackground=BUTTON_ACTIVE_BG)
        self.button_remove.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_play = tk.Button(self.frame_controls_top, text="Exécuter macro", command=self.play_macro,
                                     bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                     activebackground=BUTTON_ACTIVE_BG)
        self.button_play.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_stop = tk.Button(self.frame_controls_top, text="Arrêter macro", command=self.stop_macro, state=tk.DISABLED,
                                     bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                     activebackground=BUTTON_ACTIVE_BG)
        self.button_stop.pack(side=tk.LEFT, padx=5, pady=5)
        # Partie inférieure : Enregistrer macro, Charger macro, Nombre de répétitions
        self.frame_controls_bottom = tk.Frame(self.frame_controls, bg=FRAME_BG)
        self.frame_controls_bottom.pack(fill=tk.X, pady=(5,0))
        self.button_save = tk.Button(self.frame_controls_bottom, text="Enregistrer macro", command=self.save_macro,
                                     bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                     activebackground=BUTTON_ACTIVE_BG)
        self.button_save.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_load = tk.Button(self.frame_controls_bottom, text="Charger macro", command=self.load_macro,
                                     bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                     activebackground=BUTTON_ACTIVE_BG)
        self.button_load.pack(side=tk.LEFT, padx=5, pady=5)
        tk.Label(self.frame_controls_bottom, text="Nombre de répétitions (0: infini):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .pack(side=tk.LEFT, padx=5, pady=5)
        self.entry_loop = tk.Entry(self.frame_controls_bottom, width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_loop.insert(0, "1")
        self.entry_loop.pack(side=tk.LEFT, padx=5, pady=5)
        
        # --- Zone de configuration des raccourcis existants ---
        self.frame_hotkeys = tk.Frame(self, bg=FRAME_BG)
        self.frame_hotkeys.pack(padx=10, pady=5, fill=tk.X)
        tk.Label(self.frame_hotkeys, text="Raccourci pour exécuter la macro:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .pack(side=tk.LEFT, padx=5, pady=5)
        self.start_hotkey_var = tk.StringVar(value="f2")
        self.entry_start_hotkey = tk.Entry(self.frame_hotkeys, textvariable=self.start_hotkey_var,
                                           width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_start_hotkey.pack(side=tk.LEFT, padx=5, pady=5)
        tk.Label(self.frame_hotkeys, text="Raccourci pour arrêter la macro:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .pack(side=tk.LEFT, padx=5, pady=5)
        self.stop_hotkey_var = tk.StringVar(value="f3")
        self.entry_stop_hotkey = tk.Entry(self.frame_hotkeys, textvariable=self.stop_hotkey_var,
                                          width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_stop_hotkey.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_apply_hotkeys = tk.Button(self.frame_hotkeys, text="Appliquer raccourcis", command=self.apply_hotkeys,
                                              bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                              activebackground=BUTTON_ACTIVE_BG)
        self.button_apply_hotkeys.pack(side=tk.LEFT, padx=5, pady=5)
        
        # --- Zone d'enregistrement des actions ---
        self.frame_action_record = tk.Frame(self, bg=FRAME_BG)
        self.frame_action_record.pack(padx=10, pady=5, fill=tk.X)
        tk.Label(self.frame_action_record, text="Raccourci d'enregistrement des actions (Démarrer/Arrêter):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
            .pack(side=tk.LEFT, padx=5, pady=5)
        self.entry_action_start_hotkey = tk.Entry(self.frame_action_record, textvariable=self.action_start_hotkey_var,
                                                  width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_action_start_hotkey.pack(side=tk.LEFT, padx=5, pady=5)
        self.entry_action_stop_hotkey = tk.Entry(self.frame_action_record, textvariable=self.action_stop_hotkey_var,
                                                 width=5, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
        self.entry_action_stop_hotkey.pack(side=tk.LEFT, padx=5, pady=5)
        self.button_toggle_recording = tk.Button(self.frame_action_record, text="Démarrer enregistrement", command=self.toggle_action_recording,
                                                 bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                                 activebackground=BUTTON_ACTIVE_BG)
        self.button_toggle_recording.pack(side=tk.LEFT, padx=5, pady=5)
        
        # --- Zone d'affichage des logs ---
        self.text_log = tk.Text(self, height=10, width=90, state=tk.NORMAL, bg=LISTBOX_BG, fg=LISTBOX_FG,
                                font=FONT, relief=tk.FLAT)
        self.text_log.pack(padx=10, pady=5)
        
        self.keyboard_controller = keyboard.Controller()
        self.mouse_controller = mouse.Controller()
        
        self.hotkey_listener = None
        self.after(100, self.start_hotkey_listener)
        
    # Retirer le focus
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
        if command_type == "Appui de touche":
            tk.Label(self.frame_params, text="Touche:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=2)
            entry_key = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_key.grid(row=0, column=1, padx=5, pady=2)
            def on_key_press(event):
                entry_key.delete(0, tk.END)
                entry_key.insert(0, event.keysym)
                return "break"
            entry_key.bind("<Key>", on_key_press)
            self.param_entries["key"] = entry_key
            tk.Label(self.frame_params, text="Nombre de répétitions:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=2)
            entry_repeat = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_repeat.insert(0, "1")
            entry_repeat.grid(row=1, column=1, padx=5, pady=2)
            self.param_entries["repeat"] = entry_repeat
        elif command_type == "Attente":
            tk.Label(self.frame_params, text="Temps d'attente (sec):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=2)
            entry_duration = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.grid(row=0, column=1, padx=5, pady=2)
            self.param_entries["duration"] = entry_duration
        elif command_type == "Clic de souris":
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
            tk.Label(self.frame_params, text="Bouton:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=4, padx=5, pady=2)
            self.mouse_button_var = tk.StringVar(value="left")
            option_button = tk.OptionMenu(self.frame_params, self.mouse_button_var, "left", "right", "middle")
            option_button.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
            option_button["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
            option_button.grid(row=0, column=5, padx=5, pady=2)
            self.param_entries["button"] = self.mouse_button_var
            self.button_record_mouse = tk.Button(self.frame_params, text="Enregistrer position de la souris", command=self.record_mouse_position,
                                                 bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                                 activebackground=BUTTON_ACTIVE_BG)
            self.button_record_mouse.grid(row=1, column=0, columnspan=6, padx=5, pady=2, sticky="w")
        elif command_type == "Maintien de touche":
            tk.Label(self.frame_params, text="Touche:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=2)
            entry_key = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_key.grid(row=0, column=1, padx=5, pady=2)
            def on_key_press(event):
                entry_key.delete(0, tk.END)
                entry_key.insert(0, event.keysym)
                return "break"
            entry_key.bind("<Key>", on_key_press)
            self.param_entries["key"] = entry_key
            tk.Label(self.frame_params, text="Durée (sec):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=2)
            entry_duration = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.insert(0, "1")
            entry_duration.grid(row=1, column=1, padx=5, pady=2)
            self.param_entries["duration"] = entry_duration
        elif command_type == "Maintien de clic":
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
            tk.Label(self.frame_params, text="Bouton:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=4, padx=5, pady=2)
            self.mouse_button_var = tk.StringVar(value="left")
            option_button = tk.OptionMenu(self.frame_params, self.mouse_button_var, "left", "right", "middle")
            option_button.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
            option_button["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
            option_button.grid(row=0, column=5, padx=5, pady=2)
            self.param_entries["button"] = self.mouse_button_var
            tk.Label(self.frame_params, text="Durée (sec):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=2)
            entry_duration = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.insert(0, "1")
            entry_duration.grid(row=1, column=1, padx=5, pady=2)
            self.param_entries["duration"] = entry_duration
            self.button_record_mouse = tk.Button(self.frame_params, text="Enregistrer position de la souris", command=self.record_mouse_position,
                                                 bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                                 activebackground=BUTTON_ACTIVE_BG)
            self.button_record_mouse.grid(row=1, column=2, columnspan=4, padx=5, pady=2, sticky="w")
        elif command_type == "Défilement de souris":
            tk.Label(self.frame_params, text="Défilement horizontal:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=2)
            entry_dx = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_dx.insert(0, "0")
            entry_dx.grid(row=0, column=1, padx=5, pady=2)
            self.param_entries["dx"] = entry_dx
            tk.Label(self.frame_params, text="Défilement vertical:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=2, padx=5, pady=2)
            entry_dy = tk.Entry(self.frame_params, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_dy.insert(0, "0")
            entry_dy.grid(row=0, column=3, padx=5, pady=2)
            self.param_entries["dy"] = entry_dy

    def record_mouse_position(self):
        self.button_record_mouse.config(text="Clique gauche pour enregistrer la position", state=tk.DISABLED)
        self.log("En attente de l'enregistrement de la position de la souris... Cliquez avec le bouton gauche pour terminer.")
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
        self.log(f"Position de la souris enregistrée: ({int(x)}, {int(y)})")
        self.button_record_mouse.config(text="Enregistrer position de la souris", state=tk.NORMAL)
        
    def add_command(self):
        command_type = self.command_type_var.get()
        if command_type == "Appui de touche":
            key = self.param_entries["key"].get().strip()
            if key == "":
                messagebox.showerror("Erreur", "Veuillez saisir une touche.")
                return
            try:
                repeat = int(self.param_entries["repeat"].get().strip())
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez saisir un nombre de répétitions valide.")
                return
            cmd = {"command": "key_tap", "key": key, "repeat": repeat}
            display_text = f"Appui de touche: {key} x {repeat} fois"
        elif command_type == "Attente":
            try:
                duration = float(self.param_entries["duration"].get().strip())
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez saisir un temps d'attente valide.")
                return
            cmd = {"command": "wait", "duration": duration}
            display_text = f"Attente: {duration} sec"
        elif command_type == "Clic de souris":
            x_str = self.param_entries["x"].get().strip()
            y_str = self.param_entries["y"].get().strip()
            if not x_str or not y_str:
                messagebox.showerror("Erreur", "Veuillez saisir les valeurs X et Y.")
                return
            try:
                x = int(x_str)
                y = int(y_str)
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez saisir des valeurs entières valides pour X et Y.")
                return
            button = self.param_entries["button"].get()
            cmd = {"command": "mouse_click", "x": x, "y": y, "button": button}
            display_text = f"Clic de souris: ({x}, {y}), bouton: {button}"
        elif command_type == "Maintien de touche":
            key = self.param_entries["key"].get().strip()
            if key == "":
                messagebox.showerror("Erreur", "Veuillez saisir une touche.")
                return
            try:
                duration = float(self.param_entries["duration"].get().strip())
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez saisir une durée valide.")
                return
            cmd = {"command": "key_hold", "key": key, "duration": duration}
            display_text = f"Maintien de touche: {key} (durée: {duration} sec)"
        elif command_type == "Maintien de clic":
            x_str = self.param_entries["x"].get().strip()
            y_str = self.param_entries["y"].get().strip()
            if not x_str or not y_str:
                messagebox.showerror("Erreur", "Veuillez saisir les valeurs X et Y.")
                return
            try:
                x = int(x_str)
                y = int(y_str)
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez saisir des valeurs entières valides pour X et Y.")
                return
            button = self.param_entries["button"].get()
            try:
                duration = float(self.param_entries["duration"].get().strip())
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez saisir une durée valide.")
                return
            cmd = {"command": "mouse_hold", "x": x, "y": y, "button": button, "duration": duration}
            display_text = f"Maintien de clic: ({x}, {y}), bouton: {button} (durée: {duration} sec)"
        elif command_type == "Défilement de souris":
            try:
                dx = int(self.param_entries["dx"].get().strip())
                dy = int(self.param_entries["dy"].get().strip())
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez saisir des valeurs de défilement valides.")
                return
            cmd = {"command": "mouse_scroll", "dx": dx, "dy": dy}
            display_text = f"Défilement de souris: horizontal {dx}, vertical {dy}"
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
        self.log("Commande ajoutée: " + display_text)
        
    def remove_command(self):
        selected = self.listbox.curselection()
        if not selected:
            messagebox.showerror("Erreur", "Veuillez sélectionner une commande à supprimer.")
            return
        index = selected[0]
        self.listbox.delete(index)
        removed = self.commands.pop(index)
        self.log("Commande supprimée: " + str(removed))
        
    def play_macro(self):
        if not self.commands:
            messagebox.showinfo("Information", "Aucune commande à exécuter.")
            return
        try:
            loop_count = int(self.entry_loop.get().strip())
        except ValueError:
            messagebox.showerror("Erreur", "Veuillez saisir un nombre de répétitions valide.")
            return
        self.log("Exécution de la macro démarrée.")
        self.macro_running = True
        self.button_stop.config(state=tk.NORMAL)
        thread = threading.Thread(target=self.execute_macro, args=(loop_count,))
        thread.daemon = True
        thread.start()
        
    def execute_macro(self, loop_count):
        iteration = 0
        while self.macro_running and (loop_count == 0 or iteration < loop_count):
            self.log(f"Début de la répétition {iteration+1}.")
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
                        self.log(f"Exécution d'appui de touche: {key}")
                        time.sleep(0.05)
                elif cmd["command"] == "key_hold":
                    key = cmd["key"]
                    duration = cmd["duration"]
                    try:
                        key_obj = getattr(keyboard.Key, key.lower())
                    except AttributeError:
                        key_obj = key
                    self.keyboard_controller.press(key_obj)
                    self.log(f"Début du maintien de la touche: {key}")
                    time.sleep(duration)
                    self.keyboard_controller.release(key_obj)
                    self.log(f"Fin du maintien de la touche: {key}")
                elif cmd["command"] == "wait":
                    duration = cmd["duration"]
                    self.log(f"Début de l'attente: {duration} sec")
                    time.sleep(duration)
                    self.log("Fin de l'attente")
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
                    self.log(f"Exécution du clic de souris: ({x}, {y}), bouton: {button_str}")
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
                    self.log(f"Début du maintien du clic: ({x}, {y}), bouton: {button_str}")
                    time.sleep(cmd["duration"])
                    self.mouse_controller.release(btn)
                    self.log(f"Fin du maintien du clic: ({x}, {y}), bouton: {button_str}")
                elif cmd["command"] == "mouse_scroll":
                    dx = cmd["dx"]
                    dy = cmd["dy"]
                    self.mouse_controller.scroll(dx, dy)
                    self.log(f"Défilement de souris: horizontal {dx}, vertical {dy}")
                time.sleep(0.1)
            iteration += 1
            self.log(f"Répétition {iteration} terminée.")
        self.log("Exécution de la macro terminée.")
        self.macro_running = False
        self.button_stop.config(state=tk.DISABLED)
        
    def stop_macro(self):
        self.macro_running = False
        self.log("Demande d'arrêt de la macro.")
        
    def save_macro(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", 
                                                 filetypes=[("Fichiers JSON", "*.json")])
        if not file_path:
            return
        try:
            with open(file_path, "w") as f:
                json.dump(self.commands, f, indent=4)
            self.log("Macro enregistrée: " + file_path)
        except Exception as e:
            messagebox.showerror("Erreur", "Échec de l'enregistrement de la macro: " + str(e))
        
    def load_macro(self):
        file_path = filedialog.askopenfilename(filetypes=[("Fichiers JSON", "*.json")])
        if not file_path:
            return
        try:
            with open(file_path, "r") as f:
                self.commands = json.load(f)
            self.listbox.delete(0, tk.END)
            for cmd in self.commands:
                disp = self.get_display_text(cmd)
                self.listbox.insert(tk.END, disp)
                self.log("Commande chargée: " + disp)
            self.log("Chargement de la macro terminé: " + file_path)
        except Exception as e:
            messagebox.showerror("Erreur", "Échec du chargement de la macro: " + str(e))
            
    def get_display_text(self, cmd):
        try:
            if cmd.get("command") == "key_tap":
                return f"Appui de touche: {cmd.get('key', '')} x {cmd.get('repeat', 1)} fois"
            elif cmd.get("command") == "key_hold":
                return f"Maintien de touche: {cmd.get('key', '')} (durée: {cmd.get('duration', 0)} sec)"
            elif cmd.get("command") == "wait":
                return f"Attente: {cmd.get('duration', 0)} sec"
            elif cmd.get("command") == "mouse_click":
                return f"Clic de souris: ({cmd.get('x', 0)}, {cmd.get('y', 0)}), bouton: {cmd.get('button', '')}"
            elif cmd.get("command") == "mouse_hold":
                return f"Maintien de clic: ({cmd.get('x', 0)}, {cmd.get('y', 0)}), bouton: {cmd.get('button', '')} (durée: {cmd.get('duration', 0)} sec)"
            elif cmd.get("command") == "mouse_scroll":
                return f"Défilement de souris: horizontal {cmd.get('dx',0)}, vertical {cmd.get('dy',0)}"
            else:
                return str(cmd)
        except Exception as e:
            self.log("Erreur d'affichage de la commande: " + str(e))
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
        self.log("Écouteur de raccourcis démarré avec macro: démarrer={} arrêter={}; enregistrement: démarrer={} arrêter={}".format(
            self.format_hotkey(self.start_hotkey_var.get()),
            self.format_hotkey(self.stop_hotkey_var.get()),
            self.format_hotkey(self.action_start_hotkey_var.get()),
            self.format_hotkey(self.action_stop_hotkey_var.get())
        ))
        
    def stop_hotkey_listener(self):
        if self.hotkey_listener:
            self.hotkey_listener.stop()
            self.hotkey_listener = None
            self.log("Écouteur de raccourcis arrêté.")
        
    def apply_hotkeys(self):
        self.stop_hotkey_listener()
        self.start_hotkey_listener()
        self.log("Raccourcis appliqués.")
        
    def on_hotkey_start(self):
        self.log("Exécution de la macro demandée via raccourci.")
        if not self.macro_running:
            self.play_macro()
        
    def on_hotkey_stop(self):
        self.log("Arrêt de la macro demandé via raccourci.")
        self.stop_macro()
        
    def start_action_recording(self):
        if self.action_recording:
            return
        self.action_recording = True
        self.recorded_commands = []
        self.last_record_time = time.time()
        self.log("Enregistrement des actions démarré.")
        self.action_keyboard_listener = keyboard.Listener(on_release=self.action_on_key_release)
        self.action_mouse_listener = mouse.Listener(on_click=self.action_on_mouse_click)
        self.action_keyboard_listener.start()
        self.action_mouse_listener.start()
        self.button_toggle_recording.config(text="Arrêter enregistrement")
        
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
        self.log("Enregistrement des actions terminé. Ajout des actions enregistrées à la liste des commandes.")
        for cmd in self.recorded_commands:
            self.commands.append(cmd)
            self.listbox.insert(tk.END, self.get_display_text(cmd))
        self.button_toggle_recording.config(text="Démarrer enregistrement")
        
    def action_on_key_release(self, key):
        if not self.action_recording:
            return
        now = time.time()
        dt = now - self.last_record_time
        if dt > RECORD_WAIT_THRESHOLD:
            wait_cmd = {"command": "wait", "duration": round(dt, 2)}
            self.recorded_commands.append(wait_cmd)
            self.log("Attente enregistrée: {} sec".format(round(dt, 2)))
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
        self.log("Appui de touche enregistré: {}".format(k))
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
            self.log("Attente enregistrée: {} sec".format(round(dt, 2)))
        mouse_cmd = {"command": "mouse_click", "x": int(x), "y": int(y), "button": "left"}
        self.recorded_commands.append(mouse_cmd)
        self.log("Clic de souris enregistré: ({}, {})".format(int(x), int(y)))
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
        edit_win.title("Modifier commande")
        edit_win.configure(bg=BG_COLOR)
        edit_win.wait_visibility()
        edit_win.grab_set()
        def save_changes():
            if cmd["command"] == "key_tap":
                new_key = entry_key.get().strip()
                try:
                    new_repeat = int(entry_repeat.get().strip())
                except ValueError:
                    messagebox.showerror("Erreur", "Veuillez saisir un nombre de répétitions valide.", parent=edit_win)
                    return
                if new_key == "":
                    messagebox.showerror("Erreur", "Veuillez saisir une touche.", parent=edit_win)
                    return
                cmd["key"] = new_key
                cmd["repeat"] = new_repeat
            elif cmd["command"] == "key_hold":
                new_key = entry_key.get().strip()
                try:
                    new_duration = float(entry_duration.get().strip())
                except ValueError:
                    messagebox.showerror("Erreur", "Veuillez saisir une durée valide.", parent=edit_win)
                    return
                if new_key == "":
                    messagebox.showerror("Erreur", "Veuillez saisir une touche.", parent=edit_win)
                    return
                cmd["key"] = new_key
                cmd["duration"] = new_duration
            elif cmd["command"] == "wait":
                try:
                    new_duration = float(entry_duration.get().strip())
                except ValueError:
                    messagebox.showerror("Erreur", "Veuillez saisir un temps d'attente valide.", parent=edit_win)
                    return
                cmd["duration"] = new_duration
            elif cmd["command"] == "mouse_click":
                try:
                    new_x = int(entry_x.get().strip())
                    new_y = int(entry_y.get().strip())
                except ValueError:
                    messagebox.showerror("Erreur", "Veuillez saisir des valeurs entières valides pour X et Y.", parent=edit_win)
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
                    messagebox.showerror("Erreur", "Veuillez saisir des valeurs entières valides pour X et Y.", parent=edit_win)
                    return
                try:
                    new_duration = float(entry_duration.get().strip())
                except ValueError:
                    messagebox.showerror("Erreur", "Veuillez saisir une durée valide.", parent=edit_win)
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
                    messagebox.showerror("Erreur", "Veuillez saisir des valeurs de défilement valides.", parent=edit_win)
                    return
                cmd["dx"] = new_dx
                cmd["dy"] = new_dy
            self.listbox.delete(index)
            self.listbox.insert(index, self.get_display_text(cmd))
            self.log("Commande modifiée: " + self.get_display_text(cmd))
            edit_win.destroy()
        if cmd["command"] == "key_tap":
            tk.Label(edit_win, text="Touche:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=5)
            entry_key = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_key.insert(0, cmd["key"])
            entry_key.grid(row=0, column=1, padx=5, pady=5)
            def on_key_press(event):
                entry_key.delete(0, tk.END)
                entry_key.insert(0, event.keysym)
                return "break"
            entry_key.bind("<Key>", on_key_press)
            tk.Label(edit_win, text="Nombre de répétitions:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=5)
            entry_repeat = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_repeat.insert(0, str(cmd.get("repeat", 1)))
            entry_repeat.grid(row=1, column=1, padx=5, pady=5)
        elif cmd["command"] == "key_hold":
            tk.Label(edit_win, text="Touche:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=5)
            entry_key = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_key.insert(0, cmd["key"])
            entry_key.grid(row=0, column=1, padx=5, pady=5)
            def on_key_press(event):
                entry_key.delete(0, tk.END)
                entry_key.insert(0, event.keysym)
                return "break"
            entry_key.bind("<Key>", on_key_press)
            tk.Label(edit_win, text="Durée (sec):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=5)
            entry_duration = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.insert(0, str(cmd.get("duration", 1)))
            entry_duration.grid(row=1, column=1, padx=5, pady=5)
        elif cmd["command"] == "wait":
            tk.Label(edit_win, text="Temps d'attente (sec):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
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
            tk.Label(edit_win, text="Bouton:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=4, padx=5, pady=5)
            var_button = tk.StringVar(value=cmd["button"])
            option_button = tk.OptionMenu(edit_win, var_button, "left", "right", "middle")
            option_button.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
            option_button["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
            option_button.grid(row=0, column=5, padx=5, pady=5)
            def record_mouse_edit():
                record_button_edit.config(text="Clique gauche pour enregistrer la position", state=tk.DISABLED)
                def on_click(x, y, button, pressed):
                    if button == mouse.Button.left and pressed:
                        entry_x.delete(0, tk.END)
                        entry_x.insert(0, str(int(x)))
                        entry_y.delete(0, tk.END)
                        entry_y.insert(0, str(int(y)))
                        edit_win.after(0, lambda: self.log(f"Position de la souris enregistrée dans la modification: ({int(x)}, {int(y)})"))
                        record_button_edit.config(text="Enregistrer position de la souris", state=tk.NORMAL)
                        return False
                listener = mouse.Listener(on_click=on_click)
                listener.start()
            record_button_edit = tk.Button(edit_win, text="Enregistrer position de la souris", command=record_mouse_edit,
                                           bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                           activebackground=BUTTON_ACTIVE_BG)
            record_button_edit.grid(row=1, column=0, columnspan=6, padx=5, pady=5, sticky="w")
        elif cmd["command"] == "mouse_hold":
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
            tk.Label(edit_win, text="Bouton:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=4, padx=5, pady=5)
            var_button = tk.StringVar(value=cmd["button"])
            option_button = tk.OptionMenu(edit_win, var_button, "left", "right", "middle")
            option_button.config(bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT)
            option_button["menu"].config(bg=ENTRY_BG, fg=ENTRY_FG, font=FONT)
            option_button.grid(row=0, column=5, padx=5, pady=5)
            tk.Label(edit_win, text="Durée (sec):", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=1, column=0, padx=5, pady=5)
            entry_duration = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_duration.insert(0, str(cmd.get("duration", 1)))
            entry_duration.grid(row=1, column=1, padx=5, pady=5)
            def record_mouse_edit():
                record_button_edit.config(text="Clique gauche pour enregistrer la position", state=tk.DISABLED)
                def on_click(x, y, button, pressed):
                    if button == mouse.Button.left and pressed:
                        entry_x.delete(0, tk.END)
                        entry_x.insert(0, str(int(x)))
                        entry_y.delete(0, tk.END)
                        entry_y.insert(0, str(int(y)))
                        edit_win.after(0, lambda: self.log(f"Position de la souris enregistrée dans la modification: ({int(x)}, {int(y)})"))
                        record_button_edit.config(text="Enregistrer position de la souris", state=tk.NORMAL)
                        return False
                listener = mouse.Listener(on_click=on_click)
                listener.start()
            record_button_edit = tk.Button(edit_win, text="Enregistrer position de la souris", command=record_mouse_edit,
                                           bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT,
                                           activebackground=BUTTON_ACTIVE_BG)
            record_button_edit.grid(row=1, column=2, columnspan=4, padx=5, pady=5, sticky="w")
        elif cmd["command"] == "mouse_scroll":
            tk.Label(edit_win, text="Défilement horizontal:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=0, padx=5, pady=5)
            entry_dx = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_dx.insert(0, str(cmd["dx"]))
            entry_dx.grid(row=0, column=1, padx=5, pady=5)
            tk.Label(edit_win, text="Défilement vertical:", bg=LABEL_BG, fg=LABEL_FG, font=FONT)\
                .grid(row=0, column=2, padx=5, pady=5)
            entry_dy = tk.Entry(edit_win, width=10, bg=ENTRY_BG, fg=ENTRY_FG, font=FONT, relief=tk.FLAT)
            entry_dy.insert(0, str(cmd["dy"]))
            entry_dy.grid(row=0, column=3, padx=5, pady=5)
        tk.Button(edit_win, text="Enregistrer", command=save_changes,
                  bg=BUTTON_BG, fg=BUTTON_FG, font=FONT, relief=tk.FLAT, activebackground=BUTTON_ACTIVE_BG)\
            .grid(row=10, column=0, padx=5, pady=10)
        tk.Button(edit_win, text="Annuler", command=edit_win.destroy,
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
                self.log(f"Commande déplacée de {self.drag_original_index} à {self.drop_index}.")
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

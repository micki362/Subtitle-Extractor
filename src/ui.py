import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import subprocess
from config import LIGHT_THEME, DARK_THEME

class SubtitleExtractorUI:
    def __init__(self, master, app_logic):
        self.master = master
        self.logic = app_logic
        self.settings = self.logic.config.settings
        self.current_theme = self.logic.current_theme

        master.title("Bulk Subtitle Extractor")
        master.geometry("800x700")

        self.style = ttk.Style()
        self._create_widgets()
        self._load_and_apply_initial_state()

    def _create_widgets(self):
        self.main_frame_container = ttk.Frame(self.master, padding="10")
        self.main_frame_container.pack(fill=tk.BOTH, expand=True)

        self._create_folder_widgets()
        self._create_file_list_widgets()
        self._create_options_widgets()
        self._create_action_widgets()
        self._create_bottom_bar_widgets()

    def _create_folder_widgets(self):
        self.folder_frame = ttk.Frame(self.main_frame_container)
        self.folder_frame.pack(fill=tk.X, pady=5)
        self.folder_label = ttk.Label(self.folder_frame, text="No folder selected.", width=60, anchor="w")
        self.folder_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.select_folder_button = ttk.Button(self.folder_frame, text="Select Folder", command=self.logic.select_folder)
        self.select_folder_button.pack(side=tk.RIGHT)

    def _create_file_list_widgets(self):
        self.list_frame = ttk.Frame(self.main_frame_container)
        self.list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.scrollbar_y = ttk.Scrollbar(self.list_frame, orient=tk.VERTICAL)
        self.file_tree = ttk.Treeview(self.list_frame, columns=("File Name", "Status"), show="headings", yscrollcommand=self.scrollbar_y.set, selectmode="extended")
        self.file_tree.heading("File Name", text="File Name")
        self.file_tree.heading("Status", text="Status")
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar_y.config(command=self.file_tree.yview)
        self.scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

    def _create_options_widgets(self):
        self.options_ui_frame = ttk.Frame(self.main_frame_container)
        self.options_ui_frame.pack(fill=tk.X, pady=5)

        format_frame = ttk.Frame(self.options_ui_frame)
        format_frame.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        ttk.Label(format_frame, text="Output Format:").pack(side=tk.LEFT, padx=(5, 2))
        self.output_format_var = tk.StringVar(value=self.settings['default_output_format'])
        self.format_options = ["srt", "ass", "vtt", "copy"]
        self.format_combobox = ttk.Combobox(format_frame, textvariable=self.output_format_var,
                                            values=self.format_options, state="readonly", width=10)
        self.format_combobox.pack(side=tk.LEFT, padx=2)
        self.format_combobox.bind("<<ComboboxSelected>>", self.logic.on_format_selected)

        self.select_langs_button = ttk.Button(format_frame, text="Filter Languages...", command=self.logic.open_language_filter_dialog)
        self.select_langs_button.pack(side=tk.LEFT, padx=(10, 5))
        self.ocr_settings_button = ttk.Button(format_frame, text="OCR Settings...", command=self.open_ocr_settings_dialog)
        self.ocr_settings_button.pack(side=tk.LEFT, padx=5)

        self.current_lang_filter_label = ttk.Label(format_frame, text=self.logic._get_current_lang_filter_display())
        self.current_lang_filter_label.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)

        self.action_buttons_frame = ttk.Frame(self.options_ui_frame)
        self.action_buttons_frame.pack(side=tk.RIGHT, fill=tk.Y)
        self.skip_if_exists_var = tk.BooleanVar(value=self.settings['skip_if_exists'])
        self.skip_checkbox = ttk.Checkbutton(self.action_buttons_frame, text="Skip if exists",
                                             variable=self.skip_if_exists_var, command=self.logic.on_skip_toggle, style="TCheckbutton")
        self.skip_checkbox.pack(side=tk.LEFT, padx=(0, 10))
        self.remove_button = ttk.Button(self.action_buttons_frame, text="Remove Selected", command=self.logic.remove_selected_files)
        self.remove_button.pack(side=tk.RIGHT, padx=5)
        self.remove_with_subs_button = ttk.Button(self.action_buttons_frame, text="Remove with Subtitles", command=self.logic.remove_files_with_subtitles)
        self.remove_with_subs_button.pack(side=tk.RIGHT, padx=5)

    def _create_action_widgets(self):
        self.extract_button_frame = ttk.Frame(self.main_frame_container)
        self.extract_button_frame.pack(fill=tk.X, pady=5)
        self.extract_button = ttk.Button(self.extract_button_frame, text="Extract Subtitles", command=self.logic.start_extraction_thread, style="Accent.TButton")
        self.extract_button.pack(pady=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.main_frame_container, orient="horizontal", length=300, mode="determinate", variable=self.progress_var, style="Custom.Horizontal.TProgressbar")
        self.progress_bar.pack(fill=tk.X, pady=(5, 0), padx=5)

        self.status_label = ttk.Label(self.main_frame_container, text="Ready, Commander.", anchor="w")
        self.status_label.pack(fill=tk.X, pady=(5, 10), padx=5)

    def _create_bottom_bar_widgets(self):
        self.bottom_buttons_frame = ttk.Frame(self.main_frame_container)
        self.bottom_buttons_frame.pack(fill=tk.X, pady=5)
        self.view_log_button = ttk.Button(self.bottom_buttons_frame, text="View Log", command=self.open_log_window)
        self.view_log_button.pack(side=tk.RIGHT, padx=5)
        self.theme_button = ttk.Button(self.bottom_buttons_frame, text="Toggle Dark/Light Side", command=self.logic.toggle_theme)
        self.theme_button.pack(side=tk.RIGHT, padx=5)
        self.edit_config_button = ttk.Button(self.bottom_buttons_frame, text="Edit Holocron (Config)", command=self.open_config_file)
        self.edit_config_button.pack(side=tk.LEFT, padx=5)

    def _load_and_apply_initial_state(self):
        if self.settings['last_folder'] and os.path.isdir(self.settings['last_folder']):
            self.folder_label.config(text=self.settings['last_folder'])

        self.apply_theme()
        self.master.protocol("WM_DELETE_WINDOW", self.logic._on_closing_main)

    def apply_theme(self):
        theme = self.logic.current_theme; self.master.configure(bg=theme["bg"]); self.style.theme_use('clam')
        self.style.configure("TFrame", background=theme["bg"])
        self.style.configure("TLabel", background=theme["bg"], foreground=theme["fg"], padding=2)
        self.style.configure("TButton", background=theme["button_bg"], foreground=theme["button_fg"], padding=5, relief="flat", borderwidth=1)
        self.style.map("TButton", background=[('active', theme["accent_bg"]), ('pressed', theme["accent_bg"])], foreground=[('active', theme["accent_fg"]), ('pressed', theme["accent_fg"])])
        self.style.configure("Accent.TButton", background=theme["accent_bg"], foreground=theme["accent_fg"], font=('Helvetica', 10, 'bold'), padding=6, relief="flat", borderwidth=1)
        self.style.map("Accent.TButton", background=[('active', theme["button_bg"]), ('pressed', theme["button_bg"])], foreground=[('active', theme["button_fg"]), ('pressed', theme["button_fg"])])
        self.style.configure("Custom.Horizontal.TProgressbar", troughcolor=theme["progress_trough"], background=theme["progress_bg"], bordercolor=theme["bg"], lightcolor=theme["bg"], darkcolor=theme["bg"])
        self.style.configure("TCombobox", selectbackground=theme["list_bg"], fieldbackground=theme["list_bg"], background=theme["button_bg"], foreground=theme["fg"])
        self.style.map('TCombobox', fieldbackground=[('readonly', theme["list_bg"])], selectbackground=[('readonly', theme["accent_bg"])], selectforeground=[('readonly', theme["accent_fg"])])
        self.style.map("TCheckbutton", background=[('active', theme["bg"])], indicatorcolor=[('selected', theme["accent_bg"]), ('!selected', theme["button_bg"])], foreground=[('disabled', theme.get("disabled_fg", "#A0A0A0"))])
        self.style.configure("Treeview", background=theme["list_bg"], foreground=theme["list_fg"], fieldbackground=theme["list_bg"])
        self.style.map("Treeview", background=[('selected', theme["accent_bg"])], foreground=[('selected', theme["accent_fg"])])
        self.folder_label.configure(background=theme["bg"], foreground=theme["fg"]); self.status_label.configure(background=theme["bg"], foreground=theme["fg"])
        self.theme_button.config(text="Join the Light Side" if self.logic.is_dark_mode else "Embrace the Dark Side")
        if self.logic.log_window and self.logic.log_window.winfo_exists():
            self.logic.log_window.configure(bg=theme["bg"])
            if self.logic.log_text_widget: self.logic.log_text_widget.configure(bg=theme["log_bg"], fg=theme["log_fg"], insertbackground=theme["fg"])
            for child_frame in self.logic.log_window.winfo_children():
                if isinstance(child_frame, ttk.Frame):
                    for btn in child_frame.winfo_children():
                        if isinstance(btn, ttk.Button): btn.configure(style="TButton")
        self.select_folder_button.configure(style="TButton"); self.remove_button.configure(style="TButton"); self.view_log_button.configure(style="TButton")
        self.theme_button.configure(style="TButton"); self.edit_config_button.configure(style="TButton"); self.select_langs_button.configure(style="TButton")
        self.ocr_settings_button.configure(style="TButton")
        self.skip_checkbox.configure(style="TCheckbutton")
        self.extract_button.configure(style="Accent.TButton")

    def open_config_file(self):
        config_path = self.logic.config.get_config_path()
        try:
            if os.name == 'nt': os.startfile(config_path)
            elif 'darwin' in sys.platform: subprocess.call(('open', config_path))
            elif 'linux' in sys.platform: subprocess.call(('xdg-open', config_path))
            else: messagebox.showinfo("Info", f"Please open this Holocron manually:\n{config_path}", parent=self.master)
            self.logic.log_message(f"Opening Holocron (config file): {config_path}. Restart required for changes to take effect.", to_console=False)
        except Exception as e:
            self.logic.log_message(f"Could not open Holocron: {e}", to_console=True)
            messagebox.showerror("Error", f"Could not open Holocron editor.\nPath: {config_path}", parent=self.master)

    def open_ocr_settings_dialog(self):
        dialog = tk.Toplevel(self.master)
        dialog.title("OCR Holocron Settings")
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.configure(bg=self.current_theme["bg"])
        dialog.minsize(500, 200)

        content_frame = ttk.Frame(dialog, padding=15)
        content_frame.pack(expand=True, fill=tk.BOTH)

        ocr_enabled_var = tk.BooleanVar(value=self.settings.get('ocr_enabled', False))
        ttk.Checkbutton(content_frame, text="Enable OCR Droid (for image-based subtitles)",
                        variable=ocr_enabled_var, style="TCheckbutton").pack(anchor='w', pady=(0, 10))

        ttk.Label(content_frame, text="OCR Droid Protocol (Command Template):").pack(anchor='w')
        cmd_frame = ttk.Frame(content_frame)
        cmd_frame.pack(fill=tk.X, expand=True, pady=(0, 10))
        ocr_cmd_var = tk.StringVar(value=self.settings.get('ocr_command_template', ''))
        cmd_entry = ttk.Entry(cmd_frame, textvariable=ocr_cmd_var)
        cmd_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        def browse_for_ocr_exe():
            exe_path = filedialog.askopenfilename(title="Select OCR Executable",
                                                  filetypes=[("Executable files", "*.exe"), ("All files", "*.*")])
            if exe_path:
                current_cmd = ocr_cmd_var.get()
                parts = current_cmd.split(" ")
                if len(parts) > 1 and os.path.isfile(parts[0]):
                    new_cmd = f'"{exe_path}" {" ".join(parts[1:])}'
                else:
                    new_cmd = f'"{exe_path}" {current_cmd}'
                ocr_cmd_var.set(new_cmd.strip())

        browse_btn = ttk.Button(cmd_frame, text="Browse...", command=browse_for_ocr_exe)
        browse_btn.pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Label(content_frame, text="Use placeholders: {INPUT_FILE_PATH}, {OUTPUT_SRT_PATH}, {LANG_3_CODE}",
                  font=("Helvetica", 8)).pack(anchor='w', pady=(0, 10))

        ttk.Label(content_frame, text="Default Language for OCR (3-letter code):").pack(anchor='w')
        ocr_lang_var = tk.StringVar(value=self.settings.get('ocr_default_lang', 'eng'))
        ttk.Entry(content_frame, textvariable=ocr_lang_var, width=10).pack(anchor='w', pady=(0, 15))

        button_frame = ttk.Frame(content_frame)
        button_frame.pack(fill='x', side=tk.BOTTOM)

        def on_save():
            self.settings['ocr_enabled'] = ocr_enabled_var.get()
            self.settings['ocr_command_template'] = ocr_cmd_var.get()
            self.settings['ocr_default_lang'] = ocr_lang_var.get()
            self.logic.config.save_config(self.logic.extract_all_languages_flag, self.logic.user_selected_languages)
            self.logic.log_message("OCR Holocron settings updated and saved.", to_console=False)
            dialog.destroy()

        ok_button = ttk.Button(button_frame, text="Save Protocol", command=on_save, style="Accent.TButton")
        ok_button.pack(side=tk.RIGHT, padx=5)
        cancel_button = ttk.Button(button_frame, text="Discard Changes", command=dialog.destroy)
        cancel_button.pack(side=tk.RIGHT)
        dialog.wait_window()

    def open_log_window(self):
        if self.logic.log_window and self.logic.log_window.winfo_exists(): self.logic.log_window.lift(); self.logic.log_window.focus_set(); return
        self.logic.log_window = tk.Toplevel(self.master); self.logic.log_window.title("Mission Debrief (Log)"); self.logic.log_window.geometry("700x500"); self.logic.log_window.configure(bg=self.current_theme["bg"])
        log_button_frame = ttk.Frame(self.logic.log_window); log_button_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        copy_button = ttk.Button(log_button_frame, text="Copy to Datapad", command=self.logic.copy_log_to_clipboard); copy_button.pack(side=tk.LEFT, padx=5)
        save_button = ttk.Button(log_button_frame, text="Save Mission Log", command=self.logic.save_log_to_file)
        if not self.logic.log_dir_path: save_button.config(state=tk.DISABLED)
        save_button.pack(side=tk.LEFT, padx=5)
        self.logic.log_text_widget = scrolledtext.ScrolledText(self.logic.log_window, wrap=tk.WORD, state=tk.DISABLED, padx=5, pady=5); self.logic.log_text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        if hasattr(self.logic.log_text_widget, 'vbar'): self.logic.log_text_widget.vbar.config(width=12)
        self.logic.log_text_widget.config(state=tk.NORMAL)
        for msg in self.logic.log_buffer: self.logic.log_text_widget.insert(tk.END, msg)
        self.logic.log_text_widget.see(tk.END); self.logic.log_text_widget.config(state=tk.DISABLED)
        self.logic.log_text_widget.configure(bg=self.current_theme["log_bg"], fg=self.current_theme["log_fg"], insertbackground=self.current_theme["fg"])
        for child in log_button_frame.winfo_children():
            if isinstance(child, ttk.Button): child.configure(style="TButton")
        self.logic.log_window.protocol("WM_DELETE_WINDOW", self.logic._on_closing_log_window)
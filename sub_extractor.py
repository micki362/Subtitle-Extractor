import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import subprocess
import threading
import re
import shutil
import datetime
import configparser
import tempfile # For OCR temporary files
import random # For witty messages

# --- Configuration ---
MOVIE_EXTENSIONS = ('.mkv', '.mp4', '.avi', '.mov', '.wmv', '.flv')
SUBTITLE_EXTENSIONS = ('.srt', '.ass', '.vtt', '.sub', '.sup') # For checking existing subs
DEFAULT_FFMPEG_PATH = "ffmpeg"
DEFAULT_FFPROBE_PATH = "ffprobe"
DEFAULT_FFPROBE_TIMEOUT = 60
DEFAULT_FFMPEG_EXTRACT_TIMEOUT = 600
DEFAULT_FFMPEG_OCR_TIMEOUT = 1800 # 30 minutes for OCR
LOG_FOLDER_NAME = "logs"
CONFIG_FILENAME = "sub_extractor_settings.ini"

# --- Theme Colors ---
LIGHT_THEME = {
    "bg": "#ECECEC", "fg": "#000000", "list_bg": "#FFFFFF", "list_fg": "#000000",
    "button_bg": "#DDDDDD", "button_fg": "#000000", "accent_bg": "#777777",
    "accent_fg": "#FFFFFF", "log_bg": "#F0F0F0", "log_fg": "#000000",
    "progress_bg": "#A0A0A0", "progress_trough": "#D0D0D0"
}
DARK_THEME = {
    "bg": "#2D2D2D", "fg": "#FFFFFF", "list_bg": "#3C3C3C", "list_fg": "#FFFFFF",
    "button_bg": "#555555", "button_fg": "#FFFFFF", "accent_bg": "#0078D7",
    "accent_fg": "#FFFFFF", "log_bg": "#252525", "log_fg": "#FFFFFF",
    "progress_bg": "#0078D7", "progress_trough": "#4A4A4A"
}

# --- Constants for Subtitle Processing ---
IMAGE_BASED_CODECS = {'hdmv_pgs_subtitle', 'pgssub', 'dvd_subtitle', 'dvdsub'}
TEXT_BASED_OUTPUT_FORMATS = {'srt', 'ass', 'vtt'}

# --- Witty OCR Messages ---
OCR_PATIENCE_MESSAGES = [
    "Engaging OCR hyperdrive for {filename}... This is where the fun begins!",
    "Patience you must have, my young Padawan. OCRing {filename}...",
    "The OCR Force is strong with this one... {filename}. Just a moment.",
    "Calculating OCR trajectory for {filename}... It's not a trap!",
    "R2-D2 is decoding image subs for {filename}... Beep boop bleep!",
    "These aren't the droids you're looking for... but these OCR results for {filename} might be!",
    "Hold your blasters! OCRing {filename} can take time. The Emperor is not as forgiving as I am.",
    "Never tell me the odds! OCRing {filename} is in progress..."
]


class SubtitleExtractorApp:
    def __init__(self, master):
        self.master = master
        master.title("Bulk Subtitle Extractor")
        master.geometry("800x700") # Increased size for new button

        self.app_dir = os.path.dirname(os.path.abspath(__file__))

        self._setup_config_and_theme()

        self.extract_all_languages_flag = True
        self.user_selected_languages = set()
        self._parse_loaded_languages()

        self.movie_files_paths = []
        self.files_with_success, self.files_with_no_subs, self.files_timed_out, self.files_with_errors, self.files_skipped = [], [], [], [], []
        self.log_buffer, self.log_window, self.log_text_widget = [], None, None
        self.style = ttk.Style()
        self.cancel_requested = threading.Event()

        self._setup_logging()

        if not self.check_ffmpeg():
            messagebox.showerror("Error", f"FFmpeg/FFprobe not found (see '{CONFIG_FILENAME}').\nCheck paths and restart.")
            master.destroy()
            return

        self._create_widgets()
        self._load_and_apply_initial_state()

    def _setup_config_and_theme(self):
        """Initializes configuration and theme settings."""
        self.settings = {
            'theme': 'light', 'last_folder': '',
            'ffmpeg_path': DEFAULT_FFMPEG_PATH, 'ffprobe_path': DEFAULT_FFPROBE_PATH,
            'ffprobe_timeout': DEFAULT_FFPROBE_TIMEOUT,
            'ffmpeg_extract_timeout': DEFAULT_FFMPEG_EXTRACT_TIMEOUT,
            'ffmpeg_ocr_timeout': DEFAULT_FFMPEG_OCR_TIMEOUT,
            'default_output_format': 'srt', 'selected_languages': 'all',
            'skip_if_exists': False,
            'ocr_enabled': False, 'ocr_command_template': '', 'ocr_temp_dir': '',
            'ocr_default_lang': 'eng',
            'ocr_input_ext_map': {
                'hdmv_pgs_subtitle': '.sup', 'dvd_subtitle': '.sub'
            }
        }
        self.config = configparser.ConfigParser()
        self.load_config()

        self.current_theme_name = self.settings['theme']
        self.is_dark_mode = (self.current_theme_name == 'dark')
        self.current_theme = DARK_THEME if self.is_dark_mode else LIGHT_THEME

    def _setup_logging(self):
        """Creates the log directory if it doesn't exist."""
        self.log_dir_path = os.path.join(self.app_dir, LOG_FOLDER_NAME)
        if not os.path.exists(self.log_dir_path):
            try:
                os.makedirs(self.log_dir_path)
            except OSError as e:
                print(f"Error creating log dir: {e}")
                self.log_dir_path = None

    def _create_widgets(self):
        """Creates and packs all the GUI widgets."""
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
        self.select_folder_button = ttk.Button(self.folder_frame, text="Select Folder", command=self.select_folder)
        self.select_folder_button.pack(side=tk.RIGHT)

    def _create_file_list_widgets(self):
        self.list_frame = ttk.Frame(self.main_frame_container)
        self.list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.scrollbar_y = ttk.Scrollbar(self.list_frame, orient=tk.VERTICAL)
        self.file_listbox = tk.Listbox(self.list_frame, yscrollcommand=self.scrollbar_y.set, exportselection=False, selectmode=tk.EXTENDED)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar_y.config(command=self.file_listbox.yview)
        self.scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

    def _create_options_widgets(self):
        self.options_ui_frame = ttk.Frame(self.main_frame_container)
        self.options_ui_frame.pack(fill=tk.X, pady=5)

        # Left side: Format and Language filters
        format_frame = ttk.Frame(self.options_ui_frame)
        format_frame.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        ttk.Label(format_frame, text="Output Format:").pack(side=tk.LEFT, padx=(5, 2))
        self.output_format_var = tk.StringVar(value=self.settings['default_output_format'])
        self.format_options = ["srt", "ass", "vtt", "copy"]
        self.format_combobox = ttk.Combobox(format_frame, textvariable=self.output_format_var,
                                            values=self.format_options, state="readonly", width=10)
        self.format_combobox.pack(side=tk.LEFT, padx=2)
        self.format_combobox.bind("<<ComboboxSelected>>", self.on_format_selected)

        self.select_langs_button = ttk.Button(format_frame, text="Filter Languages...", command=self.open_language_filter_dialog)
        self.select_langs_button.pack(side=tk.LEFT, padx=(10, 5))
        self.ocr_settings_button = ttk.Button(format_frame, text="OCR Settings...", command=self.open_ocr_settings_dialog)
        self.ocr_settings_button.pack(side=tk.LEFT, padx=5)

        self.current_lang_filter_label = ttk.Label(format_frame, text=self._get_current_lang_filter_display())
        self.current_lang_filter_label.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)

        # Right side: Skip and Remove buttons
        self.action_buttons_frame = ttk.Frame(self.options_ui_frame)
        self.action_buttons_frame.pack(side=tk.RIGHT, fill=tk.Y)
        self.skip_if_exists_var = tk.BooleanVar(value=self.settings['skip_if_exists'])
        self.skip_checkbox = ttk.Checkbutton(self.action_buttons_frame, text="Skip if exists",
                                             variable=self.skip_if_exists_var, command=self.on_skip_toggle, style="TCheckbutton")
        self.skip_checkbox.pack(side=tk.LEFT, padx=(0, 10))
        self.remove_button = ttk.Button(self.action_buttons_frame, text="Remove Selected", command=self.remove_selected_files)
        self.remove_button.pack(side=tk.RIGHT, padx=5)

    def _create_action_widgets(self):
        self.extract_button_frame = ttk.Frame(self.main_frame_container)
        self.extract_button_frame.pack(fill=tk.X, pady=5)
        self.extract_button = ttk.Button(self.extract_button_frame, text="Extract Subtitles", command=self.start_extraction_thread, style="Accent.TButton")
        self.extract_button.pack(pady=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.main_frame_container, orient="horizontal", length=300, mode="determinate", variable=self.progress_var, style="Custom.Horizontal.TProgressbar")
        self.progress_bar.pack(fill=tk.X, pady=(5, 0), padx=5)

        self.status_label = ttk.Label(self.main_frame_container, text="Ready, Commander.", anchor="w") # Star Wars touch
        self.status_label.pack(fill=tk.X, pady=(5, 10), padx=5)

    def _create_bottom_bar_widgets(self):
        self.bottom_buttons_frame = ttk.Frame(self.main_frame_container)
        self.bottom_buttons_frame.pack(fill=tk.X, pady=5)
        self.view_log_button = ttk.Button(self.bottom_buttons_frame, text="View Log", command=self.open_log_window)
        self.view_log_button.pack(side=tk.RIGHT, padx=5)
        self.theme_button = ttk.Button(self.bottom_buttons_frame, text="Toggle Dark/Light Side", command=self.toggle_theme) # Star Wars touch
        self.theme_button.pack(side=tk.RIGHT, padx=5)
        self.edit_config_button = ttk.Button(self.bottom_buttons_frame, text="Edit Holocron (Config)", command=self.open_config_file) # Star Wars touch
        self.edit_config_button.pack(side=tk.LEFT, padx=5)

    def _load_and_apply_initial_state(self):
        """Applies loaded settings to the UI on startup."""
        if self.settings['last_folder'] and os.path.isdir(self.settings['last_folder']):
            self.folder_label.config(text=self.settings['last_folder'])

        self.apply_theme()
        self.master.protocol("WM_DELETE_WINDOW", self._on_closing_main)

    def on_skip_toggle(self):
        self.settings['skip_if_exists'] = self.skip_if_exists_var.get()
        self.log_message(f"Skip-if-exists set to: {self.settings['skip_if_exists']}", to_console=False)

    def _parse_loaded_languages(self):
        loaded_lang_str = self.settings.get('selected_languages', 'all').strip().lower()
        if not loaded_lang_str or loaded_lang_str == 'all':
            self.extract_all_languages_flag = True
            self.user_selected_languages = set()
        else:
            self.extract_all_languages_flag = False
            self.user_selected_languages = {lang.strip() for lang in loaded_lang_str.split(',') if lang.strip()}

    def on_format_selected(self, event=None):
        selected_format = self.output_format_var.get()
        self.settings['default_output_format'] = selected_format
        self.log_message(f"Output format set to: {selected_format}", to_console=False)

    def _get_current_lang_filter_display(self):
        if self.extract_all_languages_flag or not self.user_selected_languages:
            return "All Languages (Galactic Basic)"
        display_langs = sorted(list(self.user_selected_languages))
        if len(display_langs) > 3:
            return f"{', '.join(display_langs[:3])}, ..."
        return ', '.join(display_langs)

    def get_config_path(self):
        return os.path.join(self.app_dir, CONFIG_FILENAME)

    def load_config(self):
        config_path = self.get_config_path()
        self.config.read(config_path, encoding='utf-8')
        def get_cfg(section, option, fallback, type_func=None):
            if not self.config.has_section(section): self.config.add_section(section)
            try:
                if type_func == int: return self.config.getint(section, option)
                elif type_func == bool: return self.config.getboolean(section, option)
                else: return self.config.get(section, option)
            except (configparser.NoOptionError, ValueError):
                return fallback
        self.settings['theme'] = get_cfg('General', 'theme', self.settings['theme'])
        self.settings['last_folder'] = get_cfg('General', 'last_folder', self.settings['last_folder'])
        self.settings['ffmpeg_path'] = get_cfg('Paths', 'ffmpeg_path', self.settings['ffmpeg_path'])
        self.settings['ffprobe_path'] = get_cfg('Paths', 'ffprobe_path', self.settings['ffprobe_path'])
        self.settings['ffprobe_timeout'] = get_cfg('Timeouts', 'ffprobe_timeout', self.settings['ffprobe_timeout'], type_func=int)
        self.settings['ffmpeg_extract_timeout'] = get_cfg('Timeouts', 'ffmpeg_extract_timeout', self.settings['ffmpeg_extract_timeout'], type_func=int)
        self.settings['ffmpeg_ocr_timeout'] = get_cfg('Timeouts', 'ffmpeg_ocr_timeout', self.settings['ffmpeg_ocr_timeout'], type_func=int)
        self.settings['default_output_format'] = get_cfg('Extraction', 'default_output_format', self.settings['default_output_format'])
        self.settings['selected_languages'] = get_cfg('Extraction', 'selected_languages', self.settings['selected_languages'])
        self.settings['skip_if_exists'] = get_cfg('Extraction', 'skip_if_exists', self.settings['skip_if_exists'], type_func=bool)
        self.settings['ocr_enabled'] = get_cfg('OCR', 'ocr_enabled', self.settings['ocr_enabled'], type_func=bool)
        self.settings['ocr_command_template'] = get_cfg('OCR', 'ocr_command_template', self.settings['ocr_command_template'])
        self.settings['ocr_temp_dir'] = get_cfg('OCR', 'ocr_temp_dir', self.settings['ocr_temp_dir'])
        self.settings['ocr_default_lang'] = get_cfg('OCR', 'ocr_default_lang', self.settings['ocr_default_lang'])
        self.settings['ocr_input_ext_map'] = {
            'hdmv_pgs_subtitle': get_cfg('OCR', 'ocr_input_ext_map_hdmv_pgs_subtitle', '.sup'),
            'dvd_subtitle': get_cfg('OCR', 'ocr_input_ext_map_dvd_subtitle', '.sub')
        }
        for sec in ['General', 'Paths', 'Timeouts', 'Extraction', 'OCR']:
            if not self.config.has_section(sec): self.config.add_section(sec)

    def save_config(self):
        config_path = self.get_config_path()
        self.config.set('General', 'theme', self.settings['theme'])
        self.config.set('General', 'last_folder', self.settings['last_folder'])
        self.config.set('Paths', 'ffmpeg_path', self.settings['ffmpeg_path'])
        self.config.set('Paths', 'ffprobe_path', self.settings['ffprobe_path'])
        self.config.set('Timeouts', 'ffprobe_timeout', str(self.settings['ffprobe_timeout']))
        self.config.set('Timeouts', 'ffmpeg_extract_timeout', str(self.settings['ffmpeg_extract_timeout']))
        self.config.set('Timeouts', 'ffmpeg_ocr_timeout', str(self.settings.get('ffmpeg_ocr_timeout', DEFAULT_FFMPEG_OCR_TIMEOUT)))
        self.config.set('Extraction', 'default_output_format', self.settings['default_output_format'])
        lang_str_to_save = 'all' if self.extract_all_languages_flag or not self.user_selected_languages else ','.join(sorted(list(self.user_selected_languages)))
        self.config.set('Extraction', 'selected_languages', lang_str_to_save)
        self.config.set('Extraction', 'skip_if_exists', str(self.settings.get('skip_if_exists', False)))
        self.config.set('OCR', 'ocr_enabled', str(self.settings.get('ocr_enabled', False)))
        self.config.set('OCR', 'ocr_command_template', self.settings.get('ocr_command_template', ''))
        self.config.set('OCR', 'ocr_temp_dir', self.settings.get('ocr_temp_dir', ''))
        self.config.set('OCR', 'ocr_default_lang', self.settings.get('ocr_default_lang', 'eng'))
        ocr_ext_map = self.settings.get('ocr_input_ext_map', {})
        self.config.set('OCR', 'ocr_input_ext_map_hdmv_pgs_subtitle', ocr_ext_map.get('hdmv_pgs_subtitle', '.sup'))
        self.config.set('OCR', 'ocr_input_ext_map_dvd_subtitle', ocr_ext_map.get('dvd_subtitle', '.sub'))
        try:
            with open(config_path, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
        except IOError as e:
            self.log_message(f"Error writing configuration: {e}", to_console=True)

    def open_config_file(self):
        config_path = self.get_config_path()
        try:
            if os.name == 'nt': os.startfile(config_path)
            elif 'darwin' in sys.platform: subprocess.call(('open', config_path))
            elif 'linux' in sys.platform: subprocess.call(('xdg-open', config_path))
            else: messagebox.showinfo("Info", f"Please open this Holocron manually:\n{config_path}", parent=self.master)
            self.log_message(f"Opening Holocron (config file): {config_path}. Restart required for changes to take effect.", to_console=False)
        except Exception as e:
            self.log_message(f"Could not open Holocron: {e}", to_console=True)
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

        # OCR Enabled Checkbox
        ocr_enabled_var = tk.BooleanVar(value=self.settings.get('ocr_enabled', False))
        ttk.Checkbutton(content_frame, text="Enable OCR Droid (for image-based subtitles)",
                        variable=ocr_enabled_var, style="TCheckbutton").pack(anchor='w', pady=(0, 10))

        # OCR Command Template
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
                # Replace the first part (the executable) with the new path
                if len(parts) > 1 and os.path.isfile(parts[0]):
                    new_cmd = f'"{exe_path}" {" ".join(parts[1:])}'
                else:
                    new_cmd = f'"{exe_path}" {current_cmd}'
                ocr_cmd_var.set(new_cmd.strip())

        browse_btn = ttk.Button(cmd_frame, text="Browse...", command=browse_for_ocr_exe)
        browse_btn.pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Label(content_frame, text="Use placeholders: {INPUT_FILE_PATH}, {OUTPUT_SRT_PATH}, {LANG_3_CODE}",
                  font=("Helvetica", 8)).pack(anchor='w', pady=(0, 10))

        # OCR Default Language
        ttk.Label(content_frame, text="Default Language for OCR (3-letter code):").pack(anchor='w')
        ocr_lang_var = tk.StringVar(value=self.settings.get('ocr_default_lang', 'eng'))
        ttk.Entry(content_frame, textvariable=ocr_lang_var, width=10).pack(anchor='w', pady=(0, 15))

        # Dialog Buttons
        button_frame = ttk.Frame(content_frame)
        button_frame.pack(fill='x', side=tk.BOTTOM)

        def on_save():
            self.settings['ocr_enabled'] = ocr_enabled_var.get()
            self.settings['ocr_command_template'] = ocr_cmd_var.get()
            self.settings['ocr_default_lang'] = ocr_lang_var.get()
            self.save_config()
            self.log_message("OCR Holocron settings updated and saved.", to_console=False)
            dialog.destroy()

        ok_button = ttk.Button(button_frame, text="Save Protocol", command=on_save, style="Accent.TButton")
        ok_button.pack(side=tk.RIGHT, padx=5)
        cancel_button = ttk.Button(button_frame, text="Discard Changes", command=dialog.destroy)
        cancel_button.pack(side=tk.RIGHT)
        dialog.wait_window()

    def open_language_filter_dialog(self):
        if not self.movie_files_paths:
            messagebox.showinfo("No Targets", "Scan a star system (folder) first, Commander.", parent=self.master)
            return
        current_files_in_listbox = [self.movie_files_paths[i] for i in range(self.file_listbox.size()) if i < len(self.movie_files_paths)]
        if not current_files_in_listbox:
            messagebox.showinfo("No Targets", "No transmissions (files) in the list to scan for languages.", parent=self.master)
            return

        available_languages = set()
        lang_probe_dialog = tk.Toplevel(self.master)
        lang_probe_dialog.title("Scanning Transmissions")
        lang_probe_dialog.geometry("350x100")
        lang_probe_dialog.resizable(False, False)
        lang_probe_dialog.configure(bg=self.current_theme["bg"])
        ttk.Label(lang_probe_dialog, text="Scanning for alien languages (subtitle tracks)...").pack(pady=20, padx=10)
        lang_probe_dialog.transient(self.master)
        lang_probe_dialog.grab_set()
        self.master.update_idletasks()

        for file_path in current_files_in_listbox:
            try:
                cmd_probe = [self.settings['ffprobe_path'], '-v', 'error', '-show_entries', 'stream_tags=language', '-select_streams', 's', '-of', 'csv=p=0', file_path]
                process = subprocess.Popen(cmd_probe, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8', creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
                stdout, _ = process.communicate(timeout=self.settings['ffprobe_timeout'])
                if stdout:
                    for lang_code_line in stdout.strip().split('\n'):
                        lang_code = lang_code_line.strip()
                        if lang_code and len(lang_code) == 3: available_languages.add(lang_code.lower())
            except subprocess.TimeoutExpired: self.log_message(f"Comlink timeout probing languages in {os.path.basename(file_path)}", True)
            except Exception as e: self.log_message(f"Astromech droid malfunction probing languages in {os.path.basename(file_path)}: {e}", True)
        lang_probe_dialog.destroy()

        if not available_languages:
            messagebox.showinfo("No Languages", "No distinct alien languages found in current transmissions.", parent=self.master)
            self.extract_all_languages_flag = True
            self.user_selected_languages = set()
            self.current_lang_filter_label.config(text=self._get_current_lang_filter_display())
            return

        dialog = tk.Toplevel(self.master); dialog.title("Set Language Filters"); dialog.transient(self.master); dialog.grab_set(); dialog.configure(bg=self.current_theme["bg"]); dialog.minsize(350, 250)
        content_frame = ttk.Frame(dialog, padding=10); content_frame.pack(expand=True, fill=tk.BOTH)
        ttk.Label(content_frame, text="Select languages for translation (extraction):").pack(anchor='w', pady=(0, 5))
        scrollable_outer_frame = ttk.Frame(content_frame); scrollable_outer_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        canvas = tk.Canvas(scrollable_outer_frame, borderwidth=0, background=self.current_theme["list_bg"], highlightthickness=0)
        checkbox_frame_for_langs_in_canvas = ttk.Frame(canvas, style="TFrame"); vsb = ttk.Scrollbar(scrollable_outer_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set); vsb.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
        canvas_frame_id = canvas.create_window((0, 0), window=checkbox_frame_for_langs_in_canvas, anchor="nw")
        def on_checkbox_frame_configure(event): canvas.configure(scrollregion=canvas.bbox("all"))
        checkbox_frame_for_langs_in_canvas.bind("<Configure>", on_checkbox_frame_configure)
        def on_canvas_configure(event): canvas.itemconfig(canvas_frame_id, width=event.width)
        canvas.bind("<Configure>", on_canvas_configure)
        lang_vars = {}; sorted_langs = sorted(list(available_languages)); all_var = tk.BooleanVar(value=self.extract_all_languages_flag); lang_checkbutton_widgets = []
        def on_toggle_all_languages():
            is_all_selected = all_var.get()
            for cb_widget in lang_checkbutton_widgets: cb_widget.config(state=tk.DISABLED if is_all_selected else tk.NORMAL)
            if is_all_selected:
                for lang_code_key in lang_vars: lang_vars[lang_code_key].set(False)
            else:
                for lang_code_key, var_instance in lang_vars.items(): var_instance.set(lang_code_key in self.user_selected_languages)
        all_cb = ttk.Checkbutton(checkbox_frame_for_langs_in_canvas, text="Translate All (Galactic Basic Default)", variable=all_var, command=on_toggle_all_languages, style="TCheckbutton"); all_cb.pack(anchor='w', pady=2, padx=5)
        ttk.Separator(checkbox_frame_for_langs_in_canvas, orient='horizontal').pack(fill='x', pady=5, padx=5)
        for lang_code in sorted_langs:
            var = tk.BooleanVar();
            if not self.extract_all_languages_flag and lang_code in self.user_selected_languages: var.set(True)
            cb = ttk.Checkbutton(checkbox_frame_for_langs_in_canvas, text=lang_code, variable=var, style="TCheckbutton"); cb.pack(anchor='w', padx=10, pady=1)
            lang_vars[lang_code] = var; lang_checkbutton_widgets.append(cb)
        on_toggle_all_languages()
        button_frame = ttk.Frame(content_frame); button_frame.pack(fill='x', pady=(10, 0), side=tk.BOTTOM)
        def on_ok():
            self.extract_all_languages_flag = all_var.get()
            if self.extract_all_languages_flag: self.user_selected_languages = set()
            else: self.user_selected_languages = {lang for lang, var_cb in lang_vars.items() if var_cb.get()}
            lang_str_to_save = 'all' if self.extract_all_languages_flag or not self.user_selected_languages else ','.join(sorted(list(self.user_selected_languages)))
            self.settings['selected_languages'] = lang_str_to_save
            self.current_lang_filter_label.config(text=self._get_current_lang_filter_display())
            self.log_message(f"Language filter set to: {self._get_current_lang_filter_display()}", to_console=False); dialog.destroy()
        ok_button = ttk.Button(button_frame, text="Affirmative", command=on_ok, style="Accent.TButton"); ok_button.pack(side=tk.RIGHT, padx=5)
        cancel_button = ttk.Button(button_frame, text="Negative", command=dialog.destroy); cancel_button.pack(side=tk.RIGHT)
        dialog.wait_window()

    def _on_closing_main(self):
        self.settings['theme'] = self.current_theme_name; self.save_config()
        if self.log_window and self.log_window.winfo_exists(): self.log_window.destroy()
        self.master.destroy()

    def log_message(self, message, to_console=True):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        full_message = f"[{timestamp}] {message}"
        self.log_buffer.append(full_message + "\n")
        if self.log_window and self.log_window.winfo_exists() and self.log_text_widget:
            try:
                self.log_text_widget.config(state=tk.NORMAL); self.log_text_widget.insert(tk.END, full_message + "\n")
                self.log_text_widget.see(tk.END); self.log_text_widget.config(state=tk.DISABLED)
            except tk.TclError: pass
        if to_console: print(full_message)

    def open_log_window(self):
        if self.log_window and self.log_window.winfo_exists(): self.log_window.lift(); self.log_window.focus_set(); return
        self.log_window = tk.Toplevel(self.master); self.log_window.title("Mission Debrief (Log)"); self.log_window.geometry("700x500"); self.log_window.configure(bg=self.current_theme["bg"])
        log_button_frame = ttk.Frame(self.log_window); log_button_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        copy_button = ttk.Button(log_button_frame, text="Copy to Datapad", command=self.copy_log_to_clipboard); copy_button.pack(side=tk.LEFT, padx=5)
        save_button = ttk.Button(log_button_frame, text="Save Mission Log", command=self.save_log_to_file)
        if not self.log_dir_path: save_button.config(state=tk.DISABLED)
        save_button.pack(side=tk.LEFT, padx=5)
        self.log_text_widget = scrolledtext.ScrolledText(self.log_window, wrap=tk.WORD, state=tk.DISABLED, padx=5, pady=5); self.log_text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        if hasattr(self.log_text_widget, 'vbar'): self.log_text_widget.vbar.config(width=12)
        self.log_text_widget.config(state=tk.NORMAL)
        for msg in self.log_buffer: self.log_text_widget.insert(tk.END, msg)
        self.log_text_widget.see(tk.END); self.log_text_widget.config(state=tk.DISABLED)
        self.log_text_widget.configure(bg=self.current_theme["log_bg"], fg=self.current_theme["log_fg"], insertbackground=self.current_theme["fg"])
        for child in log_button_frame.winfo_children():
            if isinstance(child, ttk.Button): child.configure(style="TButton")
        self.log_window.protocol("WM_DELETE_WINDOW", self._on_closing_log_window)

    def copy_log_to_clipboard(self):
        if self.log_text_widget:
            log_content = self.log_text_widget.get("1.0", tk.END); self.master.clipboard_clear(); self.master.clipboard_append(log_content)
            self.log_message("Mission log copied to datapad.", to_console=False)
            parent_win = self.log_window if self.log_window and self.log_window.winfo_exists() else self.master
            messagebox.showinfo("Log Copied", "Mission log copied to datapad.", parent=parent_win)

    def save_log_to_file(self):
        if not self.log_dir_path: messagebox.showerror("Error", "Log archive directory not available.", parent=self.log_window if self.log_window and self.log_window.winfo_exists() else self.master); return
        if not self.log_buffer: messagebox.showinfo("Info", "Mission log is empty, Commander.", parent=self.log_window if self.log_window and self.log_window.winfo_exists() else self.master); return
        log_filename = f"sub_extractor_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        log_filepath = os.path.join(self.log_dir_path, log_filename)
        try:
            with open(log_filepath, "w", encoding="utf-8") as f: f.writelines(self.log_buffer)
            self.log_message(f"Mission log archived: {log_filepath}", to_console=False)
            parent_win = self.log_window if self.log_window and self.log_window.winfo_exists() else self.master
            messagebox.showinfo("Log Archived", f"Mission log archived to:\n{log_filepath}", parent=parent_win)
        except Exception as e:
            self.log_message(f"Error archiving mission log: {e}", to_console=True)
            parent_win = self.log_window if self.log_window and self.log_window.winfo_exists() else self.master
            messagebox.showerror("Archive Error", f"Failed to archive mission log:\n{e}", parent=parent_win)

    def _on_closing_log_window(self):
        if self.log_window: self.log_window.destroy(); self.log_window = None; self.log_text_widget = None

    def check_ffmpeg(self):
        ffmpeg_to_check = self.settings['ffmpeg_path']; ffprobe_to_check = self.settings['ffprobe_path']
        ffmpeg_found = shutil.which(ffmpeg_to_check) is not None; ffprobe_found = shutil.which(ffprobe_to_check) is not None
        if not ffmpeg_found: self.log_message(f"Warning: Hyperdrive motivator (FFmpeg: {ffmpeg_to_check}) offline or invalid coordinates.", to_console=True)
        if not ffprobe_found: self.log_message(f"Warning: Navigation computer (FFprobe: {ffprobe_to_check}) offline or invalid coordinates.", to_console=True)
        return ffmpeg_found and ffprobe_found

    def apply_theme(self):
        theme = self.current_theme; self.master.configure(bg=theme["bg"]); self.style.theme_use('clam')
        self.style.configure("TFrame", background=theme["bg"])
        self.style.configure("TLabel", background=theme["bg"], foreground=theme["fg"], padding=2)
        self.style.configure("TButton", background=theme["button_bg"], foreground=theme["button_fg"], padding=5, relief="flat", borderwidth=1)
        self.style.map("TButton", background=[('active', theme["accent_bg"]), ('pressed', theme["accent_bg"])], foreground=[('active', theme["accent_fg"]), ('pressed', theme["accent_fg"])])
        self.style.configure("Accent.TButton", background=theme["accent_bg"], foreground=theme["accent_fg"], font=('Helvetica', 10, 'bold'), padding=6, relief="flat", borderwidth=1)
        self.style.map("Accent.TButton", background=[('active', theme["button_bg"]), ('pressed', theme["button_bg"])], foreground=[('active', theme["button_fg"]), ('pressed', theme["button_fg"])])
        self.style.configure("Custom.Horizontal.TProgressbar", troughcolor=theme["progress_trough"], background=theme["progress_bg"], bordercolor=theme["bg"], lightcolor=theme["bg"], darkcolor=theme["bg"])
        self.style.configure("TCombobox", selectbackground=theme["list_bg"], fieldbackground=theme["list_bg"], background=theme["button_bg"], foreground=theme["fg"])
        self.style.map('TCombobox', fieldbackground=[('readonly', theme["list_bg"])], selectbackground=[('readonly', theme["accent_bg"])], selectforeground=[('readonly', theme["accent_fg"])])
        self.style.configure("TCheckbutton", background=theme["bg"], foreground=theme["fg"])
        self.style.map("TCheckbutton", background=[('active', theme["bg"])], indicatorcolor=[('selected', theme["accent_bg"]), ('!selected', theme["button_bg"])], foreground=[('disabled', theme.get("disabled_fg", "#A0A0A0"))])
        self.file_listbox.configure(bg=theme["list_bg"], fg=theme["list_fg"], selectbackground=theme["accent_bg"], selectforeground=theme["accent_fg"], highlightbackground=theme["bg"], highlightcolor=theme["accent_bg"])
        self.folder_label.configure(background=theme["bg"], foreground=theme["fg"]); self.status_label.configure(background=theme["bg"], foreground=theme["fg"])
        self.theme_button.config(text="Join the Light Side" if self.is_dark_mode else "Embrace the Dark Side") # Updated Star Wars theme toggle
        if self.log_window and self.log_window.winfo_exists():
            self.log_window.configure(bg=theme["bg"])
            if self.log_text_widget: self.log_text_widget.configure(bg=theme["log_bg"], fg=theme["log_fg"], insertbackground=theme["fg"])
            for child_frame in self.log_window.winfo_children():
                if isinstance(child_frame, ttk.Frame):
                    for btn in child_frame.winfo_children():
                        if isinstance(btn, ttk.Button): btn.configure(style="TButton")
        self.select_folder_button.configure(style="TButton"); self.remove_button.configure(style="TButton"); self.view_log_button.configure(style="TButton")
        self.theme_button.configure(style="TButton"); self.edit_config_button.configure(style="TButton"); self.select_langs_button.configure(style="TButton")
        self.ocr_settings_button.configure(style="TButton") # Theme new button
        self.skip_checkbox.configure(style="TCheckbutton")
        self.extract_button.configure(style="Accent.TButton")

    def toggle_theme(self):
        if self.is_dark_mode: self.current_theme, self.current_theme_name, self.is_dark_mode = LIGHT_THEME, "light", False
        else: self.current_theme, self.current_theme_name, self.is_dark_mode = DARK_THEME, "dark", True
        self.settings['theme'] = self.current_theme_name; self.apply_theme()

    def select_folder(self):
        initial_dir = self.settings['last_folder'] if self.settings['last_folder'] and os.path.isdir(self.settings['last_folder']) else None
        folder_path = filedialog.askdirectory(initialdir=initial_dir)
        if folder_path:
            self.folder_label.config(text=folder_path); self.settings['last_folder'] = folder_path
            self.log_message(f"Selected star system (folder): {folder_path}", to_console=False); self.scan_folder(folder_path)

    def scan_folder(self, folder_path):
        self.file_listbox.delete(0, tk.END); self.movie_files_paths = []
        self.status_label.config(text=f"Scanning {os.path.basename(folder_path)} sector..."); self.log_message(f"Scanning sector: {folder_path}...", to_console=False)
        self.master.update_idletasks(); found_count = 0
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith(MOVIE_EXTENSIONS):
                    full_path = os.path.join(root, file); self.movie_files_paths.append(full_path)
                    self.file_listbox.insert(tk.END, os.path.basename(file)); found_count += 1
        msg = f"Found {found_count} transmissions (movie files)." if found_count > 0 else "No transmissions detected in this sector."
        self.status_label.config(text=msg); self.log_message(msg, to_console=False)

    def remove_selected_files(self):
        selected_indices = self.file_listbox.curselection()
        if not selected_indices: messagebox.showinfo("Info", "No targets selected for removal, Commander.", parent=self.master); return
        removed_count = 0
        for i in sorted(selected_indices, reverse=True):
            self.log_message(f"Removing from target list: {os.path.basename(self.movie_files_paths[i])}", to_console=False)
            self.file_listbox.delete(i); del self.movie_files_paths[i]; removed_count += 1
        msg = f"{removed_count} target(s) removed from list. {len(self.movie_files_paths)} remaining."
        self.status_label.config(text=msg); self.log_message(msg, to_console=False)

    def start_extraction_thread(self):
        files_to_process_paths = list(self.movie_files_paths)
        if not files_to_process_paths: messagebox.showinfo("No Targets", "No targets acquired. Scan a system and select files.", parent=self.master); return

        self._toggle_extraction_controls(is_extracting=True)
        # Clear all reporting lists and log for the new run
        self.log_buffer.clear()
        self.files_with_success.clear(); self.files_with_no_subs.clear(); self.files_timed_out.clear(); self.files_with_errors.clear(); self.files_skipped.clear()
        self.progress_var.set(0)
        if self.log_window and self.log_window.winfo_exists() and self.log_text_widget:
            self.log_text_widget.config(state=tk.NORMAL); self.log_text_widget.delete('1.0', tk.END); self.log_text_widget.config(state=tk.DISABLED)
        self.log_message("--- Starting New Extraction Mission ---", to_console=False)
        thread = threading.Thread(target=self._extract_subtitles_logic, args=(files_to_process_paths,), daemon=True); thread.start()

    def _cancel_extraction(self):
        self.log_message("--- MISSION ABORT SIGNAL RECEIVED ---", to_console=True)
        self.status_label.config(text="Cancelling mission... Please wait for the current target to finish.")
        self.cancel_requested.set()
        # Disable the cancel button to prevent multiple clicks
        self.extract_button.config(state=tk.DISABLED, text="Cancelling...")

    def _toggle_extraction_controls(self, is_extracting):
        """Enable/disable controls and toggle the extract/cancel button."""
        if is_extracting:
            self.cancel_requested.clear()
            self.extract_button.config(text="Cancel Extraction", command=self._cancel_extraction)
            for btn in [self.remove_button, self.select_folder_button, self.select_langs_button, self.ocr_settings_button]:
                btn.config(state=tk.DISABLED)
        else:
            self.extract_button.config(text="Extract Subtitles", command=self.start_extraction_thread, state=tk.NORMAL)
            for btn in [self.remove_button, self.select_folder_button, self.select_langs_button, self.ocr_settings_button]:
                btn.config(state=tk.NORMAL)

    def _update_status_safe(self, message):
        self.master.after(0, lambda: self.status_label.config(text=message))

    def _update_progress_safe(self, value):
        self.master.after(0, lambda: self.progress_var.set(value))

    def _extraction_finished_safe(self, summary_message=None):
        if self.cancel_requested.is_set():
            summary_message = "Mission aborted by user."
            self.log_message("\n--- MISSION ABORTED BY USER ---", to_console=True)
        else:
            final_log_summary = ["\n--- MISSION DEBRIEF ---", summary_message or "Mission completed, Commander."]
            if self.files_with_success:
                final_log_summary.append("\nTransmissions successfully decoded from:")
                final_log_summary.extend([f"- {f}" for f in self.files_with_success])
            if self.files_skipped:
                final_log_summary.append("\nTargets bypassed (existing subtitles):")
                final_log_summary.extend([f"- {f}" for f in self.files_skipped])
            if self.files_with_no_subs:
                final_log_summary.append("\nTransmissions with no subtitle signals:")
                final_log_summary.extend([f"- {f}" for f in self.files_with_no_subs])
            if self.files_timed_out:
                final_log_summary.append("\nTransmissions lost in hyperspace (timed out):")
                final_log_summary.extend([f"- {f}" for f in self.files_timed_out])
            if self.files_with_errors:
                final_log_summary.append("\nTransmissions corrupted (errors):")
                final_log_summary.extend([f"- {f}" for f in self.files_with_errors])
            for line in final_log_summary: self.log_message(line, to_console=False)
            self.log_message("\n--- End of Mission Log ---", to_console=True)

        self._update_status_safe(summary_message or "Mission accomplished!"); self.progress_var.set(0)
        messagebox.showinfo("Mission Complete", (summary_message or "Extraction run complete, Commander.") + "\n\nCheck Mission Debrief (Log) for details.", parent=self.master)
        self._toggle_extraction_controls(is_extracting=False)

    def _check_for_existing_subs(self, movie_file_path):
        """Checks if a subtitle file already exists for the given movie file."""
        movie_dir = os.path.dirname(movie_file_path)
        base_name_no_ext = os.path.splitext(os.path.basename(movie_file_path))[0]
        try:
            for filename in os.listdir(movie_dir):
                if filename.startswith(base_name_no_ext) and filename.lower().endswith(SUBTITLE_EXTENSIONS):
                    return True
        except FileNotFoundError:
            self.log_message(f"[WARN] Directory not found while checking for existing subs: {movie_dir}", to_console=True)
            return False
        return False

    def _run_ocr_on_image_sub(self, movie_file_path, base_name_no_ext, stream_idx, lang_code, input_codec, target_srt_path):
        filename_short = os.path.basename(movie_file_path)
        witty_ocr_message = random.choice(OCR_PATIENCE_MESSAGES).format(filename=filename_short)
        self._update_status_safe(witty_ocr_message)
        self.log_message(f"[OCR] Attempting OCR for stream {stream_idx} ({input_codec}, lang {lang_code}) from {filename_short}", to_console=True)

        temp_dir_base = self.settings.get('ocr_temp_dir', '') or os.path.dirname(movie_file_path)
        ocr_session_temp_dir = tempfile.mkdtemp(prefix=f"ocr_{base_name_no_ext}_s{stream_idx}_", dir=temp_dir_base if os.path.isdir(temp_dir_base) else None)
        image_sub_ext = self.settings['ocr_input_ext_map'].get(input_codec, f".{input_codec}")
        temp_image_sub_basename = f"{base_name_no_ext}_s{stream_idx}_temp{image_sub_ext}"
        temp_image_sub_path = os.path.join(ocr_session_temp_dir, temp_image_sub_basename)

        cmd_extract_image = [self.settings['ffmpeg_path'], '-y', '-analyzeduration', '100M', '-probesize', '100M', '-i', movie_file_path, '-map', f'0:{stream_idx}', '-c:s', 'copy', temp_image_sub_path]
        self.log_message(f"[OCR FFmpeg CMD] {' '.join(cmd_extract_image)}")
        extract_proc = subprocess.Popen(cmd_extract_image, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8', creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
        _, ext_stderr = extract_proc.communicate(timeout=self.settings['ffmpeg_extract_timeout'])
        if ext_stderr and ext_stderr.strip():
            self.log_message(f"[OCR FFmpeg STDERR for {temp_image_sub_basename}]:\n{ext_stderr.strip()}")
            if "file ended prematurely" in ext_stderr.lower():
                 self.log_message("[INFO] Note: The 'file ended prematurely' message is often non-critical for temporary image subtitle extraction.", to_console=False)

        if extract_proc.returncode != 0 or not os.path.exists(temp_image_sub_path) or os.path.getsize(temp_image_sub_path) == 0:
            self.log_message(f"[OCR ERROR] Failed to extract temporary image subtitle or file is empty: {temp_image_sub_basename}. FFmpeg RC: {extract_proc.returncode}.", to_console=True)
            if os.path.isdir(ocr_session_temp_dir): shutil.rmtree(ocr_session_temp_dir, ignore_errors=True)
            return False

        temp_ocr_output_srt_path = os.path.join(ocr_session_temp_dir, f"{os.path.splitext(temp_image_sub_basename)[0]}.srt")
        ocr_command_raw = self.settings['ocr_command_template']
        safe_lang_code_for_ocr = lang_code if len(lang_code) == 3 else self.settings.get('ocr_default_lang', 'eng')
        ocr_command_str = ocr_command_raw.replace("{INPUT_FILE_PATH}", f'"{temp_image_sub_path}"')
        ocr_command_str = ocr_command_str.replace("{OUTPUT_SRT_PATH}", f'"{temp_ocr_output_srt_path}"')
        if "{LANG_3_CODE}" in ocr_command_raw: ocr_command_str = ocr_command_str.replace("{LANG_3_CODE}", safe_lang_code_for_ocr)
        elif "/ocrlanguage:" in ocr_command_str.lower() and "{LANG_3_CODE}" not in ocr_command_raw:
            self.log_message(f"[OCR WARNING] Obsolete '/ocrlanguage:' detected. Review Holocron (Config). Attempting to proceed without.", to_console=True)
            ocr_command_str = re.sub(r'/ocrlanguage:[a-zA-Z0-9_]+', '', ocr_command_str, flags=re.IGNORECASE).strip()
            ocr_command_str = re.sub(r'\s\s+', ' ', ocr_command_str)

        self.log_message(f"[OCR CMD] {ocr_command_str}", to_console=True)
        ocr_success = False
        try:
            import shlex
            ocr_proc = subprocess.run(shlex.split(ocr_command_str), capture_output=True, text=True, encoding='utf-8', timeout=self.settings['ffmpeg_ocr_timeout'], creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0, check=False)
            if ocr_proc.stdout and ocr_proc.stdout.strip(): self.log_message(f"[OCR STDOUT]:\n{ocr_proc.stdout.strip()}")
            if ocr_proc.stderr and ocr_proc.stderr.strip(): self.log_message(f"[OCR STDERR]:\n{ocr_proc.stderr.strip()}")
            self.log_message(f"[OCR RETURN CODE]: {ocr_proc.returncode}")
            if ocr_proc.returncode == 0 and os.path.exists(temp_ocr_output_srt_path) and os.path.getsize(temp_ocr_output_srt_path) > 0:
                shutil.move(temp_ocr_output_srt_path, target_srt_path)
                self.log_message(f"[OCR SUCCESS] Translation complete: {os.path.basename(target_srt_path)}", to_console=True); ocr_success = True
            elif ocr_proc.returncode == 0: self.log_message(f"[OCR FAILED] Droid translation unit (RC 0) but output datapad (SRT) is empty/missing: {temp_ocr_output_srt_path}", to_console=True)
            else: self.log_message(f"[OCR FAILED] Droid translation unit malfunctioned (RC {ocr_proc.returncode}).", to_console=True)
        except subprocess.TimeoutExpired: self.log_message(f"[OCR TIMEOUT] Comlink lost with OCR droid for {temp_image_sub_basename} after {self.settings['ffmpeg_ocr_timeout']}s.", to_console=True)
        except FileNotFoundError: self.log_message(f"[OCR ERROR] OCR Droid (tool) not found. Check Holocron (Config) for: {shlex.split(ocr_command_str)[0]}", to_console=True)
        except Exception as e:
            self.log_message(f"[OCR CRITICAL ERROR] Catastrophic droid failure during OCR: {e}", to_console=True); import traceback; self.log_message(traceback.format_exc(), to_console=True)
        finally:
            if os.path.isdir(ocr_session_temp_dir): shutil.rmtree(ocr_session_temp_dir, ignore_errors=True); self.log_message(f"[OCR Cleanup] Erased temporary droid memory banks: {ocr_session_temp_dir}", to_console=False)
        return ocr_success

    def _extract_subtitles_logic(self, files_to_process):
        total_files = len(files_to_process); overall_subs_extracted_count = 0; processed_for_progress_count = 0
        selected_gui_output_format = self.output_format_var.get()
        self.log_message(f"Using output format: {selected_gui_output_format}", to_console=True)
        self.log_message(f"Language filter: {self._get_current_lang_filter_display()}", to_console=True)
        if self.settings.get('ocr_enabled') and self.settings.get('ocr_command_template'):
            self.log_message(f"[OCR STATUS] OCR Droid is ONLINE. Protocol: {self.settings['ocr_command_template'][:50]}...", to_console=True)
        else:
            self.log_message("[OCR STATUS] OCR Droid OFFLINE or no protocol. Image subs will be copied or skipped (if text output chosen).", to_console=True)

        for i, movie_file_path in enumerate(files_to_process):
            if self.cancel_requested.is_set():
                break # Exit loop if cancellation was requested

            movie_filename = os.path.basename(movie_file_path); current_file_had_error_flag = False
            general_status_for_file = f"Scanning target ({i + 1}/{total_files}): {movie_filename}"
            self._update_status_safe(general_status_for_file)
            self._update_progress_safe((processed_for_progress_count / total_files) * 100 if total_files > 0 else 0)
            
            if self.settings.get('skip_if_exists'):
                if self._check_for_existing_subs(movie_file_path):
                    self.log_message(f"\n[INFO] Skipping target ({i + 1}/{total_files}): {movie_filename} - Existing subtitle file found.")
                    self.files_skipped.append(movie_filename)
                    processed_for_progress_count += 1
                    self._update_progress_safe((processed_for_progress_count / total_files) * 100 if total_files > 0 else 0)
                    continue

            self.log_message(f"\n[INFO] Processing target ({i + 1}/{total_files}): {movie_file_path}")
            movie_dir = os.path.dirname(movie_file_path); base_name_no_ext = os.path.splitext(movie_filename)[0]
            file_subs_extracted_this_file = 0; file_processed_or_skipped_this_iteration = False
            try:
                cmd_probe = [self.settings['ffprobe_path'], '-v', 'error', '-show_entries', 'stream=index,codec_type,codec_name:stream_tags=language', '-select_streams', 's', '-of', 'csv=p=0', movie_file_path]
                self.log_message(f"[FFPROBE CMD] {' '.join(cmd_probe)}")
                probe_process = subprocess.Popen(cmd_probe, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8', creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
                stdout, stderr = probe_process.communicate(timeout=self.settings['ffprobe_timeout'])
                self.log_message(f"[FFPROBE STDOUT for {movie_filename}]:\n{stdout.strip() if stdout.strip() else '<no subtitle signals detected>'}")
                if stderr and stderr.strip(): self.log_message(f"[FFPROBE STDERR for {movie_filename}]:\n{stderr.strip()}")
                self.log_message(f"[FFPROBE RETURN CODE for {movie_filename}]: {probe_process.returncode}")
                if probe_process.returncode != 0:
                    self.log_message(f"[ERROR] FFprobe malfunctioned for {movie_filename}. RC: {probe_process.returncode}. Aborting target.", to_console=True)
                    if not current_file_had_error_flag: self.files_with_errors.append(movie_filename); current_file_had_error_flag = True
                    file_processed_or_skipped_this_iteration = True; continue
                subtitle_streams_from_probe = []
                if stdout.strip():
                    lines = stdout.strip().split('\n')
                    for line_content in lines:
                        parts = line_content.strip().split(',');
                        if len(parts) < 3: continue
                        stream_index_str, codec_type_from_probe, codec_name_from_probe = parts[0].strip(), parts[1].strip().lower(), parts[2].strip().lower()
                        actual_codec_name = "unknown"; is_subtitle_stream = False
                        if codec_type_from_probe == 'subtitle': is_subtitle_stream = True; actual_codec_name = codec_name_from_probe
                        elif codec_name_from_probe == 'subtitle': is_subtitle_stream = True; actual_codec_name = codec_type_from_probe
                        elif codec_type_from_probe in IMAGE_BASED_CODECS or codec_type_from_probe in TEXT_BASED_OUTPUT_FORMATS: is_subtitle_stream = True; actual_codec_name = codec_type_from_probe
                        if is_subtitle_stream:
                            language_str = "und";
                            if len(parts) > 3 and parts[3].strip(): language_str = parts[3].strip().lower()
                            subtitle_streams_from_probe.append({"index": stream_index_str, "lang": language_str, "codec": actual_codec_name})
                            self.log_message(f"[DEBUG]   -> Detected signal: Idx='{stream_index_str}',Lang='{language_str}',Codec='{actual_codec_name}'")
                if not subtitle_streams_from_probe:
                    self.log_message(f"[INFO] No subtitle signals found/parsed for {movie_filename}."); self.files_with_no_subs.append(movie_filename)
                    file_processed_or_skipped_this_iteration = True; continue
                streams_to_extract_this_file = []
                if self.extract_all_languages_flag: streams_to_extract_this_file = subtitle_streams_from_probe
                else:
                    for stream_info in subtitle_streams_from_probe:
                        if stream_info["lang"] in self.user_selected_languages: streams_to_extract_this_file.append(stream_info)
                self.log_message(f"[INFO] Filtered to {len(streams_to_extract_this_file)} signal(s) for {movie_filename} based on language selection: {self.user_selected_languages if not self.extract_all_languages_flag else 'All (Galactic Basic)'}")
                if not streams_to_extract_this_file:
                    self.log_message(f"[INFO] No signals match language filter for {movie_filename}. Skipping this target's subtitle extraction.")
                    file_processed_or_skipped_this_iteration = True; continue

                self._update_status_safe(f"Processing subtitle signals for {movie_filename}...")

                for stream_info in streams_to_extract_this_file:
                    if self.cancel_requested.is_set(): break
                    
                    stream_idx, lang_code, input_codec = stream_info["index"], stream_info["lang"], stream_info["codec"].lower()
                    safe_lang_code = re.sub(r'[^a-zA-Z0-9_.-]', '', lang_code) or "und"
                    output_target_format_gui = selected_gui_output_format.lower()
                    ffmpeg_codec_arg_for_direct_extract = output_target_format_gui; final_output_extension = f".{output_target_format_gui}"
                    run_ocr = False; use_direct_ffmpeg_extract = True
                    if output_target_format_gui == 'copy':
                        ffmpeg_codec_arg_for_direct_extract = 'copy'
                        if input_codec in ['subrip', 'srt']: final_output_extension = ".srt"
                        elif input_codec == 'ass': final_output_extension = ".ass"
                        elif input_codec in ['webvtt', 'vtt']: final_output_extension = ".vtt"
                        elif input_codec == 'mov_text': ffmpeg_codec_arg_for_direct_extract = 'srt'; final_output_extension = '.srt'; self.log_message(f"[INFO] Forcing mov_text (signal {stream_idx}, lang {lang_code}) to SRT for comlink compatibility, despite 'copy' order.", to_console=True)
                        elif input_codec in IMAGE_BASED_CODECS: final_output_extension = self.settings['ocr_input_ext_map'].get(input_codec, f".{input_codec}"); self.log_message(f"[INFO] Copying image-based signal '{input_codec}' (stream {stream_idx}, lang {lang_code}) as is. Output ext: {final_output_extension}", to_console=True)
                        else: final_output_extension = f".{input_codec}"; self.log_message(f"[WARN] Copying unknown signal type '{input_codec}' (stream {stream_idx}, lang {lang_code}). Extension: '{final_output_extension}'.", to_console=True)
                    elif output_target_format_gui in TEXT_BASED_OUTPUT_FORMATS: # srt, ass, vtt
                        if input_codec in IMAGE_BASED_CODECS:
                            if self.settings.get('ocr_enabled') and self.settings.get('ocr_command_template'):
                                run_ocr = True; use_direct_ffmpeg_extract = False
                            else:
                                self.log_message(f"[INFO] Skipping image-based signal {stream_idx} ({input_codec}, lang {lang_code}) for {movie_filename}. Cannot convert to {output_target_format_gui.upper()} without OCR Droid. Use 'copy' or deploy OCR Droid via Holocron (Config).", to_console=True)
                                if not current_file_had_error_flag: self.files_with_errors.append(movie_filename); current_file_had_error_flag = True
                                continue # Skip this specific stream
                        elif input_codec == 'mov_text':
                            self.log_message(f"[INFO] Translating mov_text (signal {stream_idx}, lang {lang_code}) to {output_target_format_gui.upper()}.", to_console=True)
                    else:
                        self.log_message(f"[ERROR] Unexpected output format '{output_target_format_gui}' for signal {stream_idx}. Skipping.", to_console=True)
                        if not current_file_had_error_flag: self.files_with_errors.append(movie_filename); current_file_had_error_flag = True
                        continue

                    sub_filename_out = f"{base_name_no_ext}.{safe_lang_code}.{stream_idx}{final_output_extension}"
                    output_path = os.path.join(movie_dir, sub_filename_out)
                    extraction_successful_this_stream = False
                    if run_ocr:
                        if self._run_ocr_on_image_sub(movie_file_path, base_name_no_ext, stream_idx, lang_code, input_codec, output_path): extraction_successful_this_stream = True
                        else:
                            if not current_file_had_error_flag: self.files_with_errors.append(movie_filename); current_file_had_error_flag = True
                        self._update_status_safe(f"OCR Droid finished with {safe_lang_code} for {movie_filename}. Stand by...")
                    elif use_direct_ffmpeg_extract:
                        self._update_status_safe(f"Extracting signal {safe_lang_code} (idx {stream_idx}) as {ffmpeg_codec_arg_for_direct_extract.upper()} from {movie_filename}...")
                        cmd_extract = [self.settings['ffmpeg_path'], '-y', '-analyzeduration', '100M', '-probesize', '100M', '-i', movie_file_path, '-map', f'0:{stream_idx}', '-c:s', ffmpeg_codec_arg_for_direct_extract, output_path]
                        self.log_message(f"[FFMPEG CMD] {' '.join(cmd_extract)}")
                        extract_process = subprocess.Popen(cmd_extract, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8', creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
                        _, ext_stderr = extract_process.communicate(timeout=self.settings['ffmpeg_extract_timeout'])
                        if ext_stderr and ext_stderr.strip():
                            self.log_message(f"[FFMPEG STDERR for {sub_filename_out}]:\n{ext_stderr.strip()}")
                            if "file ended prematurely" in ext_stderr.lower():
                                self.log_message("[INFO] Note: The 'file ended prematurely' message from FFmpeg is often non-critical for subtitle streams and may not indicate a failure.", to_console=False)
                        self.log_message(f"[FFMPEG RETURN CODE for {sub_filename_out}]: {extract_process.returncode}")
                        if extract_process.returncode == 0 and os.path.exists(output_path) and os.path.getsize(output_path) > 0: extraction_successful_this_stream = True
                        elif extract_process.returncode == 0: self.log_message(f"[WARNING] FFmpeg reported success, but output datapad '{output_path}' is empty or missing.", to_console=True)
                    else:
                        self.log_message(f"[ERROR] Internal logic error for signal {stream_idx}. Cannot determine extraction method. Skipping.", to_console=True)
                        if not current_file_had_error_flag: self.files_with_errors.append(movie_filename); current_file_had_error_flag = True
                        continue

                    if extraction_successful_this_stream:
                        file_subs_extracted_this_file += 1
                        self.log_message(f"[SUCCESS] Successfully decoded stream {stream_idx} ({lang_code}) from {movie_filename} to {os.path.basename(output_path)}", to_console=True)
                
                if file_subs_extracted_this_file > 0:
                    overall_subs_extracted_count += file_subs_extracted_this_file
                    self.files_with_success.append(movie_filename)
                    self.log_message(f"[INFO] Target {movie_filename} processed, {file_subs_extracted_this_file} signal(s) decoded.")
                elif not current_file_had_error_flag and movie_filename not in self.files_with_no_subs and movie_filename not in self.files_with_errors:
                    self.log_message(f"[INFO] Target {movie_filename} processed, no suitable signals decoded/translated.")
                file_processed_or_skipped_this_iteration = True

            except subprocess.TimeoutExpired:
                self.log_message(f"[TIMEOUT] Comlink lost processing {movie_filename}. Skipping target.", to_console=True); self.files_timed_out.append(movie_filename); file_processed_or_skipped_this_iteration = True
                self._update_status_safe(f"Comlink lost with {movie_filename}. Moving to next target.")
            except Exception as e:
                self.log_message(f"[CRITICAL SYSTEM ERROR] Unexpected asteroid field encountered with {movie_filename}: {e}", to_console=True); import traceback; self.log_message(traceback.format_exc(), to_console=True)
                if not current_file_had_error_flag: self.files_with_errors.append(movie_filename)
                file_processed_or_skipped_this_iteration = True
                self._update_status_safe(f"Error with {movie_filename}. Jumping to next system.")
            finally:
                if file_processed_or_skipped_this_iteration: processed_for_progress_count += 1
                self._update_progress_safe((processed_for_progress_count / total_files) * 100 if total_files > 0 else 0)
        
        summary_message = f"Mission Report: {processed_for_progress_count}/{total_files} targets engaged. "
        if overall_subs_extracted_count > 0: summary_message += f"{overall_subs_extracted_count} subtitle signal(s) successfully decoded."
        else: summary_message += "No subtitle signals were decoded in this operation."
        self._extraction_finished_safe(summary_message)

if __name__ == "__main__":
    root = tk.Tk()
    app = SubtitleExtractorApp(root)
    root.mainloop()
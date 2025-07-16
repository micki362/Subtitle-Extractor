import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import subprocess
import threading
import re
import shutil
import datetime
import tempfile
import random
import ctypes
from config import AppConfig, LIGHT_THEME, DARK_THEME, OCR_PATIENCE_MESSAGES, MOVIE_EXTENSIONS, SUBTITLE_EXTENSIONS, IMAGE_BASED_CODECS, TEXT_BASED_OUTPUT_FORMATS
from ui import SubtitleExtractorUI

class SubtitleExtractorApp:
    def __init__(self, master):
        self.master = master
        self.app_dir = os.path.dirname(os.path.abspath(__file__))
        self.config = AppConfig(self.app_dir)
        self.settings = self.config.settings

        self._setup_theme()

        self.extract_all_languages_flag = True
        self.user_selected_languages = set()
        self._parse_loaded_languages()

        self.movie_files_paths = []
        self.files_with_success, self.files_with_no_subs, self.files_timed_out, self.files_with_errors, self.files_skipped = [], [], [], [], []
        self.log_buffer, self.log_window, self.log_text_widget = [], None, None
        self.cancel_requested = threading.Event()

        self._setup_logging()

        if not self.check_ffmpeg():
            messagebox.showerror("Error", f"FFmpeg/FFprobe not found (see 'sub_extractor_settings.ini').\nCheck paths and restart.")
            master.destroy()
            return

        self.ui = SubtitleExtractorUI(master, self)

    def _setup_theme(self):
        self.current_theme_name = self.settings['theme']
        self.is_dark_mode = (self.current_theme_name == 'dark')
        self.current_theme = DARK_THEME if self.is_dark_mode else LIGHT_THEME

    def _setup_logging(self):
        self.log_dir_path = os.path.join(self.app_dir, "logs")
        if not os.path.exists(self.log_dir_path):
            try:
                os.makedirs(self.log_dir_path)
            except OSError as e:
                print(f"Error creating log dir: {e}")
                self.log_dir_path = None

    def on_skip_toggle(self):
        self.settings['skip_if_exists'] = self.ui.skip_if_exists_var.get()
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
        selected_format = self.ui.output_format_var.get()
        self.settings['default_output_format'] = selected_format
        self.log_message(f"Output format set to: {selected_format}", to_console=False)

    def _get_current_lang_filter_display(self):
        if self.extract_all_languages_flag or not self.user_selected_languages:
            return "All Languages (Galactic Basic)"
        display_langs = sorted(list(self.user_selected_languages))
        if len(display_langs) > 3:
            return f"{', '.join(display_langs[:3])}, ..."
        return ', '.join(display_langs)

    def open_language_filter_dialog(self):
        if not self.movie_files_paths:
            messagebox.showinfo("No Targets", "Scan a star system (folder) first, Commander.", parent=self.master)
            return
        current_files_in_listbox = [self.movie_files_paths[i] for i in range(self.ui.file_listbox.size()) if i < len(self.movie_files_paths)]
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
            self.ui.current_lang_filter_label.config(text=self._get_current_lang_filter_display())
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
            self.ui.current_lang_filter_label.config(text=self._get_current_lang_filter_display())
            self.log_message(f"Language filter set to: {self._get_current_lang_filter_display()}", to_console=False); dialog.destroy()
        ok_button = ttk.Button(button_frame, text="Affirmative", command=on_ok, style="Accent.TButton"); ok_button.pack(side=tk.RIGHT, padx=5)
        cancel_button = ttk.Button(button_frame, text="Negative", command=dialog.destroy); cancel_button.pack(side=tk.RIGHT)
        dialog.wait_window()

    def _on_closing_main(self):
        self.settings['theme'] = self.current_theme_name; self.config.save_config(self.extract_all_languages_flag, self.user_selected_languages)
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

    def toggle_theme(self):
        if self.is_dark_mode: self.current_theme, self.current_theme_name, self.is_dark_mode = LIGHT_THEME, "light", False
        else: self.current_theme, self.current_theme_name, self.is_dark_mode = DARK_THEME, "dark", True
        self.settings['theme'] = self.current_theme_name; self.ui.apply_theme()

    def select_folder(self):
        initial_dir = self.settings['last_folder'] if self.settings['last_folder'] and os.path.isdir(self.settings['last_folder']) else None
        folder_path = filedialog.askdirectory(initialdir=initial_dir)
        if folder_path:
            self.ui.folder_label.config(text=folder_path); self.settings['last_folder'] = folder_path
            self.log_message(f"Selected star system (folder): {folder_path}", to_console=False); self.scan_folder(folder_path)

    def scan_folder(self, folder_path):
        self.ui.file_tree.delete(*self.ui.file_tree.get_children())
        self.movie_files_paths = []
        self.ui.status_label.config(text=f"Scanning {os.path.basename(folder_path)} sector..."); self.log_message(f"Scanning sector: {folder_path}...", to_console=False)
        self.master.update_idletasks(); found_count = 0
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith(MOVIE_EXTENSIONS):
                    full_path = os.path.join(root, file)
                    self.movie_files_paths.append(full_path)
                    status = "Subtitles Present" if self._check_for_existing_subs(full_path) else "Ready to Extract"
                    self.ui.file_tree.insert("", tk.END, values=(os.path.basename(file), status), iid=full_path)
                    found_count += 1
        msg = f"Found {found_count} transmissions (movie files)." if found_count > 0 else "No transmissions detected in this sector."
        self.ui.status_label.config(text=msg); self.log_message(msg, to_console=False)

    def remove_selected_files(self):
        selected_items = self.ui.file_tree.selection()
        if not selected_items: messagebox.showinfo("Info", "No targets selected for removal, Commander.", parent=self.master); return
        removed_count = 0
        for item in selected_items:
            self.log_message(f"Removing from target list: {self.ui.file_tree.item(item)['values'][0]}", to_console=False)
            self.ui.file_tree.delete(item)
            removed_count += 1
        msg = f"{removed_count} target(s) removed from list. {len(self.ui.file_tree.get_children())} remaining."
        self.ui.status_label.config(text=msg); self.log_message(msg, to_console=False)

    def remove_files_with_subtitles(self):
        items_to_remove = []
        for item in self.ui.file_tree.get_children():
            if self.ui.file_tree.item(item)['values'][1] == "Subtitles Present":
                items_to_remove.append(item)
        
        if not items_to_remove:
            messagebox.showinfo("Info", "No files with existing subtitles found to remove.", parent=self.master)
            return

        removed_count = 0
        for item in items_to_remove:
            self.log_message(f"Removing file with subtitles: {self.ui.file_tree.item(item)['values'][0]}", to_console=False)
            self.ui.file_tree.delete(item)
            removed_count += 1
        
        msg = f"{removed_count} file(s) with subtitles removed from the list."
        self.ui.status_label.config(text=msg)
        self.log_message(msg, to_console=False)

    def start_extraction_thread(self):
        files_to_process_paths = [item for item in self.ui.file_tree.get_children()]
        
        if not files_to_process_paths:
            messagebox.showinfo("No Targets", "No targets acquired. Scan a system and select files.", parent=self.master)
            return

        self._toggle_extraction_controls(is_extracting=True)
        self.log_buffer.clear()
        self.files_with_success.clear(); self.files_with_no_subs.clear(); self.files_timed_out.clear(); self.files_with_errors.clear(); self.files_skipped.clear()
        self.ui.progress_var.set(0)
        if self.log_window and self.log_window.winfo_exists() and self.log_text_widget:
            self.log_text_widget.config(state=tk.NORMAL); self.log_text_widget.delete('1.0', tk.END); self.log_text_widget.config(state=tk.DISABLED)
        self.log_message("--- Starting New Extraction Mission ---", to_console=False)
        
        thread = threading.Thread(target=self._extract_subtitles_logic, args=(files_to_process_paths,), daemon=True)
        thread.start()

        self.master.after(300000, self.show_patience_message)

    def show_patience_message(self):
        if self.ui.extract_button['text'] == "Cancel Extraction":
            messagebox.showinfo("Don't Panic!", "The extraction is taking a while. Don't panic, the Force is with you!")

    def _cancel_extraction(self):
        self.log_message("--- MISSION ABORT SIGNAL RECEIVED ---", to_console=True)
        self.ui.status_label.config(text="Cancelling mission... Please wait for the current target to finish.")
        self.cancel_requested.set()
        self.ui.extract_button.config(state=tk.DISABLED, text="Cancelling...")

    def _toggle_extraction_controls(self, is_extracting):
        if is_extracting:
            self.cancel_requested.clear()
            self.ui.extract_button.config(text="Cancel Extraction", command=self._cancel_extraction)
            for btn in [self.ui.remove_button, self.ui.select_folder_button, self.ui.select_langs_button, self.ui.ocr_settings_button, self.ui.remove_with_subs_button]:
                btn.config(state=tk.DISABLED)
        else:
            self.ui.extract_button.config(text="Extract Subtitles", command=self.start_extraction_thread, state=tk.NORMAL)
            for btn in [self.ui.remove_button, self.ui.select_folder_button, self.ui.select_langs_button, self.ui.ocr_settings_button, self.ui.remove_with_subs_button]:
                btn.config(state=tk.NORMAL)

    def _update_status_safe(self, message):
        self.master.after(0, lambda: self.ui.status_label.config(text=message))

    def _update_progress_safe(self, value):
        self.master.after(0, lambda: self.ui.progress_var.set(value))

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

        self._update_status_safe(summary_message or "Mission accomplished!"); self.ui.progress_var.set(0)
        messagebox.showinfo("Mission Complete", (summary_message or "Extraction run complete, Commander.") + "\n\nCheck Mission Debrief (Log) for details.", parent=self.master)
        self._toggle_extraction_controls(is_extracting=False)

        for item in self.ui.file_tree.get_children():
            file_path = self.ui.file_tree.item(item, "values")[0]
            if file_path in self.files_with_success:
                self.ui.file_tree.set(item, "Status", "Completed")
            elif file_path in self.files_with_errors or file_path in self.files_timed_out:
                self.ui.file_tree.set(item, "Status", "Failed")

    def _check_for_existing_subs(self, movie_file_path):
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
        
        # Build the command as a list of arguments
        command_parts = ocr_command_raw.replace("{INPUT_FILE_PATH}", temp_image_sub_path)
        command_parts = command_parts.replace("{OUTPUT_SRT_PATH}", temp_ocr_output_srt_path)
        command_parts = command_parts.replace("{LANG_3_CODE}", safe_lang_code_for_ocr)

        self.log_message(f"[OCR CMD] {command_parts}", to_console=True)
        ocr_success = False
        try:
            ocr_proc = subprocess.run(command_parts, shell=True, capture_output=True, text=True, encoding='utf-8', timeout=self.settings['ffmpeg_ocr_timeout'], creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0, check=False)
            if ocr_proc.stdout and ocr_proc.stdout.strip(): self.log_message(f"[OCR STDOUT]:\n{ocr_proc.stdout.strip()}")
            if ocr_proc.stderr and ocr_proc.stderr.strip(): self.log_message(f"[OCR STDERR]:\n{ocr_proc.stderr.strip()}")
            self.log_message(f"[OCR RETURN CODE]: {ocr_proc.returncode}")
            if ocr_proc.returncode == 0 and os.path.exists(temp_ocr_output_srt_path) and os.path.getsize(temp_ocr_output_srt_path) > 0:
                shutil.move(temp_ocr_output_srt_path, target_srt_path)
                self.log_message(f"[OCR SUCCESS] Translation complete: {os.path.basename(target_srt_path)}", to_console=True); ocr_success = True
            elif ocr_proc.returncode == 0: self.log_message(f"[OCR FAILED] Droid translation unit (RC 0) but output datapad (SRT) is empty/missing: {temp_ocr_output_srt_path}", to_console=True)
            else: self.log_message(f"[OCR FAILED] Droid translation unit malfunctioned (RC {ocr_proc.returncode}).", to_console=True)
        except subprocess.TimeoutExpired: self.log_message(f"[OCR TIMEOUT] Comlink lost with OCR droid for {temp_image_sub_basename} after {self.settings['ffmpeg_ocr_timeout']}s.", to_console=True)
        except FileNotFoundError: self.log_message(f"[OCR ERROR] OCR Droid (tool) not found. Check Holocron (Config) for: {command_parts}", to_console=True)
        except Exception as e:
            self.log_message(f"[OCR CRITICAL ERROR] Catastrophic droid failure during OCR: {e}", to_console=True); import traceback; self.log_message(traceback.format_exc(), to_console=True)
        finally:
            if os.path.isdir(ocr_session_temp_dir): shutil.rmtree(ocr_session_temp_dir, ignore_errors=True); self.log_message(f"[OCR Cleanup] Erased temporary droid memory banks: {ocr_session_temp_dir}", to_console=False)
        return ocr_success

    def _extract_subtitles_logic(self, files_to_process):
        total_files = len(files_to_process); overall_subs_extracted_count = 0; processed_for_progress_count = 0
        selected_gui_output_format = self.ui.output_format_var.get()
        self.log_message(f"Using output format: {selected_gui_output_format}", to_console=True)
        self.log_message(f"Language filter: {self._get_current_lang_filter_display()}", to_console=True)
        if self.settings.get('ocr_enabled') and self.settings.get('ocr_command_template'):
            self.log_message(f"[OCR STATUS] OCR Droid is ONLINE. Protocol: {self.settings['ocr_command_template'][:50]}...", to_console=True)
        else:
            self.log_message("[OCR STATUS] OCR Droid OFFLINE or no protocol. Image subs will be copied or skipped (if text output chosen).", to_console=True)

        for i, movie_file_path in enumerate(files_to_process):
            if self.cancel_requested.is_set():
                break

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
                    elif output_target_format_gui in TEXT_BASED_OUTPUT_FORMATS:
                        if input_codec in IMAGE_BASED_CODECS:
                            if self.settings.get('ocr_enabled') and self.settings.get('ocr_command_template'):
                                run_ocr = True; use_direct_ffmpeg_extract = False
                            else:
                                self.log_message(f"[INFO] Skipping image-based signal {stream_idx} ({input_codec}, lang {lang_code}) for {movie_filename}. Cannot convert to {output_target_format_gui.upper()} without OCR Droid. Use 'copy' or deploy OCR Droid via Holocron (Config).", to_console=True)
                                if not current_file_had_error_flag: self.files_with_errors.append(movie_filename); current_file_had_error_flag = True
                                continue
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
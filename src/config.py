
import os
import configparser

# --- Constants ---
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
    "progress_bg": "#0000FF", "progress_trough": "#D0D0D0"
}
DARK_THEME = {
    "bg": "#2D2D2D", "fg": "#FFFFFF", "list_bg": "#3C3C3C", "list_fg": "#FFFFFF",
    "button_bg": "#555555", "button_fg": "#FFFFFF", "accent_bg": "#0078D7",
    "accent_fg": "#FFFFFF", "log_bg": "#252525", "log_fg": "#FFFFFF",
    "progress_bg": "#FF0000", "progress_trough": "#4A4A4A"
}

# --- Constants for Subtitle Processing ---
IMAGE_BASED_CODECS = {'hdmv_pgs_subtitle', 'pgssub', 'dvd_subtitle', 'dvdsub', 'pgs'}
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

class AppConfig:
    def __init__(self, app_dir):
        self.app_dir = app_dir
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

    def save_config(self, extract_all_languages_flag, user_selected_languages):
        config_path = self.get_config_path()
        self.config.set('General', 'theme', self.settings['theme'])
        self.config.set('General', 'last_folder', self.settings['last_folder'])
        self.config.set('Paths', 'ffmpeg_path', self.settings['ffmpeg_path'])
        self.config.set('Paths', 'ffprobe_path', self.settings['ffprobe_path'])
        self.config.set('Timeouts', 'ffprobe_timeout', str(self.settings['ffprobe_timeout']))
        self.config.set('Timeouts', 'ffmpeg_extract_timeout', str(self.settings['ffmpeg_extract_timeout']))
        self.config.set('Timeouts', 'ffmpeg_ocr_timeout', str(self.settings.get('ffmpeg_ocr_timeout', DEFAULT_FFMPEG_OCR_TIMEOUT)))
        self.config.set('Extraction', 'default_output_format', self.settings['default_output_format'])
        lang_str_to_save = 'all' if extract_all_languages_flag or not user_selected_languages else ','.join(sorted(list(user_selected_languages)))
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
            print(f"Error writing configuration: {e}")

Sub-Extractor
=======================

I semi-vibecoded this application to solve a common problem for users of self-hosted media servers like Plex and Jellyfin: the need to transcode video files when displaying image-based subtitles (PGS, VOBSUB). By extracting these subtitles into a text-based format like SRT, media servers can direct stream the video, saving significant CPU resources.

Sub-Extractor is a user-friendly, cross-platform desktop application that automates this process. Powered by FFmpeg and your favorite OCR tool, it simplifies getting subtitles out of your media collection.

Features
--------

*   **Bulk Processing**: Select a folder and the app will automatically find all supported video files.
*   **Status Display**: Files are listed with a status indicating if they are "Ready to Extract", have "Subtitles Present", or if the extraction has "Completed" or "Failed".
*   **Multiple Output Formats**:
    *   Extract text-based subtitles (like SRT, ASS) directly into **SRT**, **ASS**, or **VTT** formats.
    *   **Copy** streams directly without re-encoding, preserving the original format (e.g., PGS/SUP, DVD/SUB).
*   **Image-Based Subtitle OCR**:
    *   Integrates with external command-line OCR tools (like [Subtitle Edit](https://www.google.com/url?sa=E&q=https%3A%2F%2Fwww.nikse.dk%2Fsubtitleedit), [VOBSUB2SRT](https://www.google.com/url?sa=E&q=https%3A%2F%2Fgithub.com%2Fruediger%2FVobSub2SRT), etc.) to convert image-based subtitles (PGS, VOBSUB) into text-based SRT files.
    *   Features a user-friendly **OCR Settings Dialog** to configure your tool without editing text files.
*   **Intelligent Filtering**:
    *   Filter extractions by one or more languages (e.g., eng, jpn, fre).
    *   Automatically skips files that already have corresponding subtitle files.
    *   Remove files from the list that already have subtitles with a single click.
*   **User-Friendly Interface**:
    *   Clean, modern UI with **Light and Dark themes**.
    *   Real-time progress bar and status updates. The progress bar is blue in light theme and red in dark theme.
    *   **Cancel** an ongoing extraction job at any time.
    *   A friendly "Don't Panic!" message will appear if an extraction takes longer than five minutes.
    *   Detailed logging for easy troubleshooting.
*   **Cross-Platform**: Built with Python and Tkinter, it runs on Windows, macOS, and Linux.

Prerequisites
-------------

*   **Python 3**: The application is written in Python. Ensure you have Python 3 installed.
*   **FFmpeg and FFprobe**: The core of the application relies on FFmpeg for probing and extracting streams.
    *   **Installation**: You must install FFmpeg and ensure that ffmpeg and ffprobe are accessible from your system's command line (i.e., they are in your system's PATH).
    *   **Verification**: Open a terminal or command prompt and type ffmpeg -version and ffprobe -version. If they run successfully, you're all set.

Installation & Setup
--------------------

*   **Download**: Place the application files in a dedicated folder.
*   **Run for the First Time**: Execute the script from your terminal:
    ```
    python src/main.py
    ```
    On its first run, the application will automatically create a configuration file (sub\_extractor\_settings.ini) and a logs directory in the same folder.
*   **Configuration**:
    *   The app will attempt to use ffmpeg and ffprobe from your system PATH by default.
    *   If your executables are located elsewhere, you can specify their full paths in the sub\_extractor\_settings.ini file or by using the **Edit Holocron (Config)** button in the app.

How to Use
----------

*   **Launch the App**: Run `python src/main.py` or run the executable file from the `dist` directory.
*   **Select a Folder**: Click the **Select Folder** button and choose the directory containing your video files. The app will scan the folder and all its subdirectories for media files and display them in the list with their subtitle status.
*   **Configure Options**:
    *   **Output Format**: Choose the desired subtitle format (srt, ass, vtt, or copy).
    *   **Filter Languages**: Click **Filter Languages...** to scan the files for available subtitle languages and select which ones you want to extract.
    *   **OCR Settings**: If you need to convert image-based subtitles, click **OCR Settings...** to enable and configure your OCR tool (see section below).
    *   **Skip if exists**: Check this box to avoid re-extracting subtitles for files that already have an associated subtitle file in the same directory.
    *   **Remove Selected**: Select one or more files from the list and click this to remove them from the current batch.
    *   **Remove with Subtitles**: Click this to remove all files from the list that already have subtitles.
*   **Start Extraction**: Click the **Extract Subtitles** button to begin the process.
    *   The progress bar will show the overall progress.
    *   The status label will provide updates on the current file being processed.
    *   You can click **Cancel Extraction** at any time to safely abort the mission.
*   **Review Results**:
    *   Extracted subtitle files will be saved in the same directory as their source video file. The naming convention is \[VideoFileName\].\[LanguageCode\].\[StreamIndex\].\[extension\].
    *   A final summary will appear when the job is complete or cancelled.
    *   Click **View Log** for a detailed mission debrief, including successes, errors, and skipped files.

Configuring OCR for Image-Based Subtitles
-----------------------------------------

This is the most powerful feature for converting subtitles like PGS (from Blu-rays) and VOBSUB (from DVDs) into text.

### 1\. Install an OCR Tool

You need a command-line tool that can convert image subtitles to SRT. A highly recommended tool is the **[Subtitle Edit](https://www.google.com/url?sa=E&q=https%3A%2F%2Fwww.nikse.dk%2Fsubtitleedit)** command-line executable (SubtitleEdit.exe).

### 2\. Open OCR Settings

In the app, click the **OCR Settings...** button.

### 3\. Configure the Dialog

*   **Enable OCR Droid**: Check this box to activate the OCR functionality.
*   **OCR Droid Protocol (Command Template)**: This is the most important part. You must provide the command that the app will use to run your OCR tool. Use the following placeholders which the app will replace for each subtitle stream:
    *   {INPUT\_FILE\_PATH}: The full path to the temporary image subtitle file (e.g., a .sup or .sub file).
    *   {OUTPUT\_SRT\_PATH}: The full path where the final .srt file should be saved.
    *   {LANG\_3\_CODE}: The 3-letter language code of the stream (e.g., eng, fre). Note: Your OCR tool must support a language switch for this to be effective.
*   **Default Language**: The 3-letter code to use if a subtitle stream has no language metadata.

### 4\. Example Command Template (for Subtitle Edit)

```
"C:\Path\To\Your\SubtitleEdit.exe" /convert {INPUT_FILE_PATH} subrip /outputfolder:"{OUTPUT_SRT_PATH}" /ocrengine:Tesseract /ocrlanguage:{LANG_3_CODE}
```

Use the Browse... button to easily find your SubtitleEdit.exe!

> **Note on Subtitle Edit & Tesseract**: When using Tesseract with Subtitle Edit's CLI, the language is primarily controlled by the Tesseract language data you have installed. The /ocrlanguage: switch may not always override this. Ensure you have the necessary language packs for Tesseract installed.

### 5\. Save

Click **Save Protocol**. The app is now ready to perform OCR.

When you select an output format like srt and the app encounters an image-based stream (e.g., hdmv\_pgs\_subtitle), it will automatically:

*   Extract the image stream to a temporary file.
*   Run your configured OCR command on that file.
*   Save the resulting SRT file.
*   Clean up all temporary files.

Building from Source
--------------------

To build the executable from the source code, you will need to have `pyinstaller` installed. You can install it using pip:

```
pip install -r requirements.txt
```

Once you have `pyinstaller` installed, you can build the executable by running the following command from the root directory of the project:

```
pyinstaller --onefile --windowed --icon=NONE src/main.py
```

This will create a single executable file in the `dist` directory.

Troubleshooting
---------------

*   **FFmpeg/FFprobe not found**: Make sure the executables are in your system's PATH or their full paths are correctly specified in sub\_extractor\_settings.ini.
*   **OCR Fails**:
    *   Check the log (View Log) for the exact command that was run and any error messages from the OCR tool.
    *   Try running the command manually in your terminal to debug it.
    *   Ensure your command template correctly quotes the file path placeholders, especially if they might contain spaces. The app now attempts to quote these, but it's good practice.
*   **"File ended prematurely" warning**: This is a common, often non-critical warning from FFmpeg when working with subtitle streams. The application log will note this, but in most cases, the extraction will still be successful.

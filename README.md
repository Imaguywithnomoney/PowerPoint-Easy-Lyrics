# PowerPoint Easy Lyrics

**Lyrics to PowerPoint in a Second**

PowerPoint Easy Lyrics is a desktop application designed to help people who have difficulty creating PowerPoint lyric presentations. It is especially useful for churches, worship teams, events, and other occasions where lyrics need to be displayed cleanly on a screen.

---

## What the App Does

This app allows you to create PowerPoint slides according to your preferences with seamless transitions:
- **Same-Slide Stacked Lyrics (Recommended)**: Place all or multiple stanzas on a single slide stacked on top of each other with click animations (previous stanza disappears and the next appears). This allows background videos and motion loops to play continuously without cuts or interruptions!
- **One Stanza per Slide**: You can also easily split lyrics across individual slides.

---

## Project Structure (Files & Folders Explained)

| File / Folder | Purpose |
| :--- | :--- |
| **`app/`** | The main application source code directory. |
| **`app/main.py`** | **The main file you run to start the application.** Launches the graphical interface. |
| **`app/ppt/`** | The PowerPoint generation pipeline and layout engine. |
| ↳ `app/ppt/layout_engine.py` | Sets up widescreen (16:9) slides, black backgrounds, and centered typography. |
| ↳ `app/ppt/animation_engine.py` | Handles PPTX XML animation timings for stacked stanzas. |
| ↳ `app/ppt/com_backend.py` | Connects directly with Microsoft PowerPoint on Windows to create native click animations. |
| ↳ `app/ppt/exporter.py` | Orchestrates lyric segmentation, slide grouping, and file saving. |
| **`app/ai/`** | Lyrics text segmentation parser. |
| ↳ `app/ai/segmenter.py` | Splits raw lyric text into stanzas and lines deterministically. |
| ↳ `app/ai/schema.py` | Data structures for lyric stanzas and sections. |
| **`dist/`** | Contains the compiled standalone Windows executable (`PowerPointEasyLyrics.exe`). |
| **`build/`** | Contains PyInstaller build specifications (`pyinstaller.spec`). |
| **`installer/`** | Contains Inno Setup configuration (`PowerPointEasyLyrics.iss`) to compile a Windows installer. |
| **`tests/`** | Contains automated testing scripts (`manual_smoke_test.py`). |
| **`requirements.txt`** | Lists the required Python packages needed to run the app from source. |
| **`.gitignore`** | Tells Git which temporary files (caches, build artifacts) to exclude from GitHub. |

---

## Simple Step-by-Step Guide: How to Install and Run

### Prerequisites
1. **Windows 10 or 11** with **Microsoft PowerPoint** installed.
2. **Python 3.11+** installed (make sure to check **"Add python.exe to PATH"** during installation).

### Step 1: Open Terminal or Command Prompt
Open Command Prompt, PowerShell, or Windows Terminal, and navigate to the project folder:
```bash
cd "path\to\PowerPoint Easy Lyrics"
```

### Step 2: Install Required Dependencies
Run this command to install the required libraries:
```bash
pip install -r requirements.txt
```
> **Note on Dependencies:** The dependencies in `requirements.txt` (`python-pptx`, `lxml`, `pywin32`, `pyinstaller`) are completely safe, legitimate, open-source libraries standard in Python development.

### Step 3: Launch the Application
Run the main file:
```bash
python -m app.main
```
*(Or simply double-click `dist/PowerPointEasyLyrics.exe` if you want to run the standalone app without using the command line).*

---

## How to Use the App

1. **Paste Song Lyrics**: Paste your raw lyrics into the large text area.
2. **Configure Settings**:
   - **Max Stanza per Slide** *(Default: 30)*: Controls how many stanzas will appear on one slide stacked on top of each other with animations. Setting this high (e.g. 30) puts the whole song on one slide for seamless looping backgrounds!
   - **Lines per Stanza** *(Default: 1)*: Controls how many lines appear in each text box.
   - **Effect**: Choose your transition animation (`Appear/Disappear`, `Fade/Fade`, or `Wipe/Fade`).
3. **Choose Output Location**: Enter a file name or click **Browse** to choose where to save the presentation.
4. **Click "Generate PPTX"**: Your presentation will be created in seconds!

> **Important Note:** If a PowerPoint window briefly pops up during generation, this is completely normal! The app uses PowerPoint to build the native presentation animations.

---

## Building the Executable (.exe)

If you make modifications to the code and want to recompile the standalone `.exe`:
```bash
pyinstaller build/pyinstaller.spec
```
The newly compiled file will be located at `dist/PowerPointEasyLyrics.exe`.

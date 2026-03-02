# NSBM-dl

A modern, fast, and fully automated utility to scrape, organize, and download all your course materials from the NSBM Moodle portal directly to your computer.

Built with Python and a fluid PyQt6 interface.

![Image](https://github.com/user-attachments/assets/7db4ee31-8325-41e5-9318-1f957bed8729)

---

## Features
- **One-Click Downloads**: Logs into your Moodle account and pulls your enrolled courses instantly.
- **Smart Organization**: Automatically sorts your downloads into categorized folders (e.g., `Downloads/NSBM-Moodle/Y2S1`).
- **Modern UI**: Clean, frameless interface with Dark and Light themes.
- **Selective Syncing**: Check or uncheck specific modules or overarching semesters using the collapsible tree widget.

---

## Building from Source

If you prefer not to use a pre-compiled `.exe` or want to modify the code yourself, you can build the executable from source.

### 1. Prerequisites 
Make sure you have **Python 3.11+** installed on your system.

Clone the repository and install the dependencies:
```bash
git clone https://github.com/pamindu-fernando/NSBM-dl.git
cd NSBM-dl
pip install -r requirements.txt
```

*(Note: Ensure your `requirements.txt` includes `requests`, `beautifulsoup4`, `lxml`, and `PyQt6`)*

### 2. Install PyInstaller
To package the app into a standalone executable, you'll need PyInstaller:
```bash
pip install pyinstaller
```

### 3. Run the Build Script
Because PyQt6 is massive and includes dozens of unused rendering libraries (like WebEngines and 3D Rendering), a standard PyInstaller build will result in a bloated file >150MB.

We have included a customized build script (`build_small.ps1`) that aggressively strips away unused PyQt6 modules, drastically reducing the file size.

Open a PowerShell window in the project directory and run:

```powershell
.\build_small.ps1
```

### 4. Locate Your Executable
Once the script finishes, navigating to the newly created `dist/NSBM-dl` directory will reveal your fully bundled `moodle_downloader.exe`. You can move this folder anywhere or create a shortcut to it on your Desktop.

---

## Troubleshooting

- **`ModuleNotFoundError: No module named 'email'`**: 
  Our source code explicitly imports `email` and `pkg_resources` to force PyInstaller's strict dependency graph to catch them. If you alter the imports, ensure these remain at the top of the file before compiling.
- **Missing Taskbar Icon on Windows**: 
  The codebase uses `ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID()` to force Windows to recognize the PyQt window icon. If you rename the internal AppID, the taskbar icon may revert to the default Python logo.

---

*Made with ❤️ by [Pamindu Fernando](https://github.com/pamindu-fernando)*

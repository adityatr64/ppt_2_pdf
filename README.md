# PPT to PDF Converter and Merger

A desktop application for converting multiple PowerPoint presentations to PDF and merging them into a single document.

## Requirements

- Windows, Linux (untested), or macOS
- Python 3.8+

### Conversion Backend (auto-detected)

Install at least one supported app for your OS:

- Windows: Microsoft PowerPoint, WPS Office, LibreOffice, or ONLYOFFICE
- Linux: LibreOffice or ONLYOFFICE
- macOS: Keynote, LibreOffice, or ONLYOFFICE

### Python Dependencies

```
pypdf
comtypes
pillow
customtkinter
```

Notes:
- `comtypes` is used only for Windows COM automation (PowerPoint/WPS backends).
- Other backends use command-line automation of installed apps.

## Installation

1. Clone or download this repository
2. Install dependencies:

bash/pwsh :
```sh
git clone https://github.com/adityatr64/ppt_2_pdf.git
cd ppt_2_pdf
```
Windows :
```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt 
```
Macos :
```
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt 

```
Linux people figure it out y'all are smart ;]

## Usage

Run the application from src:
```sh
python main.py
```

### Steps

1. Click **Add Files** to select PowerPoint files (.pptx or .ppt)
2. Reorder files as needed using:
   - Drag and drop within the list
   - **Move Up** / **Move Down** buttons
   - **Sort** to sort by name ascending
3. Organize work into tabs (each tab is its own conversion task)
   - Use **+ New Task** to create another task tab
   - Click a tab to switch between tasks
3. Configure options:
   - Open PDF after conversion
4. Start conversion for the current tab:
   - **Convert & Merge to PDF**: make one merged PDF from this tab
   - **Make Separate PDFs**: create one PDF per PowerPoint in this tab
5. Start other tabs in parallel (each tab runs independently hopefully without issues)
6. Use **Cancel** to stop conversion for the current tab only, Residuals may exist delete them youself please 🥀 

## Project Structure

```
ppt_2_pdf/
├── main.py                          # Application entry point
├── requirements.txt
├── assets/
│   ├── image.ico
│   └── fonts/
├── controllers/
│   ├── __init__.py
│   ├── app_controller.py            # Main app event orchestration
│   ├── system_ops.py                # OS-level helpers
│   └── task_manager.py              # Per-tab task lifecycle
├── services/
│   ├── __init__.py
│   ├── backend_support.py           # Backend detection/support checks
│   ├── conversion_runtime.py        # Runtime conversion flow
│   ├── conversion_types.py          # Shared conversion data types
│   └── converter_service.py         # PPT -> PDF conversion implementation
└── views/
   ├── __init__.py
   ├── main_view.py                 # UI layout and widgets
   ├── styles.py                    # UI styling
   └── icons.py                     # App/window icon helpers
```

## Architecture

The application follows the **Service-View-Controller (SVC)** pattern:

- **Services**: Business logic for PPT to PDF conversion and PDF merging
- **Views**: UI components, styling, and user interaction display
- **Controllers**: Coordinates between services and views, handles events

## Features

- Select multiple PowerPoint files at once
- Organize files into multiple task tabs
- Reorder files via drag-and-drop or buttons
- Independent per-tab conversion status and progress
- Parallel conversions across tabs (multi-threaded i think T-T)
- Automatic cleanup of temporary files
- Option to open the resulting PDF automatically

## Notes

- On startup, the app auto-detects available backends and picks the best one for your platform.
- If no backend is detected, the app shows an install prompt with supported options (Hopefully not tested on any platform).
- Large presentations may take longer to convert,BE **PATIENT**.
- The application runs conversions in a background thread to keep the UI responsive.

## Building Executable

To package the application as a standalone .exe file:

1. Install PyInstaller:

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install pyinstaller
```

2. Build the executable:

```powershell
pyinstaller --clean --noconfirm --onedir --windowed --name "PPT 2 PDF" --icon "./assets/image.ico" --add-data "assets;assets" main.py
```

3. The executable will be created in the `dist/` folder.

### Build Options

| Option | Description |
|--------|-------------|
| `--onedir` | Bundle everything into a packaged application folder |
| `--windowed` | Hide the console window |
| `--name "PPT 2 PDF"` | Set the output filename |
| `--icon ./assets/image.ico` | Set the executable icon |
| `--add-data "assets;assets"` | Bundle the `assets/` folder for runtime icon/fonts |

**Note:** The resulting .exe/.app still requires at least one supported conversion app to be installed on the target machine.

## Troubleshooting

**No backend found**: Install one supported app for your platform (PowerPoint/WPS/LibreOffice/ONLYOFFICE/Keynote).

**Conversion fails**: Make sure the PowerPoint files are not corrupted and can be opened manually.

**Permission errors**: Run the application with appropriate permissions or choose a different output directory.

**Any other errors**: Open issue on github and hope i or someone else smart sees it,or fix the issue and open PR yourself ;)

## License

GPL-2 License

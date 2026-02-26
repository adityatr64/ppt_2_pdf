# PPT to PDF Converter and Merger

A desktop application for converting multiple PowerPoint presentations to PDF and merging them into a single document.

## Requirements

- Windows, Linux, or macOS
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
3. Configure options:
   - Open PDF after conversion
4. Click **Convert and Merge to PDF**
5. Choose the output location and filename

## Project Structure

```
ppt2pdf/
├── main.py                      # Application entry point
├── controllers/
│   ├── __init__.py
│   └── app_controller.py        # Application logic and event handling
├── views/
│   ├── __init__.py
│   ├── main_view.py             # UI components and layout
│   ├── styles.py                # TTK styling configuration
│   └── icons.py                 # Icon generation utilities
└── services/
    ├── __init__.py
    └── converter_service.py     # PPT to PDF conversion and merging
```

## Architecture

The application follows the **Service-View-Controller (SVC)** pattern:

- **Services**: Business logic for PPT to PDF conversion and PDF merging
- **Views**: UI components, styling, and user interaction display
- **Controllers**: Coordinates between services and views, handles events

## Features

- Select multiple PowerPoint files at once
- Reorder files via drag-and-drop or buttons
- Progress indicator during conversion
- Automatic cleanup of temporary files
- Option to open the resulting PDF automatically

## Notes

- On startup, the app auto-detects available backends and picks the best one for your platform.
- If no backend is detected, the app shows an install prompt with supported options (Hopefully).
- Large presentations may take longer to convert.
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
pyinstaller --onedir --windowed --name "PPT 2 PDF" main.py --icon=assets/image.ico
```

3. The executable will be created in the `dist/` folder.

### Build Options

| Option | Description |
|--------|-------------|
| `--onedir` | Bundle everything into a packaged application folder |
| `--windowed` | Hide the console window |
| `--name "PPT2PDF"` | Set the output filename |
| `--icon=icon.ico` | Use a custom icon (optional) |

**Note:** The resulting .exe/.app still requires at least one supported conversion app to be installed on the target machine.

## Troubleshooting

**No backend found**: Install one supported app for your platform (PowerPoint/WPS/LibreOffice/ONLYOFFICE/Keynote).

**Conversion fails**: Make sure the PowerPoint files are not corrupted and can be opened manually.

**Permission errors**: Run the application with appropriate permissions or choose a different output directory.

## License

GPL-2 License

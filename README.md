# PowerPoint Comparison Tool

This tool converts PowerPoint slides to PNG images for comparison.

## Prerequisites

### macOS
1. Install LibreOffice:
   ```bash
   brew install --cask libreoffice
   ```

2. Install Poppler (for PDF to image conversion):
   ```bash
   brew install poppler
   ```

### Linux
```bash
sudo apt-get install libreoffice poppler-utils
```

### Windows
1. Download and install LibreOffice from https://www.libreoffice.org/
2. Download Poppler from https://github.com/oschwartz10612/poppler-windows/releases/
3. Add Poppler's bin directory to your PATH

## Installation

1. Create and activate a virtual environment:
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

2. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

```bash
python ppt_compare.py <file1.pptx> <file2.pptx>
```

Example:
```bash
python ppt_compare.py presentation1.pptx presentation2.pptx
```

## Output

The tool will:
1. Create a temporary directory (e.g., `/tmp/ppt_compare_XXXXXX`)
2. Convert each PowerPoint file's slides to PNG images
3. Save images in separate subdirectories for each file
4. Display the paths to the temporary directories

**Note:** The temporary directories are NOT automatically deleted, allowing you to review the generated images.

## Output Structure

```
/tmp/ppt_compare_XXXXXX/
├── presentation1/
│   ├── slide_001.png
│   ├── slide_002.png
│   └── slide_003.png
└── presentation2/
    ├── slide_001.png
    ├── slide_002.png
    └── slide_003.png
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
python ppt_compare.py <file1.pptx> <file2.pptx> [output_dir]
```

### Arguments
- `file1.pptx` - First PowerPoint file (source)
- `file2.pptx` - Second PowerPoint file (target)
- `output_dir` - Optional output directory for results
  - If specified: Files are saved to this directory and preserved
  - If not specified: Uses temporary directory, displays PDF, then cleans up after user confirmation

### Examples

**With temporary directory (auto-cleanup):**
```bash
python ppt_compare.py presentation1.pptx presentation2.pptx
```
The PDF will open automatically. Press Enter when done viewing to clean up temporary files.

**With specified output directory (files preserved):**
```bash
python ppt_compare.py presentation1.pptx presentation2.pptx ./output
```
All files including the PDF will be saved to `./output` and preserved.

## Output

The tool will:
1. Create an output directory (temporary or specified)
2. Convert each PowerPoint file's slides to PNG images (150 DPI)
3. Generate SHA-256 hash files for each image
4. Create a slide comparison mapping
5. Generate a side-by-side comparison PDF
6. Automatically open the PDF for viewing

### PDF Features
- Landscape orientation for side-by-side viewing
- Each page shows one comparison:
  - Matched slides: Source on left, target on right
  - Source-only slides: Source on left, blank on right
  - Target-only slides: Blank on left, target on right
- Clear labels and titles for each comparison
- Automatic scaling to fit slides on page

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
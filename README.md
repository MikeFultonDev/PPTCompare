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

### Option 1: Using the Wrapper Script (macOS/Linux only)

The easiest way to run the tool is using the provided wrapper script that automatically activates the virtual environment:

```bash
./ppt_compare.sh <file1.pptx> <file2.pptx> [output_dir] [options]
```

**Tip:** You can create a symbolic link to the script in a directory in your PATH (e.g., `~/bin`) to run it from anywhere:
```bash
ln -s /path/to/PPTCompare/ppt_compare.sh ~/bin/ppt_compare.sh
# Then run from anywhere:
ppt_compare.sh file1.pptx file2.pptx
```

**Note:** This script only works on macOS and Linux. Windows users should use Option 2 below.

### Option 2: Manual Activation (All Platforms)

**Always activate the virtual environment first:**
```bash
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

Then run the comparison tool:
```bash
python ppt_compare.py <file1.pptx> <file2.pptx> [output_dir] [options]
```

### Arguments
- `file1.pptx` - First PowerPoint file (source), or the only file when using --gitdiff or --gitpr
- `file2.pptx` - Second PowerPoint file (target), not used with --gitdiff or --gitpr
- `output_dir` - Optional output directory for results
  - If specified: Files are saved to this directory and preserved
  - If not specified: Uses temporary directory, displays PDF, then cleans up after viewer closes

### Options
- `--debug` - Enable debug output showing detailed processing information
- `--perf` - Show performance timing for different stages of processing
- `--gitdiff` - Compare current file with last committed version (only file1 is used)
- `--gitpr NUM` or `--gitpr=NUM` - Compare file from main branch with PR version (specify PR number)
- `--suppress-common-slides` - Suppress slides present in both presentations (default)
- `--no-suppress-common-slides` - Show all slides including common ones
- `--show-moved-pages` - Show slides in original order with arrows for repositioned slides (default)
- `--no-show-moved-pages` - Show slides grouped by match status without arrows
- `-h, --help` - Show help message with all options

### Examples

#### Using the Wrapper Script (macOS/Linux)

**Basic comparison with temporary directory (auto-cleanup):**
```bash
./ppt_compare.sh presentation1.pptx presentation2.pptx
```
The PDF will open automatically. Close the PDF window when done to clean up temporary files.

**With performance timing:**
```bash
./ppt_compare.sh file1.pptx file2.pptx --perf
```
Shows detailed timing breakdown of parallel processing stages.

**Compare with git (current vs last committed):**
```bash
./ppt_compare.sh presentation.pptx --gitdiff
```

**Compare file from main branch with PR version:**
```bash
./ppt_compare.sh presentation.pptx --gitpr 123
# or
./ppt_compare.sh presentation.pptx --gitpr=123
```
This fetches the file from both the main branch and the specified PR, then compares them.

**With specified output directory (files preserved):**
```bash
./ppt_compare.sh presentation1.pptx presentation2.pptx ./output
```
All files including the PDF will be saved to `./output` and preserved.

**Show all slides including common ones:**
```bash
./ppt_compare.sh file1.pptx file2.pptx --no-suppress-common-slides
```

**With debug output:**
```bash
./ppt_compare.sh file1.pptx file2.pptx --debug
```

#### Manual Method (All Platforms)

For Windows users or if you prefer manual activation:

**Basic comparison:**
```bash
source venv/bin/activate  # On Windows: venv\Scripts\activate
python ppt_compare.py presentation1.pptx presentation2.pptx
```

**Compare with PR:**
```bash
source venv/bin/activate  # On Windows: venv\Scripts\activate
python ppt_compare.py presentation.pptx --gitpr 123
```

## Output

The tool will:
1. Create an output directory (temporary or specified)
2. Convert each PowerPoint file's slides to PNG images (100 DPI)
   - Uses parallel processing with separate LibreOffice instances for optimal performance
3. Generate SHA-256 hash files for each image
4. Create a slide comparison mapping
5. Generate a side-by-side comparison PDF
6. Automatically open the PDF for viewing
7. Wait for you to close the PDF viewer before cleaning up (if using temp directory)

### PDF Features
- Landscape orientation for side-by-side viewing
- Each page shows one comparison:
  - Matched slides: Source on left, target on right
  - Source-only slides: Source on left, blank on right
  - Target-only slides: Blank on left, target on right
- Color-coded bars:
  - Light grey: Slide present in both presentations (matched)
  - Red: Slide only in source presentation
  - Green: Slide only in target presentation
- Clear labels and titles for each comparison
- Automatic scaling to fit slides on page
- Optional arrows showing repositioned slides (--show-moved-pages)

## Performance

The tool uses parallel processing for optimal performance:
- **PPTX→PDF conversion**: Runs in parallel using separate LibreOffice instances
- **PDF→PNG conversion**: Runs in parallel for both files
- **Typical performance**: ~9 seconds for two 24-28 slide presentations (50% faster than sequential)

Use `--perf` flag to see detailed timing breakdown.

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
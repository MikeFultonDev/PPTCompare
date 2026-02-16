#!/usr/bin/env python3
"""
PowerPoint Comparison Tool
Converts PowerPoint slides to PNG images for comparison
"""

import os
import sys
import tempfile
import hashlib
import shutil
import subprocess
import argparse
from pathlib import Path
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor

try:
    from pptx import Presentation
    PYTHON_PPTX_AVAILABLE = True
except ImportError:
    PYTHON_PPTX_AVAILABLE = False

try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False


def compute_sha256(file_path):
    """Compute SHA-256 hash of a file"""
    sha256_hash = hashlib.sha256()
    with open(file_path, "rb") as f:
        # Read file in chunks to handle large files efficiently
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()


def convert_ppt_to_images_libreoffice(ppt_path, output_dir):
    """Convert PowerPoint to images using LibreOffice (cross-platform)"""
    import subprocess
    
    print(f"  Converting {Path(ppt_path).name} to PDF...")
    
    # First convert to PDF
    ppt_name = Path(ppt_path).stem
    
    # Try different LibreOffice command names
    libreoffice_commands = ['libreoffice', 'soffice']
    
    pdf_created = False
    for cmd in libreoffice_commands:
        try:
            result = subprocess.run(
                [cmd, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, str(Path(ppt_path).absolute())],
                capture_output=True,
                text=True,
                timeout=60
            )
            
            if result.returncode == 0:
                pdf_created = True
                break
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
    
    if not pdf_created:
        raise RuntimeError("LibreOffice not found. Please install LibreOffice.")
    
    # Find the generated PDF
    pdf_files = list(Path(output_dir).glob("*.pdf"))
    if not pdf_files:
        raise RuntimeError("PDF conversion failed")
    
    temp_pdf = str(pdf_files[0])
    print(f"  PDF created: {temp_pdf}")
    
    # Convert PDF to images
    if not PDF2IMAGE_AVAILABLE:
        raise RuntimeError("pdf2image not installed. Run: pip install pdf2image poppler-utils")
    
    print(f"  Converting PDF to PNG images...")
    images = convert_from_path(temp_pdf, dpi=150)
    
    for i, image in enumerate(images, start=1):
        output_path = os.path.join(output_dir, f"slide_{i:03d}.png")
        image.save(output_path, "PNG")
        
        # Compute SHA-256 hash
        sha256_hash = compute_sha256(output_path)
        hash_file = os.path.join(output_dir, f"slide_{i:03d}.sha256")
        with open(hash_file, 'w') as f:
            f.write(f"{sha256_hash}  {os.path.basename(output_path)}\n")
        
        print(f"    Slide {i} -> {output_path}")
        print(f"             SHA-256: {sha256_hash}")
    
    # Clean up temporary PDF
    if os.path.exists(temp_pdf):
        os.remove(temp_pdf)
    
    return len(images)


def load_slide_hashes(output_dir):
    """Load all slide hashes from a directory into a dictionary"""
    hashes = {}
    hash_files = sorted(Path(output_dir).glob("slide_*.sha256"))
    
    for hash_file in hash_files:
        with open(hash_file, 'r') as f:
            line = f.read().strip()
            if line:
                hash_value, filename = line.split(None, 1)
                # Extract slide number from filename (e.g., slide_001.png -> 1)
                slide_num = int(filename.split('_')[1].split('.')[0])
                hashes[slide_num] = hash_value
    
    return hashes


def compare_slides(dir1, dir2):
    """Compare slides between two directories and create a mapping"""
    print("\n" + "="*60)
    print("SLIDE COMPARISON")
    print("="*60)
    
    hashes1 = load_slide_hashes(dir1)
    hashes2 = load_slide_hashes(dir2)
    
    # Create reverse mapping for dir2 (hash -> slide number)
    hash_to_slide2 = {hash_val: slide_num for slide_num, hash_val in hashes2.items()}
    
    # Create reverse mapping for dir1 (hash -> slide number)
    hash_to_slide1 = {hash_val: slide_num for slide_num, hash_val in hashes1.items()}
    
    # Track which slides in dir2 have been matched
    matched_slides2 = set()
    
    # Store comparison results for PDF generation
    comparisons = []
    
    # Compare slides from dir1
    for slide1 in sorted(hashes1.keys()):
        hash1 = hashes1[slide1]
        if hash1 in hash_to_slide2:
            slide2 = hash_to_slide2[hash1]
            matched_slides2.add(slide2)
            print(f"slide {slide1} -> slide {slide2}")
            comparisons.append(('matched', slide1, slide2))
        else:
            print(f"slide {slide1} only in source")
            comparisons.append(('source_only', slide1, None))
    
    # Find slides only in dir2
    for slide2 in sorted(hashes2.keys()):
        if slide2 not in matched_slides2:
            print(f"slide {slide2} only in target")
            comparisons.append(('target_only', None, slide2))
    
    print("="*60)
    
    return comparisons, hashes1, hashes2


def generate_comparison_pdf(dir1, dir2, output_path, comparisons, suppress_common=True):
    """Generate a PDF with side-by-side slide comparisons"""
    print("\n" + "="*60)
    print("GENERATING COMPARISON PDF")
    print("="*60)
    
    if suppress_common:
        print("Suppressing common slides (matched slides will be excluded)")
    else:
        print("Including all slides (matched slides will be shown)")
    
    # Use landscape letter size for side-by-side comparison
    page_width, page_height = landscape(letter)
    
    # Create PDF
    c = canvas.Canvas(output_path, pagesize=landscape(letter))
    
    # Calculate dimensions for side-by-side layout
    margin = 36  # 0.5 inch margins
    bar_width = 10  # Width of color bar
    available_width = (page_width - 3 * margin - 2 * bar_width) / 2  # Space for two images with bars
    available_height = page_height - 2 * margin
    
    pages_added = 0
    
    for comparison_type, slide1, slide2 in comparisons:
        # Skip matched slides if suppress_common is True
        if suppress_common and comparison_type == 'matched':
            continue
        # Determine which images to display
        left_image = None
        right_image = None
        
        if comparison_type == 'matched':
            left_image = os.path.join(dir1, f"slide_{slide1:03d}.png")
            right_image = os.path.join(dir2, f"slide_{slide2:03d}.png")
            title = f"Source Slide {slide1} = Target Slide {slide2}"
        elif comparison_type == 'source_only':
            left_image = os.path.join(dir1, f"slide_{slide1:03d}.png")
            right_image = None
            title = f"Source Slide {slide1} (not in target)"
        elif comparison_type == 'target_only':
            left_image = None
            right_image = os.path.join(dir2, f"slide_{slide2:03d}.png")
            title = f"Target Slide {slide2} (not in source)"
        
        # Determine color bar colors based on comparison type
        if comparison_type == 'matched':
            left_bar_color = HexColor('#D3D3D3')  # Light grey
            right_bar_color = HexColor('#D3D3D3')  # Light grey
        elif comparison_type == 'source_only':
            left_bar_color = HexColor('#FF0000')  # Red
            right_bar_color = None  # No bar for blank side
        elif comparison_type == 'target_only':
            left_bar_color = None  # No bar for blank side
            right_bar_color = HexColor('#00FF00')  # Green
        
        # Add title
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(page_width / 2, page_height - 20, title)
        
        # Draw left color bar and image
        if left_image and os.path.exists(left_image):
            # Draw color bar
            if left_bar_color:
                c.setFillColor(left_bar_color)
                c.rect(margin, margin, bar_width, available_height, fill=1, stroke=0)
            
            img = ImageReader(left_image)
            img_width, img_height = img.getSize()
            
            # Calculate scaling to fit in available space
            scale = min(available_width / img_width, available_height / img_height)
            scaled_width = img_width * scale
            scaled_height = img_height * scale
            
            # Center the image in the left half (after the bar)
            x = margin + bar_width + (available_width - scaled_width) / 2
            y = margin + (available_height - scaled_height) / 2
            
            c.drawImage(left_image, x, y, width=scaled_width, height=scaled_height)
            
            # Add label
            c.setFillColorRGB(0, 0, 0)  # Reset to black
            c.setFont("Helvetica", 10)
            c.drawString(margin + bar_width, margin - 15, f"Source: slide_{slide1:03d}.png")
        
        # Draw right color bar and image
        if right_image and os.path.exists(right_image):
            # Draw color bar
            if right_bar_color:
                c.setFillColor(right_bar_color)
                right_bar_x = page_width / 2 + margin
                c.rect(right_bar_x, margin, bar_width, available_height, fill=1, stroke=0)
            
            img = ImageReader(right_image)
            img_width, img_height = img.getSize()
            
            # Calculate scaling to fit in available space
            scale = min(available_width / img_width, available_height / img_height)
            scaled_width = img_width * scale
            scaled_height = img_height * scale
            
            # Center the image in the right half (after the bar)
            x = page_width / 2 + margin + bar_width + (available_width - scaled_width) / 2
            y = margin + (available_height - scaled_height) / 2
            
            c.drawImage(right_image, x, y, width=scaled_width, height=scaled_height)
            
            # Add label
            c.setFillColorRGB(0, 0, 0)  # Reset to black
            c.setFont("Helvetica", 10)
            c.drawString(page_width / 2 + margin + bar_width, margin - 15, f"Target: slide_{slide2:03d}.png")
        
        # Draw center divider line
        c.setStrokeColorRGB(0.7, 0.7, 0.7)
        c.setLineWidth(1)
        c.line(page_width / 2, margin, page_width / 2, page_height - margin)
        
        c.showPage()
        pages_added += 1
    
    c.save()
    print(f"PDF saved to: {output_path}")
    print(f"Total pages: {pages_added}")
    print("="*60)


def process_powerpoint(ppt_path, base_temp_dir):
    """Process a single PowerPoint file and convert slides to PNG images"""
    
    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"PowerPoint file not found: {ppt_path}")
    
    # Create output directory for this file
    ppt_name = Path(ppt_path).stem
    output_dir = os.path.join(base_temp_dir, ppt_name)
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"\nProcessing: {ppt_path}")
    print(f"Output directory: {output_dir}")
    
    try:
        slide_count = convert_ppt_to_images_libreoffice(ppt_path, output_dir)
        print(f"  Successfully converted {slide_count} slides")
        return output_dir
    except Exception as e:
        print(f"  Error: {e}")
        raise


def main():
    """Main function to compare two PowerPoint files"""
    
    parser = argparse.ArgumentParser(
        description='Compare two PowerPoint presentations and generate a side-by-side comparison PDF',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  # Use temporary directory with common slides suppressed (default)
  python ppt_compare.py file1.pptx file2.pptx
  
  # Show all slides including common ones
  python ppt_compare.py file1.pptx file2.pptx --no-suppress-common-slides
  
  # Save to specific directory
  python ppt_compare.py file1.pptx file2.pptx ./output
  
Color Coding:
  - Light grey bar: Slide present in both presentations (matched)
  - Red bar: Slide only in source presentation
  - Green bar: Slide only in target presentation
        '''
    )
    
    parser.add_argument('file1', help='First PowerPoint file (source)')
    parser.add_argument('file2', help='Second PowerPoint file (target)')
    parser.add_argument('output_dir', nargs='?', default=None,
                       help='Optional output directory (uses temp dir if not specified)')
    
    suppress_group = parser.add_mutually_exclusive_group()
    suppress_group.add_argument('--suppress-common-slides', dest='suppress_common',
                               action='store_true', default=True,
                               help='Suppress slides present in both presentations (default)')
    suppress_group.add_argument('--no-suppress-common-slides', dest='suppress_common',
                               action='store_false',
                               help='Show all slides including common ones')
    
    args = parser.parse_args()
    
    file1 = args.file1
    file2 = args.file2
    output_dir = args.output_dir
    suppress_common = args.suppress_common
    
    # Determine if we should use temporary directory and clean up
    use_temp_dir = output_dir is None
    
    # Validate files exist
    if not os.path.exists(file1):
        print(f"Error: File not found: {file1}")
        sys.exit(1)
    
    if not os.path.exists(file2):
        print(f"Error: File not found: {file2}")
        sys.exit(1)
    
    # Create or use output directory
    if use_temp_dir:
        base_temp_dir = tempfile.mkdtemp(prefix="ppt_compare_")
        print(f"Created temporary directory: {base_temp_dir}")
    else:
        base_temp_dir = output_dir
        os.makedirs(base_temp_dir, exist_ok=True)
        print(f"Using output directory: {base_temp_dir}")
    
    try:
        # Process both PowerPoint files
        output_dir1 = process_powerpoint(file1, base_temp_dir)
        output_dir2 = process_powerpoint(file2, base_temp_dir)
        
        print("\n" + "="*60)
        print("CONVERSION COMPLETE")
        print("="*60)
        print(f"\nFile 1 images: {output_dir1}")
        print(f"File 2 images: {output_dir2}")
        print(f"\nBase output directory: {base_temp_dir}")
        
        if not use_temp_dir:
            print("\nNote: Output files have been saved and will NOT be deleted.")
        
        # Compare slides between the two presentations
        comparisons, hashes1, hashes2 = compare_slides(output_dir1, output_dir2)
        
        # Generate comparison PDF
        pdf_path = os.path.join(base_temp_dir, "comparison.pdf")
        generate_comparison_pdf(output_dir1, output_dir2, pdf_path, comparisons, suppress_common)
        
        print(f"\nComparison PDF: {pdf_path}")
        
        # Open the PDF
        print("\nOpening PDF...")
        try:
            if sys.platform == 'darwin':  # macOS
                subprocess.run(['open', pdf_path], check=True)
            elif sys.platform == 'win32':  # Windows
                os.startfile(pdf_path)
            else:  # Linux
                subprocess.run(['xdg-open', pdf_path], check=True)
            print("PDF opened successfully")
        except Exception as e:
            print(f"Could not open PDF automatically: {e}")
            print(f"Please open manually: {pdf_path}")
        
        # If using temporary directory, wait for user to view PDF then clean up
        if use_temp_dir:
            input("\nPress Enter to close the PDF and clean up temporary files...")
            print("\nCleaning up temporary files...")
            shutil.rmtree(base_temp_dir)
            print("Temporary files deleted.")
        
    except Exception as e:
        print(f"\nError during processing: {e}")
        print(f"\nTemporary directory (may contain partial results): {base_temp_dir}")
        sys.exit(1)


if __name__ == "__main__":
    main()

# Made with Bob

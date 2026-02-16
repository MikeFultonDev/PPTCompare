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
    """Compare slides between two directories and create a mapping.
    Handles duplicate slides (same hash) correctly by tracking counts."""
    print("\n" + "="*60)
    print("SLIDE COMPARISON")
    print("="*60)
    
    hashes1 = load_slide_hashes(dir1)
    hashes2 = load_slide_hashes(dir2)
    
    # Count occurrences of each hash in both presentations
    from collections import defaultdict, Counter
    
    hash_count1 = Counter(hashes1.values())
    hash_count2 = Counter(hashes2.values())
    
    # Create lists of slides grouped by hash for matching
    hash_to_slides1 = defaultdict(list)
    hash_to_slides2 = defaultdict(list)
    
    for slide_num, hash_val in hashes1.items():
        hash_to_slides1[hash_val].append(slide_num)
    
    for slide_num, hash_val in hashes2.items():
        hash_to_slides2[hash_val].append(slide_num)
    
    # Sort the lists to ensure consistent matching
    for hash_val in hash_to_slides1:
        hash_to_slides1[hash_val].sort()
    for hash_val in hash_to_slides2:
        hash_to_slides2[hash_val].sort()
    
    # Track which slides have been matched
    matched_slides1 = set()
    matched_slides2 = set()
    
    # Store comparison results for PDF generation
    comparisons = []
    
    # Match slides with the same hash, handling duplicates
    for slide1 in sorted(hashes1.keys()):
        hash1 = hashes1[slide1]
        
        if hash1 in hash_to_slides2:
            # Find an unmatched slide2 with the same hash
            available_slides2 = [s for s in hash_to_slides2[hash1] if s not in matched_slides2]
            
            if available_slides2:
                # Match with the first available slide
                slide2 = available_slides2[0]
                matched_slides1.add(slide1)
                matched_slides2.add(slide2)
                print(f"slide {slide1} -> slide {slide2}")
                comparisons.append(('matched', slide1, slide2))
            else:
                # All slides with this hash in dir2 are already matched
                print(f"slide {slide1} only in source (duplicate)")
                comparisons.append(('source_only', slide1, None))
        else:
            print(f"slide {slide1} only in source")
            comparisons.append(('source_only', slide1, None))
    
    # Find slides only in dir2 (including unmatched duplicates)
    for slide2 in sorted(hashes2.keys()):
        if slide2 not in matched_slides2:
            print(f"slide {slide2} only in target")
            comparisons.append(('target_only', None, slide2))
    
    print("="*60)
    
    return comparisons, hashes1, hashes2


def generate_comparison_pdf(dir1, dir2, output_path, comparisons, suppress_common=True, show_moved_pages=True):
    """Generate a PDF with side-by-side slide comparisons"""
    print("\n" + "="*60)
    print("GENERATING COMPARISON PDF")
    print("="*60)
    
    if suppress_common:
        print("Suppressing common slides (matched slides will be excluded)")
    else:
        print("Including all slides (matched slides will be shown)")
    
    if show_moved_pages:
        print("Showing moved pages with arrows for repositioned slides")
        print("Slides will be shown in original file order with arrows indicating mappings")
    
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
    
    # If show_moved_pages is enabled, show slides in original order on both sides
    if show_moved_pages:
        # Build mapping from source to target
        source_to_target = {}
        target_to_source = {}
        source_only = set()
        target_only = set()
        
        for comparison_type, slide1, slide2 in comparisons:
            if comparison_type == 'matched':
                source_to_target[slide1] = slide2
                target_to_source[slide2] = slide1
            elif comparison_type == 'source_only':
                source_only.add(slide1)
            elif comparison_type == 'target_only':
                target_only.add(slide2)
        
        # Get all slides in original order
        all_source_slides = sorted(set(source_to_target.keys()) | source_only)
        all_target_slides = sorted(set(target_to_source.keys()) | target_only)
        
        # Create pages showing both sides in original order
        max_slides = max(len(all_source_slides), len(all_target_slides))
        
        for i in range(max_slides):
            slide1 = all_source_slides[i] if i < len(all_source_slides) else None
            slide2 = all_target_slides[i] if i < len(all_target_slides) else None
            
            # Determine what to show on this page
            if slide1 and slide2:
                # Both sides have slides at this position
                if slide1 in source_only and slide2 in target_only:
                    comparison_type = 'both_unmatched'
                elif slide1 in source_only:
                    comparison_type = 'mixed_source_only'
                elif slide2 in target_only:
                    comparison_type = 'mixed_target_only'
                else:
                    # Both are matched (but possibly to different slides)
                    comparison_type = 'both_matched'
            elif slide1:
                comparison_type = 'source_only'
            elif slide2:
                comparison_type = 'target_only'
            else:
                continue
            
            # Skip if suppress_common and both match at same position
            if suppress_common and comparison_type == 'both_matched':
                if slide1 in source_to_target and source_to_target[slide1] == slide2:
                    continue
            
            # Calculate arrow information for cross-page arrows
            arrow_info = None
            if slide1 and slide1 in source_to_target:
                target_slide = source_to_target[slide1]
                target_page = all_target_slides.index(target_slide) + 1 if target_slide in all_target_slides else None
                current_page = i + 1
                if target_page and target_page != current_page:
                    arrow_info = {
                        'source_slide': slide1,
                        'target_slide': target_slide,
                        'target_page': target_page,
                        'current_page': current_page,
                        'direction': 'up' if target_page < current_page else 'down'
                    }
            
            # Process this page
            _render_comparison_page_with_arrows(c, dir1, dir2, comparison_type, slide1, slide2,
                                               page_width, page_height, margin, bar_width,
                                               available_width, available_height,
                                               arrow_info, source_to_target, target_to_source)
            pages_added += 1
    else:
        # Original behavior: iterate through comparisons as-is (no arrows)
        for comparison_type, slide1, slide2 in comparisons:
            # Skip matched slides if suppress_common is True
            if suppress_common and comparison_type == 'matched':
                continue
            
            # Build simple mapping for non-arrow mode
            source_to_target = {}
            target_to_source = {}
            for ct, s1, s2 in comparisons:
                if ct == 'matched':
                    source_to_target[s1] = s2
                    target_to_source[s2] = s1
            
            _render_comparison_page_with_arrows(c, dir1, dir2, comparison_type, slide1, slide2,
                                               page_width, page_height, margin, bar_width,
                                               available_width, available_height,
                                               None, source_to_target, target_to_source)
            pages_added += 1
    
    c.save()
    print(f"PDF saved to: {output_path}")
    print(f"Total pages: {pages_added}")
    print("="*60)


def _render_comparison_page_with_arrows(c, dir1, dir2, comparison_type, slide1, slide2,
                                        page_width, page_height, margin, bar_width,
                                        available_width, available_height,
                                        arrow_info, source_to_target, target_to_source):
    """Helper function to render a single comparison page with cross-page arrow support"""
    
    # Determine which images to display
    left_image = None
    right_image = None
    title = ""
    
    # Build title and determine images based on what's on this page
    if comparison_type == 'both_matched':
        left_image = os.path.join(dir1, f"slide_{slide1:03d}.png") if slide1 else None
        right_image = os.path.join(dir2, f"slide_{slide2:03d}.png") if slide2 else None
        # Check if they match each other
        if slide1 in source_to_target and source_to_target[slide1] == slide2:
            title = f"Source Slide {slide1} = Target Slide {slide2}"
        else:
            # They're both matched but to different slides
            target_match = source_to_target.get(slide1, '?')
            source_match = target_to_source.get(slide2, '?')
            title = f"Source {slide1} (→{target_match}) | Target {slide2} (←{source_match})"
    elif comparison_type == 'mixed_source_only':
        left_image = os.path.join(dir1, f"slide_{slide1:03d}.png")
        right_image = os.path.join(dir2, f"slide_{slide2:03d}.png") if slide2 else None
        # Check if source is actually matched to a different position
        if slide1 in source_to_target:
            target_match = source_to_target[slide1]
            source_match = target_to_source.get(slide2, '?') if slide2 else '?'
            title = f"Source {slide1} (→{target_match}) | Target {slide2} (←{source_match})"
        else:
            source_match = target_to_source.get(slide2, '?') if slide2 else '?'
            title = f"Source {slide1} (not in target) | Target {slide2} (←{source_match})"
    elif comparison_type == 'mixed_target_only':
        left_image = os.path.join(dir1, f"slide_{slide1:03d}.png") if slide1 else None
        right_image = os.path.join(dir2, f"slide_{slide2:03d}.png")
        # Check if target is actually matched to a different position
        if slide2 in target_to_source:
            source_match = target_to_source[slide2]
            target_match = source_to_target.get(slide1, '?') if slide1 else '?'
            title = f"Source {slide1} (→{target_match}) | Target {slide2} (←{source_match})"
        else:
            target_match = source_to_target.get(slide1, '?') if slide1 else '?'
            title = f"Source {slide1} (→{target_match}) | Target {slide2} (not in source)"
    elif comparison_type == 'source_only':
        left_image = os.path.join(dir1, f"slide_{slide1:03d}.png")
        right_image = None
        # Check if this source slide is actually matched to a target
        if slide1 in source_to_target:
            target_match = source_to_target[slide1]
            title = f"Source Slide {slide1} (→ Target {target_match})"
        else:
            title = f"Source Slide {slide1} (not in target)"
    elif comparison_type == 'target_only':
        left_image = None
        right_image = os.path.join(dir2, f"slide_{slide2:03d}.png")
        # Check if this target slide is actually matched to a source
        if slide2 in target_to_source:
            source_match = target_to_source[slide2]
            title = f"Target Slide {slide2} (← Source {source_match})"
        else:
            title = f"Target Slide {slide2} (not in source)"
    elif comparison_type == 'both_unmatched':
        left_image = os.path.join(dir1, f"slide_{slide1:03d}.png")
        right_image = os.path.join(dir2, f"slide_{slide2:03d}.png")
        title = f"Source {slide1} (not in target) | Target {slide2} (not in source)"
    else:
        left_image = None
        right_image = None
        title = "Unknown comparison type"
    
    # Determine color bar colors
    left_bar_color = None
    right_bar_color = None
    
    if slide1:
        if slide1 in source_to_target:
            left_bar_color = HexColor('#D3D3D3')  # Grey - matched
        else:
            left_bar_color = HexColor('#FF0000')  # Red - not in target
    
    if slide2:
        if slide2 in target_to_source:
            right_bar_color = HexColor('#D3D3D3')  # Grey - matched
        else:
            right_bar_color = HexColor('#00FF00')  # Green - not in source
    
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
    
    # Draw cross-page arrow indicator if needed
    if arrow_info:
        # Draw indicator showing this source maps to a different page
        arrow_start_x = page_width / 2 - margin / 2
        arrow_y = page_height / 2
        
        # Draw arrow pointing up or down
        c.setStrokeColorRGB(0.2, 0.2, 0.8)  # Blue
        c.setLineWidth(3)
        c.setFillColorRGB(0.2, 0.2, 0.8)
        
        if arrow_info['direction'] == 'up':
            # Draw upward arrow
            arrow_top_y = margin + available_height - 20
            c.line(arrow_start_x, arrow_y, arrow_start_x, arrow_top_y)
            
            # Arrowhead pointing up
            p = c.beginPath()
            p.moveTo(arrow_start_x, arrow_top_y)
            p.lineTo(arrow_start_x - 8, arrow_top_y - 12)
            p.lineTo(arrow_start_x + 8, arrow_top_y - 12)
            p.close()
            c.drawPath(p, fill=1, stroke=0)
            
            # Label
            c.setFont("Helvetica-Bold", 10)
            label_text = f"→ Page {arrow_info['target_page']}"
            c.drawString(arrow_start_x + 15, arrow_top_y - 10, label_text)
        else:
            # Draw downward arrow
            arrow_bottom_y = margin + 20
            c.line(arrow_start_x, arrow_y, arrow_start_x, arrow_bottom_y)
            
            # Arrowhead pointing down
            p = c.beginPath()
            p.moveTo(arrow_start_x, arrow_bottom_y)
            p.lineTo(arrow_start_x - 8, arrow_bottom_y + 12)
            p.lineTo(arrow_start_x + 8, arrow_bottom_y + 12)
            p.close()
            c.drawPath(p, fill=1, stroke=0)
            
            # Label
            c.setFont("Helvetica-Bold", 10)
            label_text = f"→ Page {arrow_info['target_page']}"
            c.drawString(arrow_start_x + 15, arrow_bottom_y + 5, label_text)
    
    c.showPage()


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
    
    moved_pages_group = parser.add_mutually_exclusive_group()
    moved_pages_group.add_argument('--show-moved-pages', dest='show_moved_pages',
                                  action='store_true', default=True,
                                  help='Show slides in original order with arrows for repositioned slides (default)')
    moved_pages_group.add_argument('--no-show-moved-pages', dest='show_moved_pages',
                                  action='store_false',
                                  help='Show slides grouped by match status without arrows')
    
    args = parser.parse_args()
    
    file1 = args.file1
    file2 = args.file2
    output_dir = args.output_dir
    suppress_common = args.suppress_common
    show_moved_pages = args.show_moved_pages
    
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
        generate_comparison_pdf(output_dir1, output_dir2, pdf_path, comparisons, suppress_common, show_moved_pages)
        
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

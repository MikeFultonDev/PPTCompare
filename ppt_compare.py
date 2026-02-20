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
import time
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed
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


def convert_ppt_to_pdf_only(ppt_path, output_dir, debug=False, instance_id=None):
    """Convert PowerPoint to PDF using LibreOffice (for parallel processing)
    
    Args:
        ppt_path: Path to PowerPoint file
        output_dir: Directory for output files
        debug: Enable debug output
        instance_id: Unique ID for parallel LibreOffice instances (enables separate user profiles)
        
    Returns:
        Path to generated PDF file
    """
    import subprocess
    
    if debug:
        print(f"  Converting {Path(ppt_path).name} to PDF...")
    
    # Try different LibreOffice command names
    libreoffice_commands = ['libreoffice', 'soffice']
    
    pdf_created = False
    for cmd in libreoffice_commands:
        try:
            # Build command with separate user installation for parallel processing
            cmd_args = [cmd, '--headless', '--convert-to', 'pdf', '--outdir', output_dir]
            
            # Add separate user installation if instance_id provided (for parallel processing)
            if instance_id is not None:
                user_install_dir = f"/tmp/libreoffice_instance_{instance_id}"
                cmd_args.extend(['-env:UserInstallation=file://' + user_install_dir])
                # Also use unique port for socket connection
                port = 2001 + instance_id
                cmd_args.extend([f'--accept=socket,host=127.0.0.1,port={port};urp;'])
            
            cmd_args.append(str(Path(ppt_path).absolute()))
            
            result = subprocess.run(
                cmd_args,
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
    
    return str(pdf_files[0])


def convert_ppt_to_images_libreoffice(ppt_path, output_dir, debug=False, perf_timings=None, pdf_path=None):
    """Convert PowerPoint to images using LibreOffice (cross-platform)
    
    Args:
        ppt_path: Path to PowerPoint file
        output_dir: Directory for output files
        debug: Enable debug output
        perf_timings: Dictionary for performance timing
        pdf_path: Pre-converted PDF path (if already converted in parallel)
    """
    import subprocess
    
    # First convert to PDF (if not already done)
    ppt_name = Path(ppt_path).stem
    
    if pdf_path is None:
        if debug:
            print(f"  Converting {Path(ppt_path).name} to PDF...")
        
        if perf_timings is not None:
            pdf_start = time.time()
        
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
        
        if debug:
            print(f"  PDF created: {temp_pdf}")
        
        if perf_timings is not None:
            pdf_end = time.time()
            perf_timings.setdefault('pptx_to_pdf', 0)
            perf_timings['pptx_to_pdf'] += (pdf_end - pdf_start)
    else:
        temp_pdf = pdf_path
        if debug:
            print(f"  Using pre-converted PDF: {temp_pdf}")
    
    # Convert PDF to images
    if not PDF2IMAGE_AVAILABLE:
        raise RuntimeError("pdf2image not installed. Run: pip install pdf2image poppler-utils")
    
    if debug:
        print(f"  Converting PDF to PNG images...")
    
    if perf_timings is not None:
        png_start = time.time()
    
    images = convert_from_path(temp_pdf, dpi=100)
    
    if perf_timings is not None:
        png_end = time.time()
        perf_timings.setdefault('pdf_to_png', 0)
        perf_timings['pdf_to_png'] += (png_end - png_start)
    
    if perf_timings is not None:
        hash_start = time.time()
    
    for i, image in enumerate(images, start=1):
        output_path = os.path.join(output_dir, f"slide_{i:03d}.png")
        image.save(output_path, "PNG")
        
        # Compute SHA-256 hash
        sha256_hash = compute_sha256(output_path)
        hash_file = os.path.join(output_dir, f"slide_{i:03d}.sha256")
        with open(hash_file, 'w') as f:
            f.write(f"{sha256_hash}  {os.path.basename(output_path)}\n")
        
        if debug:
            print(f"    Slide {i} -> {output_path}")
            print(f"             SHA-256: {sha256_hash}")
    
    if perf_timings is not None:
        hash_end = time.time()
        perf_timings.setdefault('save_and_hash', 0)
        perf_timings['save_and_hash'] += (hash_end - hash_start)
    
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


def compare_slides(dir1, dir2, debug=False):
    """Compare slides between two directories and create a mapping.
    Handles duplicate slides (same hash) correctly by tracking counts."""
    if debug:
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
                if debug:
                    print(f"slide {slide1} -> slide {slide2}")
                comparisons.append(('matched', slide1, slide2))
            else:
                # All slides with this hash in dir2 are already matched
                if debug:
                    print(f"slide {slide1} only in source (duplicate)")
                comparisons.append(('source_only', slide1, None))
        else:
            if debug:
                print(f"slide {slide1} only in source")
            comparisons.append(('source_only', slide1, None))
    
    # Find slides only in dir2 (including unmatched duplicates)
    for slide2 in sorted(hashes2.keys()):
        if slide2 not in matched_slides2:
            if debug:
                print(f"slide {slide2} only in target")
            comparisons.append(('target_only', None, slide2))
    
    if debug:
        print("="*60)
    
    return comparisons, hashes1, hashes2


def generate_comparison_pdf(dir1, dir2, output_path, comparisons, suppress_common=True, show_moved_pages=True, debug=False):
    """Generate a PDF with side-by-side slide comparisons"""
    if debug:
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
    if debug:
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


def open_pdf_and_wait(pdf_path, debug=False):
    """Open a PDF file and wait for the viewer to close before returning.
    
    Args:
        pdf_path: Path to the PDF file to open
        debug: Enable debug output
        
    Returns:
        True if PDF was opened successfully, False otherwise
    """
    import time
    
    if not os.path.exists(pdf_path):
        print(f"Error: PDF file not found: {pdf_path}")
        return False
    
    if debug:
        print(f"\nOpening PDF: {pdf_path}")
        print("Waiting for PDF viewer to close...")
    
    try:
        if sys.platform == 'darwin':  # macOS
            # Get absolute path
            abs_pdf_path = os.path.abspath(pdf_path)
            pdf_filename = os.path.basename(abs_pdf_path)
            
            # Open the PDF with Preview
            subprocess.run(['open', '-a', 'Preview', abs_pdf_path], check=True)
            
            print(f"\nPDF opened in Preview. Close the PDF window when done...")
            if debug:
                print(f"Monitoring for window: {pdf_filename}")
            
            # Wait for Preview to open the file
            time.sleep(2)
            
            # Monitor using AppleScript to check if the specific PDF window is still open
            max_wait_time = 3600  # 1 hour maximum
            start_time = time.time()
            check_interval = 2.0  # Check every 2 seconds
            
            while (time.time() - start_time) < max_wait_time:
                # Use AppleScript to check if Preview has a window with this PDF's name
                applescript = f'''
                tell application "Preview"
                    set windowNames to name of every window
                    set pdfOpen to false
                    repeat with windowName in windowNames
                        if windowName contains "{pdf_filename}" then
                            set pdfOpen to true
                            exit repeat
                        end if
                    end repeat
                    return pdfOpen
                end tell
                '''
                
                try:
                    result = subprocess.run(
                        ['osascript', '-e', applescript],
                        capture_output=True,
                        text=True,
                        timeout=5
                    )
                    
                    # If the script returns "false", the window is closed
                    if result.returncode == 0 and result.stdout.strip() == 'false':
                        if debug:
                            print(f"PDF window closed: {pdf_filename}")
                        break
                    
                except subprocess.TimeoutExpired:
                    if debug:
                        print("AppleScript timeout, continuing...")
                
                time.sleep(check_interval)
            
            if (time.time() - start_time) >= max_wait_time:
                print("\nWarning: Timeout reached while waiting for PDF viewer to close")
            
            return True
            
        elif sys.platform == 'win32':  # Windows
            # Use start /wait with proper shell handling
            result = subprocess.run(
                f'start /wait "" "{pdf_path}"',
                shell=True,
                check=True
            )
            if debug:
                print("PDF viewer closed")
            return True
            
        else:  # Linux and other Unix-like systems
            # xdg-open doesn't wait, so we need to prompt the user
            subprocess.Popen(['xdg-open', pdf_path])
            print("\nPDF opened in default viewer.")
            print("Press Enter when you're done viewing to continue...")
            input()
            if debug:
                print("User confirmed viewer closed")
            return True
            
    except subprocess.CalledProcessError as e:
        print(f"Error opening PDF: {e}")
        return False
    except FileNotFoundError:
        print("Error: Could not find PDF viewer application")
        return False
    except Exception as e:
        print(f"Unexpected error opening PDF: {e}")
        return False


def process_powerpoint(ppt_path, base_temp_dir, debug=False, perf_timings=None, pdf_path=None):
    """Process a single PowerPoint file and convert slides to PNG images
    
    Args:
        ppt_path: Path to PowerPoint file
        base_temp_dir: Base directory for output
        debug: Enable debug output
        perf_timings: Dictionary for performance timing
        pdf_path: Pre-converted PDF path (if already converted in parallel)
    """
    
    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"PowerPoint file not found: {ppt_path}")
    
    # Create output directory for this file
    ppt_name = Path(ppt_path).stem
    output_dir = os.path.join(base_temp_dir, ppt_name)
    os.makedirs(output_dir, exist_ok=True)
    
    if debug:
        print(f"\nProcessing: {ppt_path}")
        print(f"Output directory: {output_dir}")
    
    try:
        slide_count = convert_ppt_to_images_libreoffice(ppt_path, output_dir, debug, perf_timings, pdf_path)
        if debug:
            print(f"  Successfully converted {slide_count} slides")
        return output_dir
    except Exception as e:
        print(f"  Error: {e}")
        raise


def get_git_committed_version(file_path, temp_dir):
    """Get the last committed version of a file from git"""
    try:
        # Convert to absolute path
        abs_file_path = os.path.abspath(file_path)
        
        # Get the git repository root
        git_root_result = subprocess.run(
            ['git', 'rev-parse', '--show-toplevel'],
            capture_output=True,
            check=True,
            text=True,
            cwd=os.path.dirname(abs_file_path)
        )
        git_root = git_root_result.stdout.strip()
        
        # Get relative path from git root
        rel_path = os.path.relpath(abs_file_path, git_root)
        
        # Get the file content from the last commit
        result = subprocess.run(
            ['git', 'show', f'HEAD:{rel_path}'],
            capture_output=True,
            check=True,
            cwd=git_root
        )
        
        # Create a temporary file with the committed version
        file_name = Path(file_path).name
        base_name = Path(file_path).stem
        extension = Path(file_path).suffix
        
        committed_file = os.path.join(temp_dir, f"{base_name}_committed{extension}")
        with open(committed_file, 'wb') as f:
            f.write(result.stdout)
        
        return committed_file
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Failed to get committed version of {file_path}: {e.stderr.decode()}")


def print_performance_report(timings):
    """Print a performance report showing time spent in each stage"""
    print("\n" + "="*60)
    print("PERFORMANCE REPORT")
    print("="*60)
    
    total_time = timings.get('total', 0)
    
    # Calculate stage times
    stages = [
        ('Setup & Validation', timings.get('setup_end', 0) - timings.get('start', 0)),
        ('PPTX→PDF (Parallel)', timings.get('pdf_convert_end', 0) - timings.get('pdf_convert_start', 0)),
        ('PDF→PNG + Hashing', timings.get('convert_end', 0) - timings.get('convert_start', 0)),
        ('Compare Slides', timings.get('compare_end', 0) - timings.get('compare_start', 0)),
        ('Generate PDF', timings.get('pdf_end', 0) - timings.get('pdf_start', 0)),
    ]
    
    print(f"\n{'Stage':<30} {'Time (s)':<12} {'% of Total':<12}")
    print("-" * 60)
    
    for stage_name, stage_time in stages:
        if stage_time > 0:
            percentage = (stage_time / total_time * 100) if total_time > 0 else 0
            print(f"{stage_name:<30} {stage_time:>10.2f}s  {percentage:>10.1f}%")
    
    print("-" * 60)
    print(f"{'TOTAL':<30} {total_time:>10.2f}s  {100.0:>10.1f}%")
    
    # Print detailed breakdown of conversion stages
    if any(key in timings for key in ['pptx_to_pdf', 'pdf_to_png', 'save_and_hash']):
        print("\nDetailed Conversion Breakdown:")
        print("-" * 60)
        
        conversion_stages = [
            ('  PPTX to PDF (LibreOffice)', timings.get('pptx_to_pdf', 0)),
            ('  PDF to PNG (pdf2image)', timings.get('pdf_to_png', 0)),
            ('  Save PNG & Compute Hashes', timings.get('save_and_hash', 0)),
        ]
        
        for stage_name, stage_time in conversion_stages:
            if stage_time > 0:
                percentage = (stage_time / total_time * 100) if total_time > 0 else 0
                print(f"{stage_name:<30} {stage_time:>10.2f}s  {percentage:>10.1f}%")
    
    print("="*60)


def main():
    """Main function to compare two PowerPoint files"""
    
    parser = argparse.ArgumentParser(
        description='Compare two PowerPoint presentations and generate a side-by-side comparison PDF',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  # Activate virtual environment first (required)
  source venv/bin/activate  # On Windows: venv\\Scripts\\activate
  
  # Compare two files
  python ppt_compare.py file1.pptx file2.pptx
  
  # Compare with performance timing
  python ppt_compare.py file1.pptx file2.pptx --perf
  
  # Compare current file with last committed version
  python ppt_compare.py presentation.pptx --git
  
  # Show all slides including common ones
  python ppt_compare.py file1.pptx file2.pptx --no-suppress-common-slides
  
  # Save to specific directory (files preserved)
  python ppt_compare.py file1.pptx file2.pptx ./output
  
  # With debug output
  python ppt_compare.py file1.pptx file2.pptx --debug

Performance:
  Uses parallel processing for optimal speed (~9s for 24-28 slide presentations)
  - PPTX→PDF: Parallel conversion with separate LibreOffice instances
  - PDF→PNG: Parallel conversion for both files
  - Use --perf flag to see detailed timing breakdown

Color Coding:
  - Light grey bar: Slide present in both presentations (matched)
  - Red bar: Slide only in source presentation
  - Green bar: Slide only in target presentation
        '''
    )
    
    parser.add_argument('file1', help='First PowerPoint file (source), or the only file when using --git')
    parser.add_argument('file2', nargs='?', default=None, help='Second PowerPoint file (target), not used with --git')
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
    
    parser.add_argument('--debug', action='store_true',
                       help='Enable debug output showing detailed processing information')
    
    parser.add_argument('--perf', action='store_true',
                       help='Show performance timing for different stages of processing')
    
    parser.add_argument('--git', action='store_true', default=False,
                       help='Compare current file with last committed version (only file1 is used)')
    
    args = parser.parse_args()
    
    file1 = args.file1
    file2 = args.file2
    output_dir = args.output_dir
    suppress_common = args.suppress_common
    show_moved_pages = args.show_moved_pages
    debug = args.debug
    perf = args.perf
    use_git = args.git
    
    # Initialize performance timing dictionary
    timings = {}
    if perf:
        timings['start'] = time.time()
    
    # Validate git mode usage
    if use_git:
        if file2 is not None and file2 != output_dir:
            print("Error: When using --git, only specify one file")
            print("Usage: python ppt_compare.py file.pptx --git [output_dir]")
            sys.exit(1)
        # In git mode, file2 becomes output_dir if provided
        if file2 is not None:
            output_dir = file2
            file2 = None
    
    # Determine if we should use temporary directory and clean up
    use_temp_dir = output_dir is None
    
    # Validate file1 exists
    if not os.path.exists(file1):
        print(f"Error: File not found: {file1}")
        sys.exit(1)
    
    # Handle git mode or regular mode
    if use_git:
        # Create temporary directory for git committed version
        git_temp_dir = tempfile.mkdtemp(prefix="ppt_git_")
        try:
            if debug:
                print(f"Getting committed version of {file1} from git...")
            file2 = get_git_committed_version(file1, git_temp_dir)
            if debug:
                print(f"Committed version saved to: {file2}")
        except Exception as e:
            print(f"Error: {e}")
            shutil.rmtree(git_temp_dir)
            sys.exit(1)
    else:
        # Regular mode - validate file2 exists
        if file2 is None:
            print("Error: Second file required when not using --git")
            print("Usage: python ppt_compare.py file1.pptx file2.pptx [output_dir]")
            print("   or: python ppt_compare.py file.pptx --git [output_dir]")
            sys.exit(1)
        if not os.path.exists(file2):
            print(f"Error: File not found: {file2}")
            sys.exit(1)
    
    # Create or use output directory
    if use_temp_dir:
        base_temp_dir = tempfile.mkdtemp(prefix="ppt_compare_")
        if debug:
            print(f"Created temporary directory: {base_temp_dir}")
    else:
        base_temp_dir = output_dir
        os.makedirs(base_temp_dir, exist_ok=True)
        if debug:
            print(f"Using output directory: {base_temp_dir}")
    
    try:
        if perf:
            timings['setup_end'] = time.time()
            timings['pdf_convert_start'] = time.time()
        
        # Step 1: Convert both PowerPoint files to PDF in parallel
        # Each LibreOffice instance gets its own user profile and socket connection
        ppt_name1 = Path(file1).stem
        ppt_name2 = Path(file2).stem
        output_dir1 = os.path.join(base_temp_dir, ppt_name1)
        output_dir2 = os.path.join(base_temp_dir, ppt_name2)
        os.makedirs(output_dir1, exist_ok=True)
        os.makedirs(output_dir2, exist_ok=True)
        
        with ProcessPoolExecutor(max_workers=2) as executor:
            future1 = executor.submit(convert_ppt_to_pdf_only, file1, output_dir1, debug, 0)
            future2 = executor.submit(convert_ppt_to_pdf_only, file2, output_dir2, debug, 1)
            
            # Wait for both PDF conversions to complete
            pdf_path1 = future1.result()
            pdf_path2 = future2.result()
        
        if perf:
            timings['pdf_convert_end'] = time.time()
            timings['convert_start'] = time.time()
        
        # Step 2: Convert PDFs to images in parallel
        if debug:
            print("\nConverting PDFs to images in parallel...")
        
        with ProcessPoolExecutor(max_workers=2) as executor:
            # Submit both conversions
            future1 = executor.submit(convert_ppt_to_images_libreoffice, file1, output_dir1, debug, None, pdf_path1)
            future2 = executor.submit(convert_ppt_to_images_libreoffice, file2, output_dir2, debug, None, pdf_path2)
            
            # Wait for both to complete and get results
            slide_count1 = future1.result()
            slide_count2 = future2.result()
        
        if debug:
            print(f"\nFile 1: Successfully converted {slide_count1} slides")
            print(f"File 2: Successfully converted {slide_count2} slides")
        
        if perf:
            timings['convert_end'] = time.time()
        
        if debug:
            print("\n" + "="*60)
            print("CONVERSION COMPLETE")
            print("="*60)
            print(f"\nFile 1 images: {output_dir1}")
            print(f"File 2 images: {output_dir2}")
            print(f"\nBase output directory: {base_temp_dir}")
            
            if not use_temp_dir:
                print("\nNote: Output files have been saved and will NOT be deleted.")
        
        if perf:
            timings['compare_start'] = time.time()
        
        # Compare slides between the two presentations
        comparisons, hashes1, hashes2 = compare_slides(output_dir1, output_dir2, debug)
        
        if perf:
            timings['compare_end'] = time.time()
            timings['pdf_start'] = time.time()
        
        # Generate comparison PDF
        pdf_path = os.path.join(base_temp_dir, "comparison.pdf")
        generate_comparison_pdf(output_dir1, output_dir2, pdf_path, comparisons, suppress_common, show_moved_pages, debug)
        
        if perf:
            timings['pdf_end'] = time.time()
            timings['total'] = timings['pdf_end'] - timings['start']
            print_performance_report(timings)
        
        # Check if PDF has content (file size > 1KB indicates it has pages)
        pdf_has_content = os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1024
        
        if pdf_has_content:
            print(f"\nComparison PDF: {pdf_path}")
            
            # Open the PDF and wait for viewer to close
            pdf_opened = open_pdf_and_wait(pdf_path, debug)
            
            if not pdf_opened:
                print(f"\nCould not open PDF automatically.")
                print(f"Please open manually: {pdf_path}")
                print("Press Enter when done viewing to clean up...")
                input()
        else:
            print(f"\nNo differences found between the presentations.")
            print(f"All slides are identical (comparison PDF not generated).")
            if suppress_common:
                print(f"Tip: Use --no-suppress-common-slides to see all slides in the comparison.")
        
        # Clean up temporary directories after PDF viewer closes
        if use_temp_dir:
            if debug:
                print("\nCleaning up temporary files...")
            shutil.rmtree(base_temp_dir)
            if debug:
                print("Temporary files deleted.")
        
        # Clean up git temporary directory if used
        if use_git:
            if debug:
                print("Cleaning up git temporary files...")
            shutil.rmtree(git_temp_dir)
        
    except Exception as e:
        print(f"\nError during processing: {e}")
        print(f"\nTemporary directory (may contain partial results): {base_temp_dir}")
        if use_git and 'git_temp_dir' in locals():
            shutil.rmtree(git_temp_dir)
        sys.exit(1)


if __name__ == "__main__":
    main()

# Made with Bob

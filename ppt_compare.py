#!/usr/bin/env python3
"""
PowerPoint Comparison Tool
Converts PowerPoint slides to PNG images for comparison
"""

import os
import sys
import tempfile
import hashlib
from pathlib import Path

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
    
    # Track which slides in dir2 have been matched
    matched_slides2 = set()
    
    # Compare slides from dir1
    for slide1 in sorted(hashes1.keys()):
        hash1 = hashes1[slide1]
        if hash1 in hash_to_slide2:
            slide2 = hash_to_slide2[hash1]
            matched_slides2.add(slide2)
            print(f"slide {slide1} -> slide {slide2}")
        else:
            print(f"slide {slide1} only in source")
    
    # Find slides only in dir2
    for slide2 in sorted(hashes2.keys()):
        if slide2 not in matched_slides2:
            print(f"slide {slide2} only in target")
    
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
    
    if len(sys.argv) != 3:
        print("Usage: python ppt_compare.py <file1.pptx> <file2.pptx>")
        sys.exit(1)
    
    file1 = sys.argv[1]
    file2 = sys.argv[2]
    
    # Validate files exist
    if not os.path.exists(file1):
        print(f"Error: File not found: {file1}")
        sys.exit(1)
    
    if not os.path.exists(file2):
        print(f"Error: File not found: {file2}")
        sys.exit(1)
    
    # Create base temporary directory
    base_temp_dir = tempfile.mkdtemp(prefix="ppt_compare_")
    print(f"Created temporary directory: {base_temp_dir}")
    
    try:
        # Process both PowerPoint files
        output_dir1 = process_powerpoint(file1, base_temp_dir)
        output_dir2 = process_powerpoint(file2, base_temp_dir)
        
        print("\n" + "="*60)
        print("CONVERSION COMPLETE")
        print("="*60)
        print(f"\nFile 1 images: {output_dir1}")
        print(f"File 2 images: {output_dir2}")
        print(f"\nBase temporary directory: {base_temp_dir}")
        print("\nNote: Temporary directories have NOT been deleted.")
        
        # Compare slides between the two presentations
        compare_slides(output_dir1, output_dir2)
        
    except Exception as e:
        print(f"\nError during processing: {e}")
        print(f"\nTemporary directory (may contain partial results): {base_temp_dir}")
        sys.exit(1)


if __name__ == "__main__":
    main()

# Made with Bob

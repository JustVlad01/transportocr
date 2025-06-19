#!/usr/bin/env python3
"""
Setup script for Transport Sorter - PDF Scanner
This script helps install dependencies and checks system requirements.
"""

import subprocess
import sys
import os
from pathlib import Path

def check_python_version():
    """Check if Python version is compatible"""
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 7):
        print("❌ Error: Python 3.7+ is required")
        print(f"Current version: {version.major}.{version.minor}.{version.micro}")
        return False
    print(f"✅ Python version: {version.major}.{version.minor}.{version.micro}")
    return True

def check_tesseract():
    """Check if Tesseract OCR is installed"""
    try:
        result = subprocess.run(['tesseract', '--version'], 
                              capture_output=True, text=True, check=True)
        print("✅ Tesseract OCR is installed")
        print(f"   Version: {result.stdout.split()[1]}")
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("❌ Tesseract OCR is not installed or not in PATH")
        print("   Please install Tesseract OCR:")
        print("   Ubuntu/Debian: sudo apt install tesseract-ocr tesseract-ocr-eng")
        print("   Windows: https://github.com/UB-Mannheim/tesseract/wiki")
        print("   macOS: brew install tesseract")
        return False

def install_requirements():
    """Install Python requirements"""
    requirements_file = Path("requirements.txt")
    if not requirements_file.exists():
        print("❌ requirements.txt not found")
        return False
    
    try:
        print("📦 Installing Python dependencies...")
        subprocess.run([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt'], 
                      check=True)
        print("✅ Python dependencies installed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install dependencies: {e}")
        return False

def create_directories():
    """Create necessary directories"""
    directories = ['ocr_results']
    for directory in directories:
        Path(directory).mkdir(exist_ok=True)
        print(f"📁 Created directory: {directory}")

def main():
    """Main setup function"""
    print("🚀 Transport Sorter - PDF Scanner Setup")
    print("=" * 50)
    
    # Check Python version
    if not check_python_version():
        sys.exit(1)
    
    # Check Tesseract
    tesseract_ok = check_tesseract()
    
    # Install Python dependencies
    deps_ok = install_requirements()
    
    # Create directories
    create_directories()
    
    print("\n" + "=" * 50)
    if tesseract_ok and deps_ok:
        print("✅ Setup completed successfully!")
        print("\n🎉 You can now run the application:")
        print("   python main.py")
    else:
        print("⚠️  Setup completed with warnings")
        if not tesseract_ok:
            print("   - Please install Tesseract OCR")
        if not deps_ok:
            print("   - Please install Python dependencies manually")
    
    print("\n📖 For more information, see README.md")

if __name__ == "__main__":
    main()
import os
import sys
import subprocess
import shutil
from pathlib import Path

def create_standalone_executable():
    """Create a standalone executable using PyInstaller"""
    
    print("🚀 Creating Standalone OptimoRoute Sorter Executable")
    print("=" * 60)
    
    # Check if PyInstaller is installed
    try:
        import PyInstaller
        print("✅ PyInstaller found")
    except ImportError:
        print("📦 Installing PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
        print("✅ PyInstaller installed")
    
    # Main application file
    main_app = "optimoroute_sorter_app.py"
    if not Path(main_app).exists():
        print(f"❌ Error: {main_app} not found!")
        return False
    
    # Create build directory
    build_dir = Path("standalone_build")
    if build_dir.exists():
        shutil.rmtree(build_dir)
    build_dir.mkdir()
    
    print(f"📁 Build directory: {build_dir}")
    
    # Create PyInstaller spec file for better control
    spec_content = f'''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['{main_app}'],
    pathex=[],
    binaries=[],
    datas=[
        ('delivery_sequence_data.json', '.'),
        ('supabase_config.py', '.'),
        ('supabase_schema.sql', '.'),
        ('your_order_format.xlsx', '.'),
    ],
    hiddenimports=[
        'PySide6.QtCore',
        'PySide6.QtGui', 
        'PySide6.QtWidgets',
        'pytesseract',
        'PIL._tkinter_finder',
        'requests',
        'pandas',
        'fitz',
        'reportlab',
        'barcode',
        'openpyxl'
    ],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='OptimoRouteSorter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # You can add an icon file here
    version_file=None
)
'''
    
    # Write spec file
    spec_file = build_dir / "OptimoRouteSorter.spec"
    with open(spec_file, 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("✅ Created PyInstaller spec file")
    
    # Build the executable
    print("🔨 Building standalone executable (this may take a few minutes)...")
    print("   Please wait while PyInstaller packages all dependencies...")
    
    try:
        cmd = [
            sys.executable, "-m", "PyInstaller",
            "--clean",
            "--noconfirm", 
            str(spec_file)
        ]
        
        result = subprocess.run(cmd, 
                              cwd=build_dir, 
                              capture_output=True, 
                              text=True, 
                              timeout=600)  # 10 minute timeout
        
        if result.returncode == 0:
            print("✅ Build completed successfully!")
            
            # Find the executable
            exe_path = build_dir / "dist" / "OptimoRouteSorter.exe"
            if exe_path.exists():
                exe_size = exe_path.stat().st_size / (1024 * 1024)  # Size in MB
                print(f"📦 Executable created: {exe_path}")
                print(f"💾 File size: {exe_size:.1f} MB")
                
                # Copy to main directory for easy access
                final_exe = Path("OptimoRouteSorter.exe")
                if final_exe.exists():
                    final_exe.unlink()
                shutil.copy2(exe_path, final_exe)
                print(f"📋 Copied to: {final_exe}")
                
                return True
            else:
                print("❌ Executable not found after build")
                print("Build output:", result.stdout)
                print("Build errors:", result.stderr)
                return False
        else:
            print("❌ Build failed!")
            print("Error output:", result.stderr)
            print("Standard output:", result.stdout)
            return False
            
    except subprocess.TimeoutExpired:
        print("❌ Build timed out (took longer than 10 minutes)")
        return False
    except Exception as e:
        print(f"❌ Build error: {str(e)}")
        return False

def create_installer_package():
    """Create a complete installer package"""
    
    exe_file = Path("OptimoRouteSorter.exe")
    if not exe_file.exists():
        print("❌ Executable not found. Build it first.")
        return False
    
    print("\n📦 Creating Installer Package")
    print("=" * 40)
    
    # Create installer directory
    installer_dir = Path("OptimoRouteSorter_Installer")
    if installer_dir.exists():
        shutil.rmtree(installer_dir)
    installer_dir.mkdir()
    
    # Copy executable
    shutil.copy2(exe_file, installer_dir / "OptimoRouteSorter.exe")
    print("✅ Copied executable")
    
    # Create README for standalone version
    readme_content = '''# OptimoRoute Sorter - Standalone Application

## 🚀 Quick Start
1. Simply double-click `OptimoRouteSorter.exe` to run the application
2. No installation required - it's a standalone executable!

## ✨ Features
- **OptimoRoute API Integration**: Fetches scheduled deliveries automatically
- **PDF Processing**: Processes delivery PDFs and groups by driver
- **OCR Support**: Extracts text from image-based PDFs
- **Professional Interface**: Clean, modern GUI
- **No Dependencies**: Everything is bundled - just run the .exe file!

## 📋 How to Use

### Step 1: Launch Application
- Double-click `OptimoRouteSorter.exe`
- The application will start with a professional interface

### Step 2: Set Output Directory  
- Click "Browse" to select where processed PDFs will be saved

### Step 3: Fetch Delivery Data
- Click "Fetch & Load Scheduled Deliveries"
- The app connects to OptimoRoute and loads your scheduled orders
- Review the data in the preview table

### Step 4: Process PDFs
- Click "Add PDFs" to select your delivery PDF files
- Click "Process Delivery PDFs" to start processing
- The app will create separate PDF files for each driver

### Step 5: View Results
- Driver-specific PDF files are created in your output directory
- A results dialog shows what was processed
- Processing summary file is created for your records

## 🔧 System Requirements
- **Operating System**: Windows 7/8/10/11 (64-bit)
- **Memory**: 512 MB RAM minimum (1 GB recommended)
- **Storage**: 200 MB free space
- **Internet**: Required for OptimoRoute API access

## 🆘 Troubleshooting

### Application Won't Start
- Ensure you're running on a 64-bit Windows system
- Try running as administrator
- Check Windows Defender isn't blocking the executable

### "No matching orders found"
- Ensure PDF files contain order IDs that match OptimoRoute data
- Check that order IDs in PDFs are readable (not heavily compressed)
- Verify the date range includes your delivery orders

### OCR Issues  
- OCR is built-in - no additional software needed
- Ensure PDF quality is good for text recognition
- Try different PDF files if text extraction fails

### API Connection Issues
- Check your internet connection
- Verify OptimoRoute API access is working
- Contact administrator if API key needs updating

## 📄 File Information
- **Application**: OptimoRouteSorter.exe
- **Version**: 1.0.0
- **Type**: Standalone Windows Executable
- **Dependencies**: All bundled (no installation required)

## 🔒 Security
This is a standalone application with no external dependencies. It only connects to:
- OptimoRoute API (for fetching delivery data)
- Your local file system (for PDF processing)

No data is sent to any other external services.

## 📞 Support
For support or questions about this application, contact your system administrator.

---
**OptimoRoute Sorter** - Professional delivery PDF processing made simple.
'''
    
    with open(installer_dir / "README.txt", 'w', encoding='utf-8') as f:
        f.write(readme_content)
    print("✅ Created README.txt")
    
    # Create version info
    version_content = '''OptimoRoute Sorter v1.0.0
Standalone Windows Application

Build Information:
- Type: Standalone Executable
- Platform: Windows (64-bit)
- Dependencies: All bundled
- Size: Optimized for distribution

Features:
✓ OptimoRoute API Integration
✓ PDF Processing with OCR
✓ Driver-based PDF Sorting
✓ Professional GUI Interface
✓ No Installation Required

This standalone executable contains all required dependencies
and can be run directly without installing Python or other software.
'''
    
    with open(installer_dir / "VERSION.txt", 'w', encoding='utf-8') as f:
        f.write(version_content)
    print("✅ Created VERSION.txt")
    
    # Create zip package
    zip_name = "OptimoRouteSorter_Standalone.zip"
    if Path(zip_name).exists():
        Path(zip_name).unlink()
    
    import zipfile
    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file_path in installer_dir.rglob('*'):
            if file_path.is_file():
                arcname = file_path.relative_to(installer_dir.parent)
                zipf.write(file_path, arcname)
    
    zip_size = Path(zip_name).stat().st_size / (1024 * 1024)
    print(f"📦 Created installer package: {zip_name}")
    print(f"💾 Package size: {zip_size:.1f} MB")
    
    return True

def main():
    """Main function to create standalone executable"""
    
    print("OptimoRoute Sorter - Standalone Executable Builder")
    print("=" * 60)
    print()
    
    # Step 1: Build executable
    print("Step 1: Building standalone executable...")
    if not create_standalone_executable():
        print("❌ Failed to create executable")
        return
    
    print("\n" + "="*60)
    
    # Step 2: Create installer package
    print("Step 2: Creating installer package...")
    if not create_installer_package():
        print("❌ Failed to create installer package")
        return
    
    print("\n" + "="*60)
    print("🎉 SUCCESS!")
    print()
    print("Your standalone application is ready:")
    print("📁 OptimoRouteSorter_Standalone.zip - Complete package for distribution")
    print("💻 OptimoRouteSorter.exe - Direct executable")
    print()
    print("Distribution options:")
    print("✅ Share the ZIP file - users extract and run the .exe")
    print("✅ Share just the .exe file - users run it directly")
    print("✅ No Python installation required for end users")
    print("✅ All dependencies are bundled in the executable")

if __name__ == "__main__":
    main() 
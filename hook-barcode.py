"""
PyInstaller runtime hook for barcode library
This ensures barcode library can find its font files and resources when running from executable
"""
import os
import sys

def setup_barcode_resources():
    """Setup barcode library resources for PyInstaller executable"""
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller executable
        base_path = sys._MEIPASS
        
        # Add barcode fonts directory to the path
        barcode_fonts_path = os.path.join(base_path, 'barcode', 'fonts')
        if os.path.exists(barcode_fonts_path):
            # Set environment variable for barcode library to find fonts
            os.environ['BARCODE_FONTS_PATH'] = barcode_fonts_path
            
        # Also try to patch the barcode library's font loading
        try:
            import barcode
            barcode_package_dir = os.path.dirname(barcode.__file__)
            
            # If we're in a PyInstaller bundle, update the font path
            if hasattr(sys, '_MEIPASS'):
                # Override the font path in barcode library
                original_font_path = os.path.join(barcode_package_dir, 'fonts')
                if not os.path.exists(original_font_path):
                    # Use the bundled fonts
                    bundled_fonts = os.path.join(sys._MEIPASS, 'barcode', 'fonts')
                    if os.path.exists(bundled_fonts):
                        # Monkey patch the font loading
                        import barcode.writer.base
                        original_get_font = barcode.writer.base.get_font
                        
                        def patched_get_font(font_name, size):
                            try:
                                return original_get_font(font_name, size)
                            except (OSError, IOError):
                                # Try bundled fonts
                                bundled_font_file = os.path.join(bundled_fonts, f"{font_name}.ttf")
                                if os.path.exists(bundled_font_file):
                                    from PIL import ImageFont
                                    return ImageFont.truetype(bundled_font_file, size)
                                raise
                        
                        barcode.writer.base.get_font = patched_get_font
        except Exception as e:
            print(f"Warning: Could not setup barcode resources: {e}")

# Run the setup when this module is imported
setup_barcode_resources()


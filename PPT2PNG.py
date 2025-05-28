#!python3
'''
Automatically convert a PPT/PPTX to PNG files.
This script extracts all slides from PowerPoint presentations as PNG images.

PyWin32 is needed. This script works on Windows only.

@author: ern (original)
@blog: www.readern.com
Modified version: Removed PDF conversion and watermark functionality
'''

import os
import win32com.client
import shutil

def ppt2png(filename, dst_filename):
    """Convert PowerPoint presentation to PNG images
    
    Args:
        filename: Path to the PPT/PPTX file
        dst_filename: Output PNG file path (will be used as base name)
    """
    print(f"Opening PowerPoint: {filename}")
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    pptSel = ppt.Presentations.Open(filename, WithWindow=False)
    print(f"Saving as PNG: {dst_filename}")
    pptSel.SaveAs(dst_filename, 18)  # 18 is the value for PNG format
    ppt.Quit()
    print("PowerPoint conversion completed")

def main():
    print("PPT to PNG Converter - Starting...")
    ppt_dir = os.getcwd()
    
    try:
        # Find all PowerPoint files in the current directory
        ppt_files = [fn for fn in os.listdir(ppt_dir) if fn.endswith(('.ppt','.pptx'))]
        
        if not ppt_files:
            print("Error: No PPT/PPTX files found in the current directory.")
            print("Please place your PowerPoint files in this directory: " + ppt_dir)
            return
            
        # Process each PowerPoint file
        for fn in ppt_files:
            file_name = os.path.splitext(fn)[0]
            print("\n" + "="*50)
            print(f"Processing: {file_name}")
            ppt_file = os.path.join(ppt_dir, fn)
            img_file = os.path.join(ppt_dir, file_name+'.png')

            print("Converting PPT to PNG images...")
            ppt2png(ppt_file, img_file)
            
            # The folder with the same name as the PPT file will contain all slides as PNG images
            img_dir = os.path.join(ppt_dir, file_name)
            
            if os.path.exists(img_dir):
                print(f"PNG images extracted successfully to: {img_dir}")
                print(f"Total slides converted: {len([f for f in os.listdir(img_dir) if f.endswith('.PNG')])}")
            else:
                print(f"Warning: Expected output directory not found: {img_dir}")
                
            print("="*50)
            
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    
    print("\nProcess completed. All PowerPoint files have been converted to PNG images.")

if __name__ == "__main__":
    main()

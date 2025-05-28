# PPT2PNG

A simple Python script to convert PowerPoint presentations (.ppt/.pptx) to PNG images.

## Description

This tool automatically converts PowerPoint presentations to individual PNG images, with each slide saved as a separate image. The images are stored in a folder with the same name as the original PowerPoint file.

## Requirements

### System Requirements
- Windows operating system (tested on Windows 10)
- Microsoft PowerPoint installed

### Python Requirements
- Python 3.x
- PyWin32 library

## Installation

1. Make sure you have Python 3 installed on your system. You can download it from [python.org](https://www.python.org/downloads/)

2. Install the required PyWin32 library:
   ```
   pip install pywin32
   ```

## Usage

1. Place your PowerPoint files (.ppt or .pptx) in the same folder as the script.

2. Open Command Prompt or PowerShell, navigate to the folder containing the script, and run:
   ```
   python PPT2PNG.py
   ```

3. The script will:
   - Find all PowerPoint files in the current directory
   - Convert each presentation to PNG images
   - Create a folder for each presentation with the same name as the PowerPoint file
   - Save all slides as PNG images in their respective folders

## Example

If you have a file named `presentation.pptx`, after running the script:
- A folder named `presentation` will be created
- All slides will be saved as PNG images in that folder (Slide1.PNG, Slide2.PNG, etc.)

## Troubleshooting

- **No images generated**: Ensure that PowerPoint is properly installed and can be accessed by Python.
- **Script crashes**: Make sure you have the PyWin32 library installed. Run `pip install pywin32` if you haven't already.
- **Access denied errors**: Try running the command prompt as administrator.

## Credits

This script is a simplified version of PPT2PDF by ern (www.readern.com), modified to focus solely on PowerPoint to PNG conversion.

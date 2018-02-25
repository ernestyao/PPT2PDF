# PPT2PDF

Automatically convert a PPT/PPTX to PNG files and combine them to a PDF after applying watermark.
Then every page of the PDF is an image. It could not be copied or converted back to PPT/PPTX.
Watermark texts should be in watermark.txt and this script watermarks every PPT/PPTX in the same folder where the script locates.

这个 Python 3 脚本自动将PPT/PPTX文件转换成PNG图片、打水印并重新合并成PDF文件。因此PDF文件都是图片构成，一般无法拷贝或无损转换回PPT/PPTX
水印是文本形式写在watermark.txt中，对每一行生成一个水印版PDF。脚本自动获取同文件夹下所有PPT/PPTX文件。

PyWin32, Pillow, reportlab, is needed. With this development, Pywin32 222, Pillow 5.0.0, reportlab 3.4.0 are used.
And so the script works on Windows only. Windows 10 64bit is used with development.

需要使用 PyWin32、Pillow、reportlab 三个包。开发时用了Pywin32 222, Pillow 5.0.0, reportlab 3.4.0。所以脚本只能在 Windows 下使用，开发环境为64位的Windows 10.

Usage: Put the script, watermark.txt and the PPT/PPTX file in the same folder and then open command line window: `python PPT2PDF.py`
使用：把脚本、watermark.txt 和 PPT/PPTX 放在同一个目录下，打开命令行窗口运行 `python PPT2PDF.py`

@author: ern

@blog: www.readern.com


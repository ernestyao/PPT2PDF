#!python3
'''
Automatically convert a PPT/PPTX to PNG files and combine them to a PDF after applying watermark.
Then every page of the PDF is an image. It could not be copied or converted back to PPT/PPTX.
Watermark texts should be in watermark.txt and this script watermarks every PPT/PPTX in the same folder where the script locates.

PyWin32, Pillow, reportlab, is needed. With this development, Pywin32 222, Pillow 5.0.0, reportlab 3.4.0 are used.
And so the script works on Windows only. Windows 10 64bit is used with development.

@author: ern
@blog: www.readern.com

v1.0 Created on Feb 25, 2018
v1.1 Updated with font fix and bug fixes
'''

import os
import time
import math
import codecs
import win32com
import shutil
from win32com.client import Dispatch, constants
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import letter, A4, landscape  
from reportlab.platypus import SimpleDocTemplate
from reportlab.lib.units import inch  
from reportlab.pdfgen import canvas
from reportlab import rl_settings

#PPT转成PNG文件
def ppt2png(filename,dst_filename):
	ppt = win32com.client.Dispatch('PowerPoint.Application')
	#ppt.DisplayAlerts = False
	pptSel = ppt.Presentations.Open(filename, WithWindow = False)
	pptSel.SaveAs(dst_filename,18); #with 17, jpeg
	ppt.Quit()

#增加水印
def add_mark(imgFile, txtMark):
	img = Image.open(imgFile)
	imgWidth, imgHeight = img.size

	#http://blog.csdn.net/Dou_CO/article/details/17715919
	textImgW = int(imgWidth * 1.5)	#确定写文字图片的尺寸，要比照片大
	textImgH = int(imgHeight * 1.5)
	blank = Image.new("RGB",(textImgW,textImgH),"white")  #创建用于添加文字的空白图像
	d = ImageDraw.Draw(blank)
	d.ink = 0 + 0 * 256 + 0 * 256 * 256
		# Try multiple fonts in case simhei.ttf is not available
	try:
		# Try SimHei first (for Chinese characters)
		markFont = ImageFont.truetype('simhei.ttf', size=180)
	except OSError:
		try:
			# Fall back to Arial which is commonly available on Windows
			markFont = ImageFont.truetype('arial.ttf', size=180)
		except OSError:
			# Last resort - use default font
			markFont = ImageFont.load_default()
			print("Warning: Using default font as neither SimHei nor Arial was found")
	
	# Use textbbox for newer Pillow versions instead of getsize
	try:
		# For newer Pillow versions
		left, top, right, bottom = markFont.getbbox(txtMark)
		fontWidth, fontHeight = right - left, bottom - top
	except AttributeError:
		try:
			# For older Pillow versions
			fontWidth, fontHeight = markFont.getsize(txtMark)
		except AttributeError:
			# Fallback for very old versions or unexpected issues
			fontWidth, fontHeight = 500, 200
			print("Warning: Could not determine font size, using default values")
	d.text(((textImgW - fontWidth)/2, (textImgH - fontHeight)/2), txtMark, font=markFont)
	textRotate = blank.rotate(30)

	rLen = math.sqrt((fontWidth/2)**2+(fontHeight/2)**2)   
	oriAngle = math.atan(fontHeight/fontWidth)
	cropW = rLen*math.cos(oriAngle + math.pi/6) *4   #被截取区域的宽高
	cropH = rLen*math.sin(oriAngle + math.pi/6) *4
	box = [int((textImgW-cropW)/2-1),int((textImgH-cropH)/2-1)-50,int((textImgW+cropW)/2+1),int((textImgH+cropH)/2+1)]
	textImg = textRotate.crop(box)  #截取文字图片
	pasteW,pasteH = textImg.size
	#旋转后的文字图片粘贴在一个新的blank图像上
	textBlank = Image.new("RGB",(imgWidth,imgHeight),"white")
	pasteBox = (int((imgWidth-pasteW)/2-1),int((imgHeight-pasteH)/2-1))
	textBlank.paste(textImg,pasteBox)
	waterImage = Image.blend(img.convert('RGB'),textBlank,0.1)

	fileDir = os.path.dirname(imgFile) + '_' + txtMark
	fileName = os.path.join(fileDir, os.path.basename(imgFile))
	waterImage.save(fileName,'png')

#合并输出PDF
def topdf(path,recursion=None,pictureType=None,sizeMode=None,width=None,height=None,fit=None,save=None):
	"""
	Parameters
	----------
	path : string
		   path of the pictures
	pictureType : list
				  type of pictures,for example :jpg,png...
	sizeMode : int 
		   None or 0 for pdf's pagesize is the biggest of all the pictures
		   1 for pdf's pagesize is the min of all the pictures
		   2 for pdf's pagesize is the given value of width and height
		   to choose how to determine the size of pdf
	width : int
			width of the pdf page
	height : int
			height of the pdf page
	fit : boolean
		   None or False for fit the picture size to pagesize
		   True for keep the size of the pictures
		   wether to keep the picture size or not
	save : string 
		   path to save the pdf 
	"""

	filelist = os.listdir(path)
	filelist = [os.path.join(path, f) for f in filelist]
	filelist.sort(key=lambda x: os.path.getmtime(x))

	maxw = 0
	maxh = 0
	if sizeMode == None or sizeMode == 0:
		for i in filelist:
			#print('----'+i)
			im = Image.open(i)
			if maxw < im.size[0]:
				maxw = im.size[0]
			if maxh < im.size[1]:
				maxh = im.size[1]
	elif sizeMode == 1:
		maxw = 999999
		maxh = 999999
		for i in filelist:
			# Fixed Image2 to Image
			im = Image.open(i)
			if maxw > im.size[0]:
				maxw = im.size[0]
			if maxh > im.size[1]:
				maxh = im.size[1]
	else:
		if width == None or height == None:
			raise Exception("no width or height provid")
		maxw = width
		maxh = height

	maxsize = (maxw,maxh)
	if save == None:
		filename_pdf = os.path.join(path, path.split('\\')[-1])
	else:
		filename_pdf = os.path.join(save, path.split('\\')[-1])

	filename_pdf = filename_pdf + '.pdf'
	print('准备生成' + filename_pdf)
	c = canvas.Canvas(filename_pdf, pagesize=maxsize )
	 
	l = len(filelist)
	for i in range(l): 
		print(filelist[i])
		(w, h) =maxsize
		width, height = letter 
		if fit == True:
			c.drawImage(filelist[i] , 0,0) 
		else:
			c.drawImage(filelist[i] , 0,0,maxw,maxh) 
		c.showPage()  
	c.save()


def main():
    print("PPT2PDF Converter - Starting...")
    ppt_dir = os.getcwd()
    markTFileName = os.path.join(ppt_dir, 'watermark.txt')
    
    try:
        markTFile = open(markTFileName, encoding="utf8")
        watermarks = markTFile.readlines()
        markTFile.close()
        
        if not watermarks:
            print("Warning: watermark.txt is empty. No watermarks will be applied.")
            
        ppt_files = [fn for fn in os.listdir(ppt_dir) if fn.endswith(('.ppt','.pptx'))]
        
        if not ppt_files:
            print("Error: No PPT/PPTX files found in the current directory.")
            print("Please place your PowerPoint files in this directory: " + ppt_dir)
            return
            
        for fn in ppt_files:
            file_name = os.path.splitext(fn)[0]
            print("Processing: " + file_name)
            ppt_file = os.path.join(ppt_dir, fn)
            img_file = os.path.join(ppt_dir, file_name+'.png')

            print("Converting PPT to PNG images...")
            ppt2png(ppt_file, img_file)
            img_dir = os.path.join(ppt_dir, file_name)

            imgFileList = os.listdir(img_dir)
            imgFileList = [os.path.join(img_dir, f) for f in imgFileList]
            imgFileList.sort(key=lambda x: os.path.getmtime(x))

            for markText in watermarks:
                markText = markText.strip('\n')
                print(f"Applying watermark: '{markText}'")
                os.makedirs(img_dir + '_' + markText, exist_ok=True)
                
                for imgFile in imgFileList:
                    add_mark(imgFile, markText)
                
                print('Watermark completed')
                markDir = img_dir + '_' + markText
                topdf(path=markDir, save=ppt_dir)
                print(f"PDF generation completed: {markDir}.pdf")
                
                # Clean up temporary files
                shutil.rmtree(markDir)
            
            print(f"All processing completed for {file_name}")
            
    except FileNotFoundError:
        print("Error: watermark.txt file not found in the current directory.")
        print("Make sure the watermark.txt file is in: " + ppt_dir)
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    
    print("Process completed.")

if __name__ == "__main__":
    main()

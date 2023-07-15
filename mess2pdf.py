## mess2pdf.py version 1.1
##
## Last edited: Feb 8, 2022

import glob
import os
import re
import platform
import subprocess
import docx2pdf
import PIL
from PIL import Image
from PyPDF2 import PdfFileMerger

#nwd = "/Users/davidflenner/Downloads/Exam 1 Scratchwork Download Nov 22, 2021 803 PM"
DEBUG = False
ATWORK = True

img_types = [ 'jpeg', 'jpg', 'png' ]
heif_types = [ 'heic', 'heif' ]
ms_types = [ 'docx' ]

conv_types = img_types + ms_types + heif_types

class fileID:
    def __init__(self, path, fname):
        self.path = path
        self.fname = fname
        self.ext = self.getExt(fname)
        self.pdfname = self.getpdfname()

    def __str__(self):
        return self.fname

    def getExt(self, fname):
        if '.' in fname:
            split = fname.split('.')
            return split[-1]
        return ''

    def getpdfname(self):
        if self.ext.lower() in conv_types:
            pdfname = self.fname.replace(self.ext, 'pdf')
            return pdfname
        else:
            return None


    def getfilepath(self):
        return os.path.join(self.path, self.fname)

    def getpdfpath(self):
        return os.path.join(self.path, self.pdfname)

    def getFID(self):
        isAssignment = re.search(r"\d+-\d+", self.fname) != None
        if isAssignment:
            fparts = self.fname.split('-')
            return f'{fparts[0]}-{fparts[1]}'
        else:
            return None
        

# Change to the specified directory
#os.chdir("/Users/davidflenner/Downloads/Exam 1 Scratch Work Download Dec 1, 2021 800 PM")

# Get a list of all files in the current working directory
filelist = glob.glob('*')

# Now make a list of all files that can be converted
convlist = []
for fname in filelist:
    # Ignore any directories
    if os.path.isfile(fname):
        fileid = fileID(os.getcwd(), fname)
        if fileid.ext.lower() in conv_types:
            convlist.append(fileid)


# Phase 1: Convert all heic files to pdf

# sips utlity only runs on mac os
if platform.system() == 'Darwin':

    toConvert = []
    for fileid in convlist:
        if fileid.ext.lower() in heif_types:
            toConvert.append(fileid)

    if len(toConvert) > 0:
        print("\nThe following HEIF file(s) were located:\n")
        for fileid in toConvert:
            print(f'  ./{fileid.fname}')
        if input("\nOkay to convert these to PDF [Y\\n]? ") == 'Y':
            for fileid in toConvert:
                evt = subprocess.run(["sips", "-s", "format", "pdf", f'{fileid.fname}', "--out", f'{fileid.getpdfname()}'], 
                    stdout=subprocess.DEVNULL)
                if DEBUG == False:
                    if evt.returncode == 0:
                        os.remove(fileid.fname)

# on windows we have to use Image Magick to convert HEIF
if platform.system() == 'Windows':

    toConvert = []
    for fileid in convlist:
        if fileid.ext.lower() in heif_types:
            toConvert.append(fileid)

    if len(toConvert) > 0:
        print("\nThe following HEIF file(s) were located:\n")
        for fileid in toConvert:
            print(f'  ./{fileid.fname}')
        if input("\nOkay to convert these to PDF [Y\\n]? ") == 'Y':
            for fileid in toConvert:
                if ATWORK:
                    exefile = 'c:\\Users\\davidflenner\\bin\\magick.exe'
                else:
                    exefile = "magick"
                evt = subprocess.run([exefile, f'{fileid.fname}', f'{fileid.getpdfname()}'], 
                    stdout=subprocess.DEVNULL)
                if DEBUG == False:
                    if evt.returncode == 0:
                        os.remove(fileid.fname)

# Phase 2: Convert all image files to pdf

toConvert = []
for fileid in convlist:
    if fileid.ext.lower() in img_types:
        toConvert.append(fileid)

if len(toConvert) > 0:
    print("\nThe following images files were located:\n")
    for fileid in toConvert:
        print(f' ./{fileid.fname}')
    if input("\nOkay to convert these to PDF [Y\\n]? ") == 'Y':
        for fileid in toConvert:
            if fileid.ext.lower() in img_types:
                try:
                    image = Image.open(fileid.getfilepath())
                    width, height = image.size
                    if width > height:
                        image = image.rotate(-90, PIL.Image.NEAREST, expand=1)
                    image.convert('RGB').save(fileid.getpdfpath())
                    if DEBUG == False:
                        os.remove(fileid.fname)
                except:
                    print(f'{fileid.fname} could not be converted')


# Phase 3: Convert all docx files to pdf

toConvert = []
for fileid in convlist:
    if fileid.ext.lower() in ms_types:
        toConvert.append(fileid)

if len(toConvert) > 0:
    print("\nThe following MS Word files were located:\n")
    for fileid in toConvert:
        print(f' ./{fileid.fname}')
    if input("\nOkay to use Word to convert these to PDF [Y\\n]? ") == 'Y':
        for fileid in convlist:
            if fileid.ext.lower() in ms_types:
                try:
                    docx2pdf.convert(fileid.getfilepath())
                    print(f'{fileid.fname} converted to {fileid.pdfname}')
                    if DEBUG == False:
                        os.remove(fileid.fname)
                except:
                    print(f'{fileid.fname} could not be converted')

# Phase 4: Combine all similar files into a single pdf

filelist = glob.glob('*')
pdflist = []
for fname in filelist:
    if os.path.isfile(fname):
        fileid = fileID(os.getcwd(), fname)
        if fileid.ext.lower() == 'pdf':
            pdflist.append(fileid)

duplicatepdfs = []
# loop through all the pdf files
for i in range(0, len(pdflist)):
    targetFID = pdflist[i].getFID()
    if targetFID != None:
        duplicates = []
        for j in range(0, len(pdflist)):
            if pdflist[j].getFID() == targetFID:
                duplicates.append(j)
        if (len(duplicates) > 1) and (duplicates not in duplicatepdfs):
            duplicatepdfs.append(duplicates)

# Go through each duplicate group and join them into a single pdf

if len(duplicatepdfs) > 0:

    print('\nThe following groups of files were found:\n')
    for group in duplicatepdfs:
        for pdf in group:
            print(f'  ./{pdflist[pdf].fname}')
        print()
    
    if input('Join these groups into individual pdfs? [Y\\n] ') == 'Y':

        for group in duplicatepdfs:

            pdfs = []

            # Convert the list item number to file names
            for pdf in group:
                pdfs.append(pdflist[pdf].fname)

            try:

                # Write the group of files to a single PDF
                merger = PdfFileMerger()
                for pdf in pdfs:
                    merger.append(pdf)

                # Write the group to a single pdf file
                merger.write('group.pdf')
                merger.close()

                # Delete all the pdfs in the group
                for i in range(0, len(pdfs)):
                    if DEBUG == False:
                        os.remove(pdfs[i])

                # Rename 'group.pdf' to use the first file name in the group
                if DEBUG == False:
                    os.rename('group.pdf', pdfs[0])

            except TypeError:
                print('\nERROR: Could not write group: \n')
                for pdf in group:
                    print(f'  ./{pdflist[pdf].fname}')
            

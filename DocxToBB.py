import os
import re
import sys
import codecs
import datetime

import Tkinter, tkFileDialog
#import pyperclip # Needs to be imported late, incompatible with open tkFileDialog

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from ConfigParser import RawConfigParser
from itertools import izip

class paraStyles:
    align = False
    right = False
    justify = False

    def closeInOrder(self):
        stack = ''
        if self.justify:
            stack += ('[/j]')
            self.justify = False
        if self.right:
            stack += ('[/r]')
            self.right = False
        if self.align:
            stack += ('[/c]')
            self.align = False
        return stack

    ##deprecated
    def setStyle(self, align, right, justify): 
        stack = ''
        if not justify == self.justify:
            stack += ('[/j]')
            self.justify = False
        if not right == self.right:
            stack += ('[/r]')
            self.right = False
        if not align == self.align:
            stack += ('[/c]')
            self.align = True
        return stack


def main(file):
    document = Document(file)
    config = readConfig()
    newFileString = u''
    paraStyle = paraStyles()
    first = True
    br = config['endlinechar'].decode('string_escape')
    if eval(config['emptylineafterparagraph']):
        print br
        endline = br + br
    else:
        endline = br
    for para in document.paragraphs:
        if para.text == '':
            continue
        if first:
            first = False
            if config['preamble']:
                newFileString += config['preamble']
            newFileString += re.sub('(?<!\\\)\$', para.text, config['titleformat'])
            if eval(config['addcopyright']):
                newFileString += br + u"\u00A9"
                dateStr = datetime.datetime.now().strftime(config['copyrightdateformat']) + ' '
                newFileString += re.sub('(?<!\\\)\$', dateStr, config['copyrightauthor'])
            continue
        else:
            newFileString += endline               

        newPara = u''
        
        newPara, paraStyle = preamblePara(newPara, para, paraStyle)

        newPara = parsePara(newPara, para, paraStyle)

        for special, replace in izip(eval(config['searchfor']), eval(config['replacewith'])):
            newPara = re.sub(special, replace, newPara)
        if eval(config['prunewhitespace']):
            while newPara[-1] == ' ':
                newPara = newPara[:-1]

        newFileString += newPara
    #close all open code-fragments and add postamble
    newFileString += paraStyle.closeInOrder()
    if config['postamble']:
                newFileString += config['postamble']
    writeTxt(newFileString, file, config)

def preamblePara(newPara, para, style):
    if style.justify and not para.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
        newPara += ('[/j]') 
        style.justify = False
    if style.right and not para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
        newPara += ('[/r]') 
        style.right = False
    if style.align and not para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
        newPara += ('[/c]') 
        style.align = False
    if para.alignment == WD_ALIGN_PARAGRAPH.CENTER and not style.align:
        newPara += ('[c]') 
        style.align = True
    if para.alignment == WD_ALIGN_PARAGRAPH.RIGHT and not style.right:
        newPara += ('[r]') 
        style.right = True
    if para.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY and not style.justify:
        newPara += ('[j]') 
        style.justify = True
    return newPara, style


def parsePara(newPara, para, style):
    #Definitions
    bold = False
    italic = False
    underline = False
    strikethrough = False
    #main
    for run in para.runs:
        if strikethrough and not run.font.strike:
            newPara += ('[/s]')
            strikethrough = False
        if underline and not run.underline:
            newPara += ('[/u]')
            underline = False
        if italic and not run.italic:
            newPara += ('[/i]')
            italic = False
        if bold and not run.bold:
            newPara += ('[/b]')
            bold = False
        if run.bold and not bold:
            newPara += ('[b]')
            bold = True
        if run.italic and not italic:
            newPara += ('[i]')
            italic = True
        if run.underline and not underline:
            newPara += ('[u]')
            underline = True
        if run.font.strike and not strikethrough:
            newPara += ('[/s]')
            strikethrough = True
        newPara += (run.text)
    if strikethrough:
        newPara += ('[/u]')
    if underline:
        newPara += ('[/u]')
    if italic:
        newPara += ('[/i]')
    if bold:
        newPara += ('[/b]')
    return newPara

def getFile():
    root = Tkinter.Tk()
    file_path = tkFileDialog.askopenfilename(filetypes = (("Word/Office docx","*.docx"),("All files","*.*")))
    root.withdraw()
    root.destroy()
    return file_path

def writeTxt(txt, source, config):
    print 'Conversion Successfull'
    if eval(config['clipboard']):
        import pyperclip # hacky workaround
        pyperclip.copy(txt)
        print 'Output copied to clipboard'
    outputpath = eval(config['outputpath'])
    if outputpath:
        outputName = os.path.splitext(os.path.basename(source))[0]
        outputName = re.sub('(?<!\\\)\$', outputName, config['outputname'])
        if outputpath == 1 or outputpath == 2:
            outputDir = os.path.dirname(source)
            if outputpath ==2:
                if os.path.isfile(os.path.join(outputDir, outputName)):
                    cnt = 1
                    while os.path.isfile(os.path.join(outputDir, outputName + '(' + cnt + ')')):
                        cnt += 1
                    outputName += '(' + cnt + ')'
        else:
            raise NotImplementedError;
            outputDir = config['outputpath'] 
        with codecs.open(os.path.join(outputDir, outputName + '.txt'), "w", 'utf-8') as text_file:
            text_file.write(txt)
        print 'Output saved as ' + outputName + '.txt'
        print str(len(txt)) + ' characters'
        if eval(config['keepopen']):           
            raw_input()
        raise SystemExit
        print 'Unknown Error, please exit manually'

def readConfig():
    config = RawConfigParser()
    config.readfp(codecs.open("DocxToBB.ini", "r", "utf-8"))

    return  dict(config.items('DEFAULT'))


if __name__ == '__main__':
    file = getFile()
    main(file)
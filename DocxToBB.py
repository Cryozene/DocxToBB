import os
import re
import sys
import codecs
import datetime
import string
import random
from unidecode import unidecode
from threading import Thread
import functools

import Tkinter, tkFileDialog
#import pyperclip # Needs to be imported late, incompatible with open tkFileDialog

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from ConfigParser import RawConfigParser
from itertools import izip

import win32com.client as win32
from win32com.client import constants

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

## timeout decorator
def timeout(timeout):
    def deco(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            res = [Exception('function [%s] timeout [%s seconds] exceeded!' % (func.__name__, timeout))]
            def newFunc():
                try:
                    res[0] = func(*args, **kwargs)
                except Exception, e:
                    res[0] = e
            t = Thread(target=newFunc)
            t.daemon = True
            try:
                t.start()
                t.join(timeout)
            except Exception, je:
                print 'error starting thread'
                raise je
            ret = res[0]
            if isinstance(ret, BaseException):
                raise ret
            return ret
        return wrapper
    return deco


def main():
    file = getFile()
    filename = os.path.basename(file)
    file, cleanup = genDocx(file)
    try:
        docx = Document(file)
        config = readConfig()
        txt = parseDocx(docx, config)
        writeTxt(txt, file, filename, config)
    except Exception as e:
        raise Exception, "The code is buggy: %s" % e, sys.exc_info()[2]
    finally:
        if cleanup:
            try:
                os.remove(file)
            except OSError as e:  ## if failed, report it back to the user ##
                print ("Error: %s - %s." % (e.filename, e.strerror))
        if config['keepopen']:           
                raw_input()
        raise SystemExit
    print 'Unknown Error, please exit manually'

def parseDocx(document, config):
    newFileString = u''
    paraStyle = paraStyles()
    first = True
    paraStack = False
    br = config['endlinechar']
    if config['emptylineafterparagraph']:
        print br
        endline = br + br
    else:
        endline = br
    for para in document.paragraphs:

        #skip empty lines
        if config['skipemptylines'] and  para.text == '':
            continue

        #handle first line and special lines
        if first:
            first = False
            if config['preamble']:
                newFileString += config['preamble']
            newFileString += re.sub('(?<!\\\)\$', para.text, config['titleformat'])
            if config['addcopyright']:
                cpr = u"\u00A9"
                dateStr = datetime.datetime.now().strftime(config['copyrightdateformat']) + ' '
                cpr += re.sub('(?<!\\\)\$', dateStr, config['copyrightauthor'])
                cpr = re.sub('(?<!\\\)\$', cpr, config['copyrightstyle'])
                newFileString += br + cpr
            continue
       
        #parse paragraph
        newPara = u''       
        newPara, paraStyle = preamblePara(newPara, para, paraStyle, br)
        if config['parsecolors']:
            newPara = parseColoredPara(newPara, para, paraStyle)
        else:
            newPara = parsePara(newPara, para, paraStyle)

        #handle special replacement options
        for special, replace in izip(config['searchfor'], config['replacewith']):
            newPara = re.sub(special, replace, newPara, flags=re.UNICODE)
        if config['prunewhitespace']:
            while newPara[-1] == ' ':
                newPara = newPara[:-1]

        #handle linebreaks
        if newFileString[-1].endswith(u':'):
            newFileString += br
        elif newFileString[-1].endswith(u','):
            newFileString += br
        elif config['holdtogetherspeech']:
            line = unidecode(newPara)
            if paraStack:
                if line.startswith(',') or line.startswith('"'):
                    newFileString += br
                else:
                    newFileString += endline
            else:
                newFileString += endline
            if line.endswith('"') and len(line.split(' '))<config['holdtogetherspeech'] :
                if paraStack:
                    paraStack += newPara
                else:
                    paraStack = newPara
            else:
                paraStack = False
        else:
            newFileString += endline

        #add to output          
        newFileString += newPara

    #close all open code-fragments and add postamble
    newFileString += paraStyle.closeInOrder()
    if config['postamble']:
                newFileString += config['postamble']

    return newFileString
    

def preamblePara(newPara, para, style, br):
    if style.justify and not para.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
        newPara += ('[/j]') 
        style.justify = False
    if style.right and not para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
        newPara += ('[/r]') 
        style.right = False
    if style.align and not para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
        newPara += ('[/c]') 
        style.align = False
    #additional linebreaks needed, for enclosing environment
    if para.alignment == WD_ALIGN_PARAGRAPH.CENTER and not style.align:
        newPara += (br +'[c]') 
        style.align = True
    if para.alignment == WD_ALIGN_PARAGRAPH.RIGHT and not style.right:
        newPara += (br + '[r]') 
        style.right = True
    if para.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY and not style.justify:
        newPara += (br + '[j]') 
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
            newPara += ('[s]')
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

#copy of parsePara with added capability for parsing special colors to Text
#heavy implication for adding additional characters and parsing time
def parseColoredPara(newPara, para, style):
    #Definitions
    bold = False
    italic = False
    underline = False
    strikethrough = False
    color = '00000'
    colorchanged = False
    #main
    for run in para.runs:
        if colorchanged:
            if run.font.color.rgb:
                col = str(run.font.color.rgb)
                if not col == color:
                    newPara += ('[/color]')
                    newPara += ('[color=#' + col + ']')
                else:
                    pass # no color change
            else:
                newPara += ('[/color]')
                colorchanged = False
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
            newPara += ('[s]')
            strikethrough = True
        if not colorchanged and run.font.color.rgb:
            color = str(run.font.color.rgb)
            newPara += ('[color=#' + color + ']')
            colorchanged = True
        newPara += (run.text)
    if strikethrough:
        newPara += ('[/u]')
    if underline:
        newPara += ('[/u]')
    if italic:
        newPara += ('[/i]')
    if bold:
        newPara += ('[/b]')
    if colorchanged:
        newPara += ('[/color]')
    return newPara

def getFile():
    root = Tkinter.Tk()
    file_path = tkFileDialog.askopenfilename(filetypes = (("Word/Office",("*.docx", "*.doc")),("All files","*.*")))
    root.withdraw()
    root.destroy()
    if not file_path:
        print "No file selected, exiting ..."
        raise SystemExit
    return file_path

def writeTxt(txt, source, filename, config):
    print 'Conversion Successfull'
    if config['clipboard']:
        import pyperclip # hacky workaround
        pyperclip.copy(txt)
        print 'Output copied to clipboard'
    outputpath = config['outputpath']
    if outputpath:
        outputName = os.path.splitext(filename)[0]
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
        

def readConfig():
    config = RawConfigParser()
    config.readfp(codecs.open("DocxToBB.ini", "r", "utf-8"))
    default = dict(config.items('DEFAULT'))
    try:
        default['keepopen'] = eval(default['keepopen'])
        default['skipemptylines'] = eval(default['skipemptylines'])
        default['outputpath'] = int(eval(default['outputpath']))
        default['clipboard'] = eval(default['clipboard'])
        default['emptylineafterparagraph'] = eval(default['emptylineafterparagraph'])
        default['addcopyright'] = eval(default['addcopyright'])
        default['prunewhitespace'] = eval(default['prunewhitespace'])
        default['parsecolors'] = eval(default['parsecolors'])
        default['holdtogetherspeech'] = int(eval(default['holdtogetherspeech']))
        default['datetime'] = datetime.datetime.now().strftime(default['copyrightdateformat'])
        default['searchfor'] = eval(default['searchfor'])
        default['replacewith'] = eval(default['replacewith'])
        default['endlinechar'] = default['endlinechar'].decode('string_escape')
        default['preamble'] = replaceLinebreaks(default['endlinechar'], default['preamble'])
        default['postamble'] = replaceLinebreaks(default['endlinechar'], default['postamble'])
        default['titleformat'] = replaceLinebreaks(default['endlinechar'], default['titleformat'])
        default['copyrightauthor'] = replaceLinebreaks(default['endlinechar'], default['copyrightauthor'])
        default['copyrightstyle'] = replaceLinebreaks(default['endlinechar'], default['copyrightstyle'])
        for i in range(len(default['replacewith'])):
            default['replacewith'][i] = replaceLinebreaks(default['endlinechar'], default['replacewith'][i])   
        return  default
    except:
        print "Error while parsing config:", sys.exc_info()[0]
        raise

def replaceLinebreaks(endline, input):
    return re.sub('(?<!\\\)\[\/br\]', endline, input)

@timeout(30)
def genDocx(path):
    dir = os.path.dirname(path)
    file = os.path.basename(path)
    base, ext = os.path.splitext(path)
    if ext == ".docx":
        return path, False

    duplicateFile = True
    while duplicateFile:
        out_filename = id_generator(size=16) + ".docx"
        out_file = os.path.abspath(os.path.join(dir, out_filename))
        duplicateFile = os.path.exists(out_file)

    word = win32.Dispatch("Word.Application")
    word.visible = 0
    wb = word.Documents.Open(os.path.abspath(path))
    wb.SaveAs2(out_file, FileFormat=16) # file format for docx
    wb.Close()

    return out_file, True

def id_generator(size=16, chars=string.ascii_uppercase + string.digits):
   return ''.join(random.choice(chars) for _ in range(size))


if __name__ == '__main__':
    main()
    

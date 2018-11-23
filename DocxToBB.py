import os
import re
import sys
import traceback
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

##constants
VERSION = 0.5
NO_BREAK_SPACE = u"\u00A0"
COPYRIGHT = u"\u00A9"
PARASTACK_MAX = 6
COLOR_BLACK = '00000'

"""
Holds all information about current formatting
ALL format-Values accept None as a disabled state
True and False should be checked explicitly
"""
class paraStyles:
    bold = None
    italic = None
    underline = None
    strikethrough = None
    align = None
    right = None
    justify = None
    color = None
    ident = u''
    endline = 1
    cEndlS = 1

    def __init__(self, endline, ident, styleOptions):
        self.endline = endline
        self.cEndlS = endline
        if styleOptions['justify']:
            self.justify = False
        if styleOptions['align']:
            self.align = False
        if styleOptions['floatright']:
            self.right = False
        if styleOptions['bold']:
            self.bold = False 
        if styleOptions['underline']:
            self.underline = False  
        if styleOptions['strikethrough']:
            self.strikethrough = False
        if styleOptions['italic']:
            self.italic = False  
        if styleOptions['parsecolors']:
            self.color = False
        if ident and ident > 0:
            self.ident = u"\u00A0"*ident

    def closeInOrder(self):
        stack = ''
        if self.color:
            stack += ('[/color]')
            self.color = False
        if self.justify:
            stack += ('[/j]')
            self.justify = False
        if self.right:
            stack += ('[/r]')
            self.right = False
        if self.align:
            stack += ('[/c]')
            self.align = False
        if self.strikethrough:
            stack += ('[/s]')
            self.strikethrough = False
        if self.underline:
            stack += ('[/u]')
            self.underline = False
        if self.italic:
            stack += ('[/i]')
            self.italic = False
        if self.bold:
            stack += ('[/b]')
            self.bold = False
        return stack

    def nextLine(self):
        self.cEndlS = self.endline

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
    try:
        cleanup = False
        file = getFile()
        filename = os.path.basename(file)
        file, cleanup = genDocx(file)
        docx = Document(file)
        config, styleOptions = readConfig()
        txt = parseDocx(docx, config, styleOptions)
        writeTxt(txt, file, filename, config)
    except Exception as e:
        print 'Damn, you found another bug.'
        print "Please report the issue with the following information:"
        traceback.print_exc(file=sys.stdout)
        raw_input()
    finally:
        if cleanup:
            try:
                os.remove(file)
            except OSError as e:  ## if failed, report it back to the user ##
                print ("Error: %s - %s." % (e.filename, e.strerror))
        try:
            if config['keepopen']:           
                    raw_input()
        except:
            pass
    raise SystemExit

def parseDocx(document, config, styleOptions):
    newFileString = u''
    first = True
    paraStack = PARASTACK_MAX
    br = config['endlinechar']
    if config['emptylineafterparagraph']:
        endline = 2
    else:
        endline = 1
    paraStyle = paraStyles(endline, config['identfirstline'], styleOptions)
    for para in document.paragraphs:
        #skip empty lines
        if config['skipemptylines'] and  para.text == '':
            continue

        #handle first line and special lines
        if first:
            first = False
            if config['preamble']:
                newFileString += config['preamble']
            dateStr = datetime.datetime.now().strftime(config['copyrightdateformat'])
            firstline = re.sub(r'(?<!\\)\\cr', COPYRIGHT, config['header'], flags=re.UNICODE)
            firstline = re.sub(r'(?<!\\)\\date', dateStr, firstline, flags=re.UNICODE)
            firstline = re.sub(r'(?<!\\)\\title', para.text, firstline, flags=re.UNICODE)
            newFileString += firstline
            continue
       
        #parse paragraph
        newPara = u''       
        newPara, paraStyle = preamblePara(newPara, para, paraStyle, br)
        newPara = parsePara(newPara, para, paraStyle)

        #handle special replacement options
        for special, replace in izip(config['searchfor'], config['replacewith']):
            newPara = re.sub(special, replace, newPara, flags=re.UNICODE)
        if config['prunewhitespace']:
            while newPara and newPara[-1] == ' ':
                newPara = newPara[:-1]
        if not newPara and config['skipemptylines']: # empty line created by replacements or pruning
            continue

        #handle linebreaks
        if (        config['holdtogetherspeech']
            and not newFileString[-1].endswith(u':') 
            and not newFileString[-1].endswith(u',')
           ):
            line = unidecode(newPara)        
            if 0 < paraStack <= PARASTACK_MAX:
                if paraStack < PARASTACK_MAX and startsWithSpeech(line):
                    paraStyle.cEndlS -= 1
                if endsWithSpeech(line) and len(line.split(' '))<=config['holdtogetherspeech'] :
                    paraStack -= 1
                else:
                    paraStack = PARASTACK_MAX
            else:
                paraStack = PARASTACK_MAX

        #create linebreaks if necessary
        for _ in range(paraStyle.cEndlS):
            newFileString += br

        #reset endline
        paraStyle.nextLine()

        #add to output          
        newFileString += newPara

    #close all open code-fragments and add postamble
    newFileString += paraStyle.closeInOrder()
    if config['postamble']:
                newFileString += config['postamble']

    return newFileString
    

def preamblePara(newPara, para, style, parseColor):
    paraAlignment = getParaAlignment(para)
    
    if style.justify is True and not paraAlignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
        space, style = checkSpecialEndline(style)
        newPara += (space + '[/j]') 
        style.justify = False
    elif style.right is True and not paraAlignment == WD_ALIGN_PARAGRAPH.RIGHT:
        space, style = checkSpecialEndline(style)
        newPara += (space + '[/r]') 
        style.right = False
    elif style.align is True and not paraAlignment == WD_ALIGN_PARAGRAPH.CENTER:
        space, style = checkSpecialEndline(style)
        newPara += (space + '[/c]') 
        style.align = False

    if style.align is False and paraAlignment == WD_ALIGN_PARAGRAPH.CENTER:
        space, style = checkSpecialEndline(style)
        newPara += (space + '[c]') 
        style.align = True
    elif style.right is False and paraAlignment == WD_ALIGN_PARAGRAPH.RIGHT:
        space, style = checkSpecialEndline(style)
        newPara += (space + '[r]') 
        style.right = True
    elif style.justify is False and paraAlignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
        space, style = checkSpecialEndline(style)
        newPara += (space + '[j]') 
        style.justify = True
    newPara += style.ident
    return newPara, style

def getParaAlignment(para):
    if para.alignment:
        return para.alignment
    if para.paragraph_format.alignment:
        return para.paragraph_format.alignment
    if para.style.paragraph_format.alignment:
        return para.style.paragraph_format.alignment
    return None

#removes a linebreak and adds a special NoBreakSpace, if necessary
def checkSpecialEndline(style):
     space = u''
     if style.cEndlS > 0:
        style.cEndlS -= 1
        space = NO_BREAK_SPACE
     return space, style

#copy of parsePara with added capability for parsing special colors to Text
#heavy implication for adding additional characters and parsing time
def parsePara(newPara, para, style):
    #Definitions
    #main
    newcolor = None
    for run in para.runs:
        if not style.color is None and run.font.color.rgb:
            col = str(run.font.color.rgb)
            if col != style.color:
                if style.color:
                    newPara += ('[/color]')
                newcolor = ('[color=#' + col + ']')
                style.color = col
            else:
                pass # no color change
        elif style.color:
            newPara += ('[/color]')
            style.color = False
        if style.strikethrough is True and not run.font.strike:
            newPara += ('[/s]')
            style.strikethrough = False
        if style.underline is True and not run.underline:
            newPara += ('[/u]')
            style.underline = False
        if style.italic is True and not run.italic:
            newPara += ('[/i]')
            style.italic = False
        if style.bold is True and not run.bold:
            newPara += ('[/b]')
            style.bold = False
        if style.bold is False and run.bold:
            newPara += ('[b]')
            style.bold = True
        if style.italic is False and run.italic:
            newPara += ('[i]')
            style.italic = True
        if style.underline is False and run.underline:
            newPara += ('[u]')
            style.underline = True
        if style.strikethrough is False and run.font.strike:
            newPara += ('[s]')
            style.strikethrough = True
        if style.color is False and newcolor:
            newPara += newcolor
            newcolor = None
        newPara += (run.text)
    return newPara

def endsWithSpeech(line):
    if line.endswith('"'):
        return True
    else:
        return False

def startsWithSpeech(line):
    if (    line.startswith(',') 
        or  line.startswith('"')
        ):
        return True
    else:
        return False

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
    version = dict(config.items('Version'))
    default = dict(config.items('DEFAULT'))
    styleOptions = dict(config.items('StyleOptions'))
    try:
        version = str(version['version'])
        if version != str(VERSION):
            handleVersionError(config)

        default['keepopen'] = eval(default['keepopen'])
        default['skipemptylines'] = eval(default['skipemptylines'])
        default['outputpath'] = int(eval(default['outputpath']))
        default['clipboard'] = eval(default['clipboard'])
        default['emptylineafterparagraph'] = eval(default['emptylineafterparagraph'])
        default['prunewhitespace'] = eval(default['prunewhitespace'])
        default['holdtogetherspeech'] = int(eval(default['holdtogetherspeech']))
        default['identfirstline'] = int(eval(default['identfirstline']))
        default['datetime'] = datetime.datetime.now().strftime(default['copyrightdateformat'])
        default['searchfor'] = eval(default['searchfor'])
        default['replacewith'] = eval(default['replacewith'])
        default['endlinechar'] = default['endlinechar'].decode('string_escape')
        default['preamble'] = replaceLinebreaks(default['endlinechar'], default['preamble'])
        default['postamble'] = replaceLinebreaks(default['endlinechar'], default['postamble'])
        default['header'] = replaceLinebreaks(default['endlinechar'], default['header'])
        for i in range(len(default['replacewith'])):
            default['replacewith'][i] = replaceLinebreaks(default['endlinechar'], default['replacewith'][i])
         
        styleOptions['justify'] = eval(styleOptions['justify'])
        styleOptions['align'] = eval(styleOptions['align'])    
        styleOptions['floatright'] = eval(styleOptions['floatright'])    
        styleOptions['bold'] = eval(styleOptions['bold'])    
        styleOptions['underline'] = eval(styleOptions['underline'])    
        styleOptions['strikethrough'] = eval(styleOptions['strikethrough'])    
        styleOptions['italic'] = eval(styleOptions['italic'])    
        styleOptions['parsecolors'] = eval(styleOptions['parsecolors'])         
        return  default, styleOptions
        #additional validation
        if not default['emptylineafterparagraph']:
            default['holdtogetherspeech'] = 0
    except:
        print "Error while parsing config:", sys.exc_info()[0]
        raise

def replaceLinebreaks(endline, input):
    return re.sub(r'(?<!\\)\[\/br\]', endline, input, flags=re.UNICODE)

def handleVersionError(config):
    print 'Version of config incompatible with current Version of Parser.'
    print 'Trying to update automatically ...'
    raise NotImplementedError

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
    

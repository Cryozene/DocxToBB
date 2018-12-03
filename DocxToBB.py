import documentation

import ast
import codecs
import datetime
import itertools
import os
import pyperclip
import random
import re
import string
import sys
import tkinter as tk
import traceback
from configparser import RawConfigParser, NoSectionError
from copy import copy
from tkinter import NSEW, NW, E, N, S, W, END, filedialog, messagebox, ttk, font

import win32com.client as win32
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from unidecode import unidecode

##constants
VERSION = "0.5"
NO_BREAK_SPACE = u"\u00A0"
COPYRIGHT = u"\u00A9"
PARASTACK_MAX = 6
COLOR_BLACK = '00000'

#globals
tk_root = tk.Tk() # pre-initialization of tkinter objects needs running instance of tk.TK() present


## Exception Bases
class BadConfigException(Exception):

    def __init__(self, *args, origin="", infomsg="", **kwargs):
        super().__init__(*args, **kwargs)
        self.origin = origin
        self.infomsg = infomsg

    def getInfo(self):
        return self.origin, self.infomsg

    def setInfomsg(self, value):
        self.infomsg = value
        return self

    def setOrigin(self, value):
        self.origin = value
        return self

class DeprecatedConfigException(Exception):
    pass
      
class BBTag:
    '''
    Holds all information about converting a BB-Code Tag into a tkinter Tag
    '''
    #constants
    SPACING2_BOOK = 0
    SPACING2_SEITE = 2
    SPACING2_SERIF = 2
    SPACING2_SANS = 2
    FONT_BOOK = font.Font(family="Times", size=14)
    FONT_SEITE = font.Font(family="Helvetica", size=12)
    FONT_SERIF = font.Font(family="Times", size=12)
    FONT_SANS = font.Font(family="Helvetica", size=10)
    TEXTWIDTH_BOOK = 580
    TEXTWIDTH_SEITE = 700
    BGCOLOR_BOOK = '#F9F9F9'
    BGCOLOR_SEITE = '#F8F3ED'
    BOOK = 0
    SEITE = 1
    SERIF = 2
    SANS = 3

    def __init__(self, family=BOOK, bold=False, italic=False, underline=False, strikethrough=False, justify=tk.LEFT, color='#000000', name='default', size=False):
        self.setTag(family=family, bold=bold, italic=italic, underline=underline, strikethrough=strikethrough, justify=justify, color=color, name=name, size=size)

    def setTag(self, family=BOOK, bold=False, italic=False, underline=False, strikethrough=False, justify=tk.LEFT, color='#000000', name='default', size = False):
        self.setFamily(family)
        self.setSize(size)
        self.setBold(bold)
        self.setItalic(italic)
        self.setUnderline(underline)
        self.setStrikethrough(strikethrough)
        self.setJustify(justify)
        self.setTextColor(color)
        self.name = name

    def setJustify(self, justify):
        self.justify = justify

    def setFamily(self, family):
        self.family = family
        if family == self.BOOK:
            self.font = self.FONT_BOOK.copy()
            self.spacing = self.SPACING2_BOOK
            self.size = 14
        elif family == self.SEITE:
            self.font = self.FONT_SEITE.copy()
            self.spacing = self.SPACING2_SEITE
            self.size = 12
        elif family == self.SERIF:
            self.font = self.FONT_SERIF.copy()
            self.spacing = self.SPACING2_SERIF
            self.size = 12
        elif family == self.SANS:
            self.font = self.FONT_SANS.copy()
            self.spacing = self.SPACING2_SANS
            self.size = 10
        else:
            raise ValueError

    def setBold(self, enable):
        if enable:
            self.font.configure(weight=font.BOLD)
            self.bold = True
        else:
            self.font.configure(weight=font.NORMAL)
            self.bold = False

    def setItalic(self, enable):
        if enable:
            self.font.configure(slant=font.ITALIC)
            self.italic = True
        else:
            self.font.configure(slant=font.ROMAN)
            self.italic = False

    def setUnderline(self, enable):
        if enable:
            self.font.configure(underline=1)
            self.underline = True
        else:
            self.font.configure(underline=0)
            self.underline = False

    def setStrikethrough(self, enable):
        if enable:
            self.font.configure(overstrike=1)
            self.strikethrough = True
        else:
            self.font.configure(overstrike=0)
            self.strikethrough = False
    
    def setSize(self, size):
        if not size: return
        self.size = size
        self.font.configure(size=size)

    def setTextColor(self, color):
        self.color = color

    def setName(self, name):
        self.name = name

    def getFont(self):
        return self.font

    def getTextcolor(self):
        return self.color

    def getSpacing(self):
        return self.spacing

    def getJustify(self):
        return self.justify

    def getName(self):
        return self.name

    def configureTag(self, textbox):
        textbox.tag_configure(self.name, spacing2=self.spacing, font=self.font, foreground=self.color, justify=self.justify)

    def __eq__(self, other):
        return (isinstance(other, (BBTag, ))
                and all([self.getSpacing() == other.getSpacing(),
                        self.getTextcolor() == other.getTextcolor(),
                        self.getJustify() == other.getJustify(),
                        self.size == other.size,
                        self.italic == other.italic,
                        self.bold == other.bold,
                        self.underline == other.underline,
                        self.strikethrough == other.strikethrough,
                        self.family == other.family]))
    def __copy__(self):
        new = type(self)(family=self.family, 
                        bold=self.bold, italic=self.italic, underline=self.underline, strikethrough=self.strikethrough,
                        justify=self.justify, color=self.color, name=self.name)
        return new
            
class BBTagStack:   
    bold = 0
    italic = 0
    underline = 0
    strikethrough = 0
    justify = []
    size = []
    family = []
    color = []

class BBToTkText:
    ''' has all method to parse a Text with BBCode to the Preview area
        style in line with display at romane-forum.de
    '''
    def __init__(self):
        pass

    def parse(self, textbox, frame, s='[seite] [/seite]'):
        self.textbox = textbox
        self.clearTextbox()
        self.currentTag = BBTag()
        self.tagcnt = itertools.count()
        if re.search(r'(?<!\\)\[book\]', s):
            frame.configure(width=BBTag.TEXTWIDTH_BOOK)
            self.closing = '[/book]'
            self.currentTag.setFamily(BBTag.BOOK)
        else:
            frame.configure(width=BBTag.TEXTWIDTH_SEITE)
            self.closing = '[/seite]'
            self.currentTag.setFamily(BBTag.SEITE)
        
        sList = re.split(r'((?<!\\)\[.*?\])', s) # matches any bbCodeTag
        self.parseSList(sList)

    def parseSList(self, sList):
        self.textbuffer = ''
        self.tagStack = BBTagStack()
        for s in sList:
            s = re.split(r'(\[.*?(?=\[))', s) # filter out preceding '['
            for paraTag in s:
                if paraTag == '':
                    continue
                if paraTag.endswith(']') and paraTag.startswith('['): # command
                    self.parseText()
                    self.parseBBTag(paraTag)
                else:
                    self.textbuffer += paraTag
        self.parseText()
        self.insertText('\n\n\n\n')
            
    def parseText(self):
        if not self.textbuffer:
            return
        text = self.textbuffer.splitlines(True)
        self.textbuffer = ''
        self.insertText(text)

    def insertText(self, txt):
        linebuffer = ''
        for line in txt:
            line = re.sub(' +', ' ', line, flags=re.UNICODE)#strip double whitespace
            line = re.sub('^ *', '', line, flags=re.UNICODE)#strip leading whitespace
            line = re.sub(' *$', '', line, flags=re.UNICODE)#strip trailing whitepace
            line = re.sub(r'\\\[', '[', line, flags=re.UNICODE)#substitute escaped [ whitepace
            linebuffer += line
        tagname = 'Tag' + str(self.tagcnt.__next__())
        t = BBTag(family=self.currentTag.family, 
                bold=self.currentTag.bold, italic=self.currentTag.italic, underline=self.currentTag.underline, strikethrough=self.currentTag.strikethrough, 
                justify=self.currentTag.justify, color=self.currentTag.color, name=tagname, size=self.currentTag.size)
        t.configureTag(self.textbox)
        self.textbox.insert(tk.END, linebuffer, t.name)

    def stripLastLinebreak(self):
        self.textbuffer = re.sub(r'[\r]?\n *$', '', self.textbuffer)

    def parseBBTag(self, paraCommand):
        lpc = paraCommand.lower()
        if lpc == '[b]':
            self.tagStack.bold += 1
            self.currentTag.setBold(True)
        elif lpc == '[i]':
            self.tagStack.italic += 1
            self.currentTag.setItalic(True)
        elif lpc == '[u]':
            self.tagStack.underline += 1
            self.currentTag.setUnderline(True)      
        elif lpc == '[s]':
            self.tagStack.strikethrough += 1
            self.currentTag.setStrikethrough(True)
        elif lpc == '[c]':
            self.tagStack.justify.append(self.currentTag.getJustify())
            self.currentTag.setJustify(tk.CENTER)
            self.stripLastLinebreak()
        elif lpc == '[r]':
            self.tagStack.justify.append(self.currentTag.getJustify())
            self.currentTag.setJustify(tk.RIGHT)
            self.stripLastLinebreak()
        elif lpc == '[j]':
            self.stripLastLinebreak()
        elif lpc == '[serif]':
            self.tagStack.family.append(self.currentTag.family)
            self.currentTag.setFamily(BBTag.SERIF)
        elif lpc == '[sans]':
            self.tagStack.family.append(self.currentTag.family)
            self.currentTag.setFamily(BBTag.SANS)
        elif lpc == '[/b]':
            if self.tagStack.bold:
                self.tagStack.bold -= 1
                if not self.tagStack.bold:
                    self.currentTag.setBold(False)
        elif lpc == '[/u]':
            if self.tagStack.underline:
                self.tagStack.underline -= 1
                if not self.tagStack.underline:
                    self.currentTag.setUnderline(False)
        elif lpc == '[/i]':
            if self.tagStack.italic:
                self.tagStack.italic -= 1
                if not self.tagStack.italic:
                    self.currentTag.setItalic(False)
        elif lpc == '[/s]':
            if self.tagStack.strikethrough:
                self.tagStack.strikethrough -= 1
                if not self.tagStack.strikethrough:
                    self.currentTag.setStrikethrough(False)
        elif lpc == '[/c]':
            if self.currentTag.getJustify() == tk.CENTER:
                self.currentTag.setJustify(self.tagStack.justify.pop())
                self.textbuffer += '\n'
        elif lpc == '[/r]':
            if self.currentTag.getJustify() == tk.RIGHT:
                self.currentTag.setJustify(self.tagStack.justify.pop())
                self.textbuffer += '\n'
        elif lpc == '[/j]':
            self.textbuffer += '\n'
        elif lpc == '[/sans]':
            if self.currentTag.family == BBTag.SANS:
                self.currentTag.setFamily(self.tagStack.family.pop())
        elif lpc == '[/serif]':
            if self.currentTag.family == BBTag.SERIF:
                self.currentTag.setFamily(self.tagStack.family.pop())
        elif re.match(r'\[color=#[0-9a-f]{6}\]', lpc):
            color = re.search(r'#[0-9a-f]{6}',lpc).group()
            self.tagStack.color.append(self.currentTag.getTextcolor())
            self.currentTag.setTextColor(color)
        elif lpc == '[/color]':
            if self.tagStack.color:
                self.currentTag.setTextColor(self.tagStack.color.pop())
        elif re.match(r'\[size=\d{2,3}\]', lpc):
            self.tagStack.size.append(self.currentTag.size)
            size = re.search(r'\d{2,3}', lpc).group()
            size = int(self.tagStack.size[0]*int(size)/100)
            self.currentTag.setSize(size)
        elif lpc == '[/size]':
            if self.tagStack.size:
                self.currentTag.setSize(self.tagStack.size.pop())
        elif lpc in ['[book]', '[/book]', '[seite]', '[/seite]']:
            pass
        else:
            self.textbuffer += paraCommand
            return

    def clearTextbox(self):
        self.textbox.delete('1.0', END)
        for tag in self.textbox.tag_names():
            self.textbox.tag_delete(tag)
                  
class InputDialog:

    def __init__(self, parent, text='Enter String', ok='Ok', cancel='Cancel'):
        self.top = tk.Toplevel(parent)
        self.input_Label = tk.Label(self.top, text=text)
        self.input_Label.pack()
        self.input_EntryBox = tk.Entry(self.top, width=50)
        self.input_EntryBox.pack()
        self.ok_Frame = ttk.Frame(self.top)
        self.ok_Frame.pack()
        self.ok_Button = tk.Button(self.ok_Frame, text=ok, command=self.send)
        self.ok_Button.pack(anchor=(W,), side=tk.LEFT)
        self.okSpace_Label = tk.Label(self.ok_Frame, width = 30)
        self.okSpace_Label.pack(side=tk.LEFT)
        self.cancel_Button = tk.Button(self.ok_Frame, text=cancel, command=lambda *args: self.top .destroy())
        self.cancel_Button.pack(anchor=(E,), side=tk.LEFT)
        self.entry = ''

    def send(self):
        self.entry = self.input_EntryBox.get()
        self.top.destroy()

    def getEntry(self):
        return self.entry

class ConfigValidator:
    '''
    Holds Validation functions for settings and general config
    '''
    def __init__(self):
        self.configdefault = {
            'preamble'                  : (self.isBB, '[book]'),
            'postamble'                 : (self.isBB, '[/book]'),
            'emptylineafterparagraph'   : (self.isBool, True),
            'skipemptylines'            : (self.isBool, True),
            'holdtogetherspeech'        : (self.isPositiveInt, 0),
            'endlinechar'               : (self.isEndline, '\r\n'),
            'prunewhitespace'           : (self.isBool, True),
            'header'                    : (self.isBB, '[c][b][size=150]$title[/size][/b] [/br] $cr$date by [b]Name[/b] [/c]'),
            'copyrightdateformat'       : (self.isDate, '%m/%Y'),
            'identfirstline'            : (self.isPositiveInt, 0),
            'outputpath'                : (lambda value: self.isIn(value, [0, 1, 2]), 1),
            'clipboard'                 : (self.isBool, True),
            'searchandreplace'          : (self.isSRTuple, [[2, 'Fix [book] Ellipsis', "\u2026", "..."],
                                                            [0, 'Fix [seite] Ellipsis', "(?<!\.)\.\.\.(?!\.)", "\u2026"],
                                                            [1, 'Space Before TriplePoints', "(?<![\. ])\.\.\.(?!\.)", "\u00A0..."],
                                                            [0, 'Space Before Ellipsis', "(?<![\. ])\u2026", "\u00A0\u2026"],
                                                            [0, 'Inherit Double Spaces', "  ", " \u00A0"]]),
            'keepopen'                  : (self.isBool, False),
        }
        self.style_enabled = {
            'justify'                   : (self.isBool, True),
            'align'                     : (self.isBool, True),
            'floatright'                : (self.isBool, True),
            'bold'                      : (self.isBool, True),
            'underline'                 : (self.isBool, True),
            'strikethrough'             : (self.isBool, True),
            'italic'                    : (self.isBool, True),
            'colors'                    : (self.isBool, False),
        }
        self.version = {'version' : ((lambda x: x), VERSION)}
        self.configs = {
            'Version'       : self.version,
            'Settings'       : self.configdefault,
            'StyleOptions'  : self.style_enabled,
        }
        self.filePath = os.path.abspath('DocxToBB.ini')
        self.labelLookup = {}

    def addConfigLabel(self, name, labelVar):
        name = name.lower()
        self.labelLookup[name] = labelVar

    def getFilePath(self):
        return self.filePath

    def getDefault(self, name):
        default = self.getConfigDict(name)[name][1]
        if name in self.labelLookup:
            print(default)
            self.labelLookup[name].set(parsed)
        return default

    def parseValue(self, name, value):
        parsed = self.getConfigDict(name)[name][0](value)
        return parsed

    def getConfigDict(self, name):
        dct = {}
        if name.lower() in self.configdefault:
            return self.configdefault
        if name.lower() in self.style_enabled:
            return self.style_enabled
        if name.lower() in self.version:
            return self.version
        return ValueError

    def parseConfig(self, configlist):
        configs = set(self.configdefault.keys())
        configs.update(self.style_enabled.keys())
        for i in range(len(configlist)):
            for key, value in configlist[i].items():
                try:
                    configs.remove(key)
                except KeyError: continue
                try:
                    configlist[i][key] = self.parseValue(key, value)
                except BadConfigException as e:
                    raise e.setOrigin(key)
        if configs:
            raise DeprecatedConfigException()
        return tuple(configlist)

    def isBB(self, value):
        try:
            value = str(value).encode().decode('unicode-escape')
            #TODO validate BB
        except Exception as e:
            raise BadConfigException(infomsg="Can't parse string")
        return value

    def isBool(self, value):
        try:
            if isinstance(value, (str,)):
                if value.lower() == 'true' or value == '0':
                    value = True
                elif  value.lower() == 'false' or value == '1':
                    value = False
                else:
                    raise BadConfigException(infomsg="Can't parse Value, must be \"True\" or \"False\"")
            if isinstance(value, (bool,)):
                return value
            else:
                raise BadConfigException(infomsg="Must be True or False")
        except:
            raise BadConfigException(infomsg="Must be Boolean Value")

    def isPositiveInt(self, value):
        try:
            if isinstance(value, (str,)):
                value = int(ast.literal_eval(value))
            if isinstance(value, (int,)) and value >= 0:
                return value
            else:
                raise BadConfigException(infomsg="Must be Integer equal or greater than 0")
        except:
            raise BadConfigException(infomsg="Must be Integer Value")

    def isEndline(self, value):
        try:
            value = str(value)
            if value in ['\n', '\r\n', '\r']:
                return value
            else:
                raise BadConfigException(infomsg="Must be valid linebreak")
        except:
            raise BadConfigException(infomsg="Can't parse endline")

    def isDate(self, value):
        try:
            value = str(value)
            datetime.datetime.now().strftime(value)
        except:
            raise BadConfigException(infomsg="Can't parse date string")
        return value

    def isSRTuple(self, value):
        value = self.isList(value)
        for i, sr in enumerate(value):
            try:
                enabled, name, search, replace = sr
            except:
                raise BadConfigException(infomsg="Ill-formed Search&Replace")
            try:
                enabled = self.isPositiveInt(enabled)
            except BadConfigException as e:
                raise e.setInfomsg('Enable_State of search & replace tuple must be a Boolean')
            try:
                name = str(name)
            except:
                raise BadConfigException(infomsg='Name of search & replace tuple must be a String')
            try:
                search = self.isRegex(search)
            except BadConfigException as e:
                raise e.setInfomsg('Search of search & replace tuple must be a valid Regex-Expression')
            try:
                replace = str(replace)
            except:
                raise BadConfigException(infomsg='Replace of search & replace tuple must be a String')
            value[i] = [enabled, name, search, replace]
        return value

    def isRegex(self, value):
        try:
            value = re.compile(value).pattern
        except re.error:
            raise BadConfigException(infomsg="Not a valid Regex Rexpression")
        return value

    def isList(self, value):
        try:
            if isinstance(value, (str,)):
                value = ast.literal_eval(value)
            if isinstance(value, (list, tuple)):
                return value
        except:
            pass
        raise BadConfigException(infomsg="Must be a list ([value1, value2, ...])")

    def isIn(self, value, lst):
        try:
            if isinstance(value, (str,)):
                value = eval(value)
            if value in lst:
                return value
            else:
                msg = "Value must be in" + str(lst)
                raise BadConfigException(infomsg=msg)
        except:
            raise BadConfigException(infomsg="Can't parse value")

    def saveConfig(self, default, style_enabled):
        parser = RawConfigParser()
        cnt = len(default['searchandreplace'])
        for i, sr in reversed(list(enumerate(default['searchandreplace']))): #normalize Importance
            if sr[0] > 0:
                default['searchandreplace'][i][0] = cnt
            cnt -= 1
        parser['Settings'] = {}
        for key, value in default.items():
            if key == 'endlinechar': continue
            parser.set('Settings', str(key), str(value))
        parser['StyleOptions'] = {}
        for key, value in style_enabled.items():
            parser.set('StyleOptions', str(key), str(value))
        parser['Version'] = {'version' : VERSION}
        self.writeConfigfile(parser)


    def generateNewConfig(self):
        parser = RawConfigParser()
        for section, sectiondict in self.configs.items():
            parser[section] = {}
            for key, value in sectiondict.items():
                if key == 'endlinechar': continue
                parser.set(section, str(key), str(self.getDefault(key)))
                print(section, str(key), str(self.getDefault(key)))
        self.writeConfigfile(parser)

    def writeConfigfile(self, parser):
        with open(self.filePath, mode="w", encoding='utf-8') as configfile:
            parser.write(configfile)

class ParaStyles:
    """
    Holds all information about current formatting
    ALL format-Values accept None as a disabled state
    True and False should be checked explicitly
    """
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
        if styleOptions['colors']:
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

class CreateToolTip:
    '''
    create a tooltip for a given widget
    '''
    def __init__(self, widget, text='widget info'):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.close)

    def enter(self, event=None):
        x = y = 0
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # creates a toplevel window
        self.tw = tk.Toplevel(self.widget)
        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(self.tw, text=self.text, justify='left',
                       background='white', relief='solid', borderwidth=1,
                       font=("times", "12", "normal"))
        label.pack(ipadx=1)
    def close(self, event=None):
        if self.tw:
            self.tw.destroy()

class TextToBB(object):

    def __init__(self):
        #variables needed for initialization
        self.parsedTXT = ''
        self.init = True

        #root
        self.root = tk_root
        self.root.title("TextToBB")
        self.root.option_add('*tearOff', tk.FALSE)
        self.root.minsize(600, 400)

        #create mainframe
        self.mainframe = ttk.Frame(self.root, padding="3 3 12 12")
        self.mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        try:
            self.configValidator = ConfigValidator()
            self.config, self.style_enabled = readConfig()
        except Exception as e:
            print ('couldn\'t read Config.ini')
            raise

        #draw and start GUI
        self.drawGUI()
        self.init = False
        self.startGUI()

    def startGUI(self):
        for child in self.mainframe.winfo_children():
            child.grid_configure(padx=5, pady=5)
        self.root.mainloop()

    def drawGUI(self):
        self.status = tk.StringVar()
        self.status.set('')
        self.col1Label = ttk.Label(self.mainframe, textvariable=self.status, font=("Helvetica", "18", "bold"))
        self.col1Label.grid(column=0, columnspan=3, row=0)

        self.filePath = tk.StringVar()
        self.selectFile_Button = ttk.Button(self.mainframe, text="Select File", command=self.getFile)
        self.selectFile_Button.grid(column=0, row=2, sticky=(W))
        self.filePath_Field = ttk.Entry(self.mainframe, textvariable=self.filePath, width=120)
        self.filePath_Field.grid(column=0, row=2, columnspan=3, sticky=(E))

        self.convert_Button = ttk.Button(self.mainframe, text="Convert", command=self.convert)
        self.convert_Button.grid(column=1, row=3, sticky=S)

        self.showPreview = tk.StringVar()
        self.showPreview.set('Disable Preview')

        self.drawStyles()
        self.drawSettings()
        self.drawSearchAndReplace()
        self.drawExtraInfo()
        
        self.warnings = tk.StringVar()
        self.warnings.set('')
        self.warnings_Stack = []
        self.warning_Label = ttk.Label(self.mainframe, textvariable=self.warnings, width=10)
        self.warning_Label.grid(column=1, row=10, columnspan=3, sticky=(S))
        self.warning_ttp = CreateToolTip(self.warning_Label, text='Warning')

    def drawStyles(self):
        #MainFrame
        self.styles_Frame = ttk.Labelframe(self.mainframe, text='Enabled Formatting', width=300, height=30)
        self.styles_Frame.grid(column=0, row=4, columnspan=3, sticky=(NW))

        #Label
        self.justify = tk.BooleanVar()
        self.justify.set(self.style_enabled['justify'])
        self.justify.trace_add('write', lambda *args: self.changeSettings(self.justify, 'justify'))
        self.justify_Check = ttk.Checkbutton(self.styles_Frame, text='justify', variable=self.justify, onvalue=True, offvalue=False, width=30)
        self.justify_Check.grid(column=0, row=2, sticky=(W))
        self.justify_ttp = CreateToolTip(self.justify_Check, text=("Inherit justify from source?\n"
                                                                   "Discouraged when Paragraph ident > 0"))

        self.floatright = tk.BooleanVar()
        self.floatright.set(self.style_enabled['floatright'])
        self.floatright.trace_add('write', lambda *args: self.changeSettings(self.floatright, 'floatright'))
        self.floatright_Check = ttk.Checkbutton(self.styles_Frame, text='floatright', variable=self.floatright, onvalue=True, offvalue=False, width=30)
        self.floatright_Check.grid(column=1, row=2, sticky=(W))
        self.floatright_ttp = CreateToolTip(self.floatright_Check, text="Inherit floatright from source?")

        self.align = tk.BooleanVar()
        self.align.set(self.style_enabled['align'])
        self.align.trace_add('write', lambda *args: self.changeSettings(self.align, 'align'))
        self.align_Check = ttk.Checkbutton(self.styles_Frame, text='align', variable=self.align, onvalue=True, offvalue=False, width=30)
        self.align_Check.grid(column=2, row=2, sticky=(W))
        self.align_ttp = CreateToolTip(self.align_Check, text="Inherit align from source?")

        self.colors = tk.BooleanVar()
        self.colors.set(self.style_enabled['colors'])
        self.colors.trace_add('write', lambda *args: self.changeSettings(self.colors, 'colors'))
        self.colors_Check = ttk.Checkbutton(self.styles_Frame, text='colors', variable=self.colors, onvalue=True, offvalue=False, width=30)
        self.colors_Check.grid(column=3, row=2, sticky=(W))
        self.colors_ttp = CreateToolTip(self.colors_Check, text=("Inherit font-colors from source?\n"
                                                                 "Discouraged when not actively in use\n"
                                                                 "Colors may appear differently in final result"))
                                                            
        self.italic = tk.BooleanVar()
        self.italic.set(self.style_enabled['italic'])
        self.italic.trace_add('write', lambda *args: self.changeSettings(self.italic, 'italic'))
        self.italic_Check = ttk.Checkbutton(self.styles_Frame, text='italic', variable=self.italic, onvalue=True, offvalue=False, width=30)
        self.italic_Check.grid(column=0, row=3, sticky=(W))
        self.italic_ttp = CreateToolTip(self.italic_Check, text="Inherit italic from source?")

        self.bold = tk.BooleanVar()
        self.bold.set(self.style_enabled['bold'])
        self.bold.trace_add('write', lambda *args: self.changeSettings(self.bold, 'bold'))
        self.bold_Check = ttk.Checkbutton(self.styles_Frame, text='bold', variable=self.bold, onvalue=True, offvalue=False, width=30)
        self.bold_Check.grid(column=1, row=3, sticky=(W))
        self.bold_ttp = CreateToolTip(self.bold_Check, text="Inherit bold from source?")

        self.underline = tk.BooleanVar()
        self.underline.set(self.style_enabled['underline'])
        self.underline.trace_add('write', lambda *args: self.changeSettings(self.underline, 'underline'))
        self.underline_Check = ttk.Checkbutton(self.styles_Frame, text='underline', variable=self.underline, onvalue=True, offvalue=False, width=30)
        self.underline_Check.grid(column=2, row=3, sticky=(W))
        self.underline_ttp = CreateToolTip(self.underline_Check, text="Inherit underline from source?")

        self.strikethrough = tk.BooleanVar()
        self.strikethrough.set(self.style_enabled['strikethrough'])
        self.strikethrough.trace_add('write', lambda *args: self.changeSettings(self.strikethrough, 'strikethrough'))
        self.strikethrough_Check = ttk.Checkbutton(self.styles_Frame, text='strikethrough', variable=self.strikethrough, onvalue=True, offvalue=False, width=30)
        self.strikethrough_Check.grid(column=3, row=3, sticky=(W))
        self.strikethrough_ttp = CreateToolTip(self.strikethrough_Check, text="Inherit strikethrough from source?")

    def drawSettings(self):
        #MainFrame
        self.settings_Frame = ttk.Labelframe(self.mainframe, text='Settings', width=300, height=100)
        self.settings_Frame.grid(column=0, row=5, columnspan=3, sticky=(NW))

        
        #Label
        self.preamble = tk.StringVar()
        self.preamble.set(self.config['preamble'])
        self.configValidator.addConfigLabel('preamble', self.preamble)
        self.preamble.trace_add('write', lambda *args: self.changeSettings(self.preamble, 'preamble'))
        self.preambleInfo_Button = ttk.Button(self.settings_Frame, text="i", command=lambda: self.ShowInfoText('preamble'), width=3)
        self.preambleInfo_Button.grid(column=1, row=1, sticky=(W))
        self.preambleInfo_ttp = CreateToolTip(self.preambleInfo_Button, text='Extra Info')
        self.preamble_Label = ttk.Label(self.settings_Frame, text='Preamble', width=12)
        self.preamble_Label.grid(column=2, row=1, sticky=(W))
        self.preamble_ttp = CreateToolTip(self.preamble_Label, text=("Preamble of the whole File\n"
                                                                     "Defines general styles in BBCode applied to the whole text"))
        self.preamble_Field = ttk.Entry(self.settings_Frame, textvariable=self.preamble, width=40)
        self.preamble_Field.grid(column=3, row=1, columnspan=2, sticky=(W))

        ttk.Label(self.settings_Frame, width=5).grid(column=5, row=1, sticky=(NSEW))

        self.postamble = tk.StringVar()
        self.postamble.set(self.config['postamble'])
        self.configValidator.addConfigLabel('postamble', self.preamble)
        self.postamble.trace_add('write', lambda *args: self.changeSettings(self.postamble, 'postamble'))
        self.postambleInfo_Button = ttk.Button(self.settings_Frame, text="i", command=lambda: self.ShowInfoText('postamble'), width=3)
        self.postambleInfo_Button.grid(column=6, row=1, sticky=(E))
        self.postambleInfo_ttp = CreateToolTip(self.postambleInfo_Button, text='Extra Info')
        self.postamble_Label = ttk.Label(self.settings_Frame, text='Postamble', width=12)
        self.postamble_Label.grid(column=7, row=1, sticky=(W))
        self.postamble_ttp = CreateToolTip(self.postamble_Label, text=("Postamble of the whole File\n"
                                                                       "Any open style-definitions in preamble should be closed here"))
        self.postamble_Field = ttk.Entry(self.settings_Frame, textvariable=self.postamble, width=40)
        self.postamble_Field.grid(column=8, row=1, columnspan=2, sticky=(E))

        self.header = tk.StringVar()
        self.header.set(self.config['header'])
        self.configValidator.addConfigLabel('header', self.preamble)
        self.header.trace_add('write', lambda *args: self.changeSettings(self.header, 'header'))
        self.headerInfo_Button = ttk.Button(self.settings_Frame, text="i", command=lambda: self.ShowInfoText('header'), width=3)
        self.headerInfo_Button.grid(column=1, row=2, sticky=(W))
        self.headerInfo_ttp = CreateToolTip(self.headerInfo_Button, text='Extra Info')
        self.header_Label = ttk.Label(self.settings_Frame, text='Header', width=12)
        self.header_Label.grid(column=2, row=2, sticky=(W))
        self.header_ttp = CreateToolTip(self.header_Label, text=("Header defines formatting of title and coypright-mark\n"
                                                                "Refer to  extra-info for Formatting-Guidelines"))
        self.header_Field = ttk.Entry(self.settings_Frame, textvariable=self.header, width=117)
        self.header_Field.grid(column=3, row=2, columnspan=7, sticky=(W))

        self.copyrightDateFormat = tk.StringVar()
        self.copyrightDateFormat.set(self.config['copyrightdateformat'])
        self.copyrightDateFormat.trace_add('write', lambda *args: self.changeSettings(self.copyrightDateFormat, 'copyrightDateFormat'))
        self.copyrightDateFormatInfo_Button = ttk.Button(self.settings_Frame, text="i", command=lambda: self.ShowInfoText('copyrightDateFormat'), width=3)
        self.copyrightDateFormatInfo_Button.grid(column=1, row=3, sticky=(W))
        self.copyrightDateFormatInfo_ttp = CreateToolTip(self.copyrightDateFormatInfo_Button, text='Extra info')
        self.copyrightDateFormat_Label = ttk.Label(self.settings_Frame, text='Date Format', width=12)
        self.copyrightDateFormat_Label.grid(column=2, row=3, sticky=(W))
        self.copyrightDateFormat_ttp = CreateToolTip(self.copyrightDateFormat_Label, text=("Defines date format for copyright mark\n"
                                                                                            "Refer to extra-info for formatting guidelines"))
        self.copyrightDateFormat_Field = ttk.Entry(self.settings_Frame, textvariable=self.copyrightDateFormat, width=50)
        self.copyrightDateFormat_Field.grid(column=3, row=3, columnspan=2, sticky=(W))

        self.emptyLineAfterParagraph = tk.BooleanVar()
        self.emptyLineAfterParagraph.set(self.config['emptylineafterparagraph'])
        self.emptyLineAfterParagraph.trace_add('write', lambda *args: self.changeSettings(self.emptyLineAfterParagraph, 'emptyLineAfterParagraph'))
        self.emptyLineAfterParagraphInfo_Button = ttk.Button(self.settings_Frame, text="i", command=lambda: self.ShowInfoText('emptyLineAfterParagraph'), width=3)
        self.emptyLineAfterParagraphInfo_Button.grid(column=1, row=4, sticky=(W))
        self.emptyLineAfterParagraphInfo_ttp = CreateToolTip(self.emptyLineAfterParagraphInfo_Button, text='Extra info')
        self.emptyLineAfterParagraph_Check = ttk.Checkbutton(self.settings_Frame, text='Add empty lines between Paragraphs',
                                                             variable=self.emptyLineAfterParagraph, onvalue=True, offvalue=False, width=50)
        self.emptyLineAfterParagraph_Check.grid(column=2, row=4, columnspan=3, sticky=(W))
        self.emptyLineAfterParagraph_ttp = CreateToolTip(self.emptyLineAfterParagraph_Check, text="Controls if empty lines will be automatically inserted between paragraphs")

        self.skipEmptyLines = tk.BooleanVar()
        self.skipEmptyLines.set(self.config['skipemptylines'])
        self.skipEmptyLines.trace_add('write', lambda *args: self.changeSettings(self.skipEmptyLines, 'skipEmptyLines'))
        self.skipEmptyLinesInfo_Button = ttk.Button(self.settings_Frame, text="i", command=lambda: self.ShowInfoText('skipEmptyLines'), width=3)
        self.skipEmptyLinesInfo_Button.grid(column=6, row=4, sticky=(E))
        self.skipEmptyLinesInfo_ttp = CreateToolTip(self.skipEmptyLinesInfo_Button, text='Extra Info')
        self.skipEmptyLines_Check = ttk.Checkbutton(self.settings_Frame, text='Skip empty lines',
                                                            variable=self.skipEmptyLines, onvalue=True, offvalue=False, width=50)
        self.skipEmptyLines_Check.grid(column=7, row=4, columnspan=4, sticky=(W))
        self.skipEmptyLines_ttp = CreateToolTip(self.skipEmptyLines_Check, text=("Controls if empty lines in the original text should be ignored\n"
                                                                                "Enable this option, if linebreaks should be managed by TextToBB alone"))

        self.holdTogetherSpeech = tk.StringVar()
        self.holdTogetherSpeech.set(self.config['holdtogetherspeech'])
        self.configValidator.addConfigLabel('holdtogetherspeech', self.preamble)
        self.holdTogetherSpeech.trace_add('write', lambda *args: self.changeSettings(self.holdTogetherSpeech, 'holdTogetherSpeech'))
        self.holdTogetherSpeech_Button = ttk.Button(self.settings_Frame, text="i", command=lambda: self.ShowInfoText('holdTogetherSpeech'), width=3)
        self.holdTogetherSpeech_Button.grid(column=1, row=5, sticky=(W))
        self.holdTogetherSpeech_ttp = CreateToolTip(self.holdTogetherSpeech_Button, text='Extra Info')
        self.holdTogetherSpeech_Label = ttk.Label(self.settings_Frame, text='Hold direct speech together', width=30)
        self.holdTogetherSpeech_Label.grid(column=2, row=5, columnspan=2, sticky=(W))
        self.holdTogetherSpeech_ttp = CreateToolTip(self.holdTogetherSpeech_Label, text=("Holds up to 6 paragraphs together (no empty line in bewteeen),\n"
                                                                                        "that end and begin with direct speech\n"
                                                                                        "Maximum of \{Value\} words per Paragraph"))
        self.holdTogetherSpeech_Field = ttk.Entry(self.settings_Frame, textvariable=self.holdTogetherSpeech, width=5)
        self.holdTogetherSpeech_Field.grid(column=4, row=5, sticky=(W))

        ttk.Label(self.settings_Frame, width=5).grid(column=5, row=5, sticky=(NSEW))

        self.identFirstLine = tk.StringVar()
        self.identFirstLine.set(self.config['identfirstline'])
        self.configValidator.addConfigLabel('identfirstline', self.preamble)
        self.identFirstLine.trace_add('write', lambda *args: self.changeSettings(self.identFirstLine, 'indentFirstLine'))
        self.identFirstLineInfo_Button = ttk.Button(self.settings_Frame, text="i", command=lambda: self.ShowInfoText('identFirstLine'), width=3)
        self.identFirstLineInfo_Button.grid(column=6, row=5, sticky=(E))
        self.identFirstLineInfo_ttp = CreateToolTip(self.identFirstLineInfo_Button, text='Extra Info')
        self.identFirstLine_Label = ttk.Label(self.settings_Frame, text='Ident First Paragraph', width=30)
        self.identFirstLine_Label.grid(column=7, row=5, columnspan=3, sticky=(W))
        self.identFirstLine_ttp = CreateToolTip(self.identFirstLine_Label, text=("Indents first Line of a Paragraph by \{Value\} spaces\n"
                                                                                "Discouraged with justification enabled"))
        self.identFirstLine_Field = ttk.Entry(self.settings_Frame, textvariable=self.identFirstLine, width=5)
        self.identFirstLine_Field.grid(column=9, row=5, sticky=(W))


        self.outputPath = tk.IntVar()
        self.outputPath.set(self.config['outputpath'])
        self.outputPath.trace_add('write', lambda *args: self.changeSettings(self.outputPath, 'outputPath'))
        self.outputPathInfo_Button = ttk.Button(self.settings_Frame, text="i", command=lambda: self.ShowInfoText('outputPath'), width=3)
        self.outputPathInfo_Button.grid(column=1, row=6, sticky=(W))
        self.outputPath_Info = CreateToolTip(self.outputPathInfo_Button, text='Extra Info')
        self.outputpath0 = ttk.Radiobutton(self.settings_Frame, text='Don\'t generate an output-file', variable=self.outputPath, value=0)
        self.outputpath0.grid(column=2, row=6, columnspan=3, sticky=(W))
        self.outputpath0_ttk = CreateToolTip(self.outputpath0, text=("Won\'t generate an output file when called"))
        self.outputpath1 = ttk.Radiobutton(self.settings_Frame, text='Save next to source', variable=self.outputPath, value=1)
        self.outputpath1.grid(column=4, row=6, columnspan=3, sticky=(W))
        self.outputpath1_ttk = CreateToolTip(self.outputpath1, text=("Will spawn a file named \{OriginalName\}_BB.txt next to provided source."))
        self.outputpath2 = ttk.Radiobutton(self.settings_Frame, text='Select file destination', variable=self.outputPath, value=2)
        self.outputpath2.grid(column=7, row=6, columnspan=2, sticky=(W))
        self.outputpath2_ttk = CreateToolTip(self.outputpath2, text=("Will open a dialog to select file destination and name manually"))

        self.clipboard = tk.BooleanVar()
        self.clipboard.set(self.config['clipboard'])
        self.clipboard.trace_add('write', lambda *args: self.changeSettings(self.clipboard, 'clipboard'))
        self.clipboard_Button = ttk.Button(self.settings_Frame, text="i", command=lambda: self.ShowInfoText('clipboard'), width=3)
        self.clipboard_Button.grid(column=1, row=7, sticky=(W))
        self.clipboard_ttp = CreateToolTip(self.clipboard_Button, text='Extra info')
        self.clipboard_Check = ttk.Checkbutton(self.settings_Frame, text='Copy result to clipboard',
                                                            variable=self.clipboard, onvalue=True, offvalue=False, width=50)
        self.clipboard_Check.grid(column=2, row=7, columnspan=3, sticky=(W))
        self.clipboard_ttp = CreateToolTip(self.clipboard_Check, text="Controls if result should be pasted directly to clipboard")

        self.preview = tk.BooleanVar()
        self.preview.set(True)
        self.preview.trace_add('write', lambda *args: self.changeSettings(self.preview, 'preview'))
        self.preview_Button = ttk.Button(self.settings_Frame, text="i", command=lambda: self.ShowInfoText('preview'), width=3)
        self.preview_Button.grid(column=6, row=7, sticky=(E))
        self.preview_ttp = CreateToolTip(self.preview_Button, text='Extra info')
        self.preview_Check = ttk.Checkbutton(self.settings_Frame, text='Show Preview',
                                                            variable=self.clipboard, onvalue=True, offvalue=False, width=50)
        self.preview_Check.grid(column=7, row=7, columnspan=3, sticky=(W))
        self.preview_ttp = CreateToolTip(self.preview_Check, text=("Shows a preview of converted Text\n"
                                                                    "Note that this is a simplified custom rendering and may look different from the real result\n"
                                                                    "Justification will not be displayed"))

    def drawSearchAndReplace(self):
        self.sr_Frame = ttk.Labelframe(self.mainframe, text='Search & Replace', width=300, height=50)
        self.sr_Frame.grid(column=0, row=6, columnspan=3, sticky=(NW))


        self.srInfo_Button = ttk.Button(self.sr_Frame, text="i", command=lambda: self.ShowInfoText('searchandreplace'), width=3)
        self.srInfo_Button.grid(column=0, row=0, sticky=(W))
        self.srInfo_ttp = CreateToolTip(self.srInfo_Button, text='Extra info')
        self.sr_Listbox = tk.Listbox(self.sr_Frame, height=4, width = 30)
        self.sr_Listbox.grid(column=1, row=0, sticky=(NW))
        self.srList_Scrollbar = ttk.Scrollbar(self.sr_Frame, orient= tk.VERTICAL, command=self.sr_Listbox.yview)
        self.srList_Scrollbar.grid(column=2, row=0, sticky=(N,S))
        self.sr_Listbox['yscrollcommand'] = self.srList_Scrollbar.set
        self.sr_Listbox.bind("<<ListboxSelect>>", self.srSelect)

        self.srEdit_Frame = ttk.Frame(self.sr_Frame, width=200, height=40)
        self.srEdit_Frame.grid(column=3, row=0, sticky=(NW))

        self.deleteSR_Button = ttk.Button(self.srEdit_Frame, text="Delete", command=self.deleteSR)
        self.deleteSR_Button.grid(column=0, row=0, sticky=(W))
        self.deleteSR_ttp = CreateToolTip(self.deleteSR_Button, text='Permanently delete search & replace rule')

        self.addSR_Button = ttk.Button(self.srEdit_Frame, text="Add", command=self.addSR)
        self.addSR_Button.grid(column=0, row=2, sticky=(W))
        self.addSR_ttp = CreateToolTip(self.addSR_Button, text='Add a new search & replace rule')

        self.moveUpSR_Button = ttk.Button(self.srEdit_Frame, text="Move Up", command=self.moveUp)
        self.moveUpSR_Button.grid(column=1, row=0, sticky=(W))
        self.moveUpSR_ttp = CreateToolTip(self.moveUpSR_Button, text='Move rule up in the order of execution')

        self.moveDownSR_Button = ttk.Button(self.srEdit_Frame, text="Move Down", command=self.moveDown)
        self.moveDownSR_Button.grid(column=1, row=2, sticky=(W))
        self.moveDownSR_ttp = CreateToolTip(self.moveDownSR_Button, text='Move rule down in the order of execution')

        #self.srFieldSpace_Label = tk.Label(self.srEdit_Frame, width = 3)
        #self.srFieldSpace_Label.grid(column=1, row=0)

        self.srSearch = tk.StringVar()
        self.srSearch.trace_add('write', lambda *args: self.changeSettings(self.srSearch, 'srSearch'))
        self.srSearch_Label = ttk.Label(self.srEdit_Frame, text='Search', width=12)
        self.srSearch_Label.grid(column=2, row=0, sticky=(E))
        self.srSearch_ttp = CreateToolTip(self.srSearch_Label, text=("Set Search Parameter\n"
                                                                    "Must be valid Regex: See extra info for details"))
        self.srSearch_Field = ttk.Entry(self.srEdit_Frame, textvariable=self.srSearch, width=58)
        self.srSearch_Field.grid(column=3, row=0, sticky=(E))

        self.srReplace = tk.StringVar()
        self.srReplace.trace_add('write', lambda *args: self.changeSettings(self.srReplace, 'srReplace'))
        self.srReplace_Label = ttk.Label(self.srEdit_Frame, text='Replace', width=12)
        self.srReplace_Label.grid(column=2, row=1, sticky=(E))
        self.srReplace_ttp = CreateToolTip(self.srReplace_Label, text=("Set Replace Value\n"
                                                                        "Enter as a normal String, for special characters see extra info"))
        self.srReplace_Field = ttk.Entry(self.srEdit_Frame, textvariable=self.srReplace, width=58)
        self.srReplace_Field.grid(column=3, row=1, sticky=(E))

        self.enableSR = tk.BooleanVar()
        self.enableSR.trace_add('write', lambda *args: self.changeSettings(self.enableSR, 'enableSR'))
        self.enableSR_Check = ttk.Checkbutton(self.srEdit_Frame, text='Enabled', variable=self.enableSR, onvalue=True, offvalue=False, width=10)
        self.enableSR_Check.grid(column=2, row=2, columnspan=2, sticky=(W))
        self.enableSR_ttp = CreateToolTip(self.enableSR_Check, text=("Enable/Disable Rule"))

        self.reorderAndSetSRRule(changeSelection=True)

    def drawExtraInfo(self):

        self.extraInfo_Frame = ttk.Labelframe(self.mainframe, text="Info & Preview", width = 300)
        self.extraInfo_Frame.grid(column=0, row=7, columnspan=3, sticky=(NW))

        self.extraInfoLeftSpace_Label = ttk.Label(self.extraInfo_Frame, width = 10)
        self.extraInfoLeftSpace_Label.grid(column=0, row=0)

        self.extraInfoText_Frame = ttk.Frame(self.extraInfo_Frame, width=700, height = 250)
        self.extraInfoText_Frame.grid(column = 1, row=0, sticky=(NW))
        self.extraInfoText_Frame.columnconfigure(0, weight=10)
        self.extraInfoText_Frame.grid_propagate(False)

        self.extraInfo_Box = tk.Text(self.extraInfoText_Frame, height=20, padx=15, pady=15, wrap=tk.WORD)
        self.extraInfo_Box.grid(sticky="we")
        self.extraInfo_Box.insert("1.0", "Extra Info")
        self.extraInfo_Box.configure(state='disabled')
        self.extraInfoText_Frame.configure(width=580)

        self.extraInfo_Scrollbar = ttk.Scrollbar(self.extraInfo_Frame, orient= tk.VERTICAL, command=self.extraInfo_Box.yview)
        self.extraInfo_Scrollbar.grid(column=2, row=0, sticky=(N,S))
        self.extraInfo_Box['yscrollcommand'] = self.extraInfo_Scrollbar.set

    def reorderAndSetSRRule(self, selectName=False, changeSelection=False):
        if not self.config['searchandreplace']:
            return
        self.config['searchandreplace'] = sorted(self.config['searchandreplace'], key=lambda x: x[0], reverse=True)
        self.sr_Listbox.delete(0, tk.END)
        select = 0
        for i, sr in enumerate(self.config['searchandreplace']):
            self.sr_Listbox.insert('end', sr[1])
            if sr[0] > 0:
                self.sr_Listbox.itemconfig(i, foreground="black")
            else:
                self.sr_Listbox.itemconfig(i, foreground="grey")
            if sr[1] == selectName:
                select = i
        if changeSelection:
            self.sr_Listbox.select_set(select)
            self.sr_Listbox.event_generate("<<ListboxSelect>>")
    
    def srSelect(self, event):
        self.srWidget = event.widget
        try:
            self.srSelection=self.srWidget.curselection()[0]
        except IndexError:
            return
        sr =  self.config['searchandreplace'][self.srSelection]
        if sr[0] > 0:
            self.enableSR.set(True)
        else:
            self.enableSR.set(False)
        self.srSearch.set(sr[2])
        self.srReplace.set(sr[3])

    def deleteSR(self, *args):
        msg = 'Are you sure to permanently delete "' + self.srWidget.get(self.srSelection) + '"?' 
        response = messagebox.askyesno(message=msg, icon='question', title='Delete Search & Replace Rule')
        if response:
            self.config['searchandreplace'].pop(self.srSelection)
            self.reorderAndSetSRRule(changeSelection=True)
    
    def addSR(self, *args):
        inputDialog = InputDialog(self.root, text='Enter a name for the new Rule')
        self.root.wait_window(inputDialog.top)
        name = inputDialog.getEntry()
        if name:
            imp = self.getMaxSrImportance() + 1
            self.config['searchandreplace'].insert(0, (imp, name, "", ""))
            self.reorderAndSetSRRule(selectName=name, changeSelection=True)

    def moveDown(self, *args):
        if not self.enableSR: #already not executed
            return
        if self.srSelection + 1 >= len(self.config['searchandreplace']): #already last
            return
        if not self.config['searchandreplace'][self.srSelection+1][0]: #last enabled
            return
        #swap
        self.config['searchandreplace'][self.srSelection+1][0], self.config['searchandreplace'][self.srSelection][0] = \
        self.config['searchandreplace'][self.srSelection][0], self.config['searchandreplace'][self.srSelection+1][0]
        self.reorderAndSetSRRule(selectName=self.config['searchandreplace'][self.srSelection][1], changeSelection=True)

    def moveUp(self, *args):
        if not self.enableSR: #already not executed
            return
        if self.srSelection == 0: #already first
            return
        #swap
        self.config['searchandreplace'][self.srSelection-1][0], self.config['searchandreplace'][self.srSelection][0] = \
        self.config['searchandreplace'][self.srSelection][0], self.config['searchandreplace'][self.srSelection-1][0]
        self.reorderAndSetSRRule(selectName=self.config['searchandreplace'][self.srSelection][1], changeSelection=True)

    def getMaxSrImportance(self):
        return self.config['searchandreplace'][0][0]

    def convert(self):
        try:
            tryParseConfig(self.config, self.style_enabled, validator=self.configValidator)
        except BadConfigException as e:
            origin, msg = e.getInfo()
            self.status.set('Error while parsing' + str(origin) +' : ' + msg)
            return
        except DeprecatedConfigException as e:
            self.status.set('Missing config-data: Backup/rename your current config-file manually and try again')
            return
        self.configValidator.saveConfig(self.config, self.style_enabled)
        if self.filePath.get() == '':
            self.status.set('No file selected to convert')
            return
        try:
            cleanup = False
            filename = os.path.basename(self.filePath.get())
            file, cleanup = genDocx(self.filePath.get())
            docx = Document(file)
            self.parsedTXT = parseDocx(docx, self.config, self.style_enabled)
            self.writeTxt(self.parsedTXT, file, filename)
            if self.preview.get():
                if self.parsedTXT:
                    self.ShowInfoText(None, info=self.parsedTXT)
        except TimeoutError:
            self.status('Timeout while trying to write. Please ensure TextToBB has WriteAccess to the directory your source-file is in.')
        except Exception as e:
            print('Damn, you found another bug.')
            print("Please report the issue with the following information:")
            traceback.print_exc(file=sys.stdout)
            self.status.set('Error while parsing file, please see console log for further information')
        finally:
            if cleanup:
                try:
                    os.remove(file)
                except OSError as e:  ## if failed, report it back to the user ##
                    print("Error: %s - %s." % (e.filename, e.strerror))
            try:
                if config['keepopen']:
                    input()
            except:
                pass

    def previewButtonPressed(self):
        if self.showPreview.get() == 'Show Preview':
            self.showPreview.set('Disable Preview')
        else:
            self.showPreview.set('Show Preview')

    def ShowInfoText(self, caller, info=''):
        try:
            if info == '':
                info = documentation.extrainfo.get(caller)
            self.extraInfo_Box.configure(state='normal')
            BBToTkText().parse(self.extraInfo_Box, self.extraInfoText_Frame, s=info)
            self.extraInfo_Box.configure(state='disabled')
        except Exception as e:
            if self.init: return
            raise

    def changeSettings(self, caller, name):
        name = str(name).lower()
        if name in self.config:
            self.config[name] = caller.get()
        elif name in self.style_enabled:
            self.style_enabled[name] = caller.get()
        elif name =='enablesr':
            if caller.get() and self.config['searchandreplace'][self.srSelection][0] <=0:
                self.config['searchandreplace'][self.srSelection][0] = self.getMaxSrImportance() + 1
                self.reorderAndSetSRRule(changeSelection=True)
            if not caller.get() and self.config['searchandreplace'][self.srSelection][0] > 0:
                self.config['searchandreplace'][self.srSelection][0] = 0
                self.reorderAndSetSRRule(changeSelection=True, selectName=self.config['searchandreplace'][self.srSelection][1])
        elif name == 'srsearch':
            self.config['searchandreplace'][self.srSelection][2] = caller.get()
        elif name =='srreplace':
            self.config['searchandreplace'][self.srSelection][3] = caller.get()
        if self.preview.get():
            self.convert()
        

    def getFile(self):
        file_path = filedialog.askopenfilename(filetypes = (("Word/Office",("*.docx", "*.doc")),("All files","*.*")))
        self.filePath.set(file_path)

    def writeTxt(self, txt, source, filename):
        self.status.set('Success!   ')
        if self.config['clipboard']:
            pyperclip.copy(txt)
            self.status.set(self.status.get()+" Copied to Clipoard")
        outputpath = self.config['outputpath']
        if outputpath:
            outputName = os.path.splitext(filename)[0] + "_BB"
            #outputName = re.sub('(?<!\\\)\$', outputName, self.config['outputname'])
            if outputpath == 1 or outputpath == 2:
                outputDir = os.path.dirname(source)
                if outputpath == 2:
                    if os.path.isfile(os.path.join(outputDir, outputName)):
                        cnt = 1
                        while os.path.isfile(os.path.join(outputDir, outputName + '(' + cnt + ')')):
                            cnt += 1
                        outputName += '(' + cnt + ')'
            with codecs.open(os.path.join(outputDir, outputName + '.txt'), "w", 'utf-8') as text_file:
                text_file.write(txt)
            #print 'Output saved as ' + outputName + '.txt'
            #print str(len(txt)) + ' characters'

#Main Parser
def parseDocx(document, config, styleOptions, maxparagraphs=sys.maxsize):
    newFileString = u''
    first = True
    paraStack = PARASTACK_MAX
    br = config['endlinechar']
    if config['emptylineafterparagraph']:
        endline = 2
    else:
        endline = 1
    paraStyle = ParaStyles(endline, config['identfirstline'], styleOptions)
    for i, para in enumerate(document.paragraphs):
        if i > maxparagraphs:
            break

        #skip empty lines
        if config['skipemptylines'] and  para.text == '':
            continue

        #handle first line and special lines
        if first:
            first = False
            if config['preamble']:
                newFileString += replaceLinebreaks(config['endlinechar'], config['preamble'])
            dateStr = datetime.datetime.now().strftime(replaceLinebreaks(config['endlinechar'], config['copyrightdateformat']))
            firstline = re.sub(r'(?<!\\)\$cr', COPYRIGHT, replaceLinebreaks(config['endlinechar'], config['header']), flags=re.UNICODE)
            firstline = re.sub(r'(?<!\\)\$date', dateStr, firstline, flags=re.UNICODE)
            firstline = re.sub(r'(?<!\\)\$title', para.text, firstline, flags=re.UNICODE)
            firstline = re.sub(r'\\\$', '$', firstline, flags=re.UNICODE)
            newFileString += firstline
            continue
    
        #parse paragraph
        newPara = u''     
        newPara, paraStyle = preamblePara(newPara, para, paraStyle)
        newPara = parsePara(newPara, para, paraStyle)

        #handle special replacement options
        for enabled, name, special, replace in config['searchandreplace']:
            if not enabled: continue
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
        newFileString += replaceLinebreaks(config['endlinechar'], config['preamble'])

    return newFileString   

def preamblePara(newPara, para, style):
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

#Read Config file
def readConfig(validator = ConfigValidator()):
    config = RawConfigParser()
    if not os.path.exists(validator.getFilePath()):
        handleMissingConfig(validator==validator)
    while True:
        config.read_file(codecs.open(validator.getFilePath(), "r", "utf-8"))
        try:
            version = dict(config.items('Version'), raw=True)
            default = dict(config.items('Settings'), raw=True)
            styleOptions = dict(config.items('StyleOptions'), raw=True)
            break
        except NoSectionError:
            handleMissingConfig(validator=validator)
    version = str(version['version'])
    if version != str(VERSION):
        handleVersionError(config)
    return tryParseConfig(default, styleOptions) 

def tryParseConfig(default, styleOptions, validator = ConfigValidator()):   
    try:
        default['endlinechar'] = u"\r\n"
        while True:
            try:
                default, styleOptions = validator.parseConfig((default, styleOptions))
                break
            except BadConfigException as e:
                    handleBadConfig(e, (default, styleOptions), validator=validator)     
        if not default['emptylineafterparagraph']:
            default['holdtogetherspeech'] = 0
        return  default, styleOptions
        #additional validation
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        raise SystemExit

#handle Config Errors
def handleMissingConfig(validator=ConfigValidator()):
    errormsg = 'Can\'t find or read ' + validator.getFilePath() + '\n Generate a new config file with default values?'
    response = messagebox.askyesno(message=errormsg, icon='error', title='Missing Config')
    if response:
        validator.generateNewConfig()
    else:
        raise ValueError

def handleBadConfig(e, configlist, validator=ConfigValidator()):
    origin, msg = e.getInfo()
    errormsg = 'Error while parsing ' + str(origin) +' : ' + msg + '\n Use default setting for "' + str(origin) + '"?'
    response = messagebox.askyesno(message=errormsg, icon='error', title='Bad Config')
    if response:
        for i, config in enumerate(configlist):
            if origin in config:
                configlist[i][origin] = validator.getDefault(origin)
    else:
        raise e

def replaceLinebreaks(endline, input):
    return re.sub(r'(?<!\\)\[\/br\]', endline, input, flags=re.UNICODE)

def handleVersionError(config):
    raise NotImplementedError

#utility
def id_generator(size=16, chars=string.ascii_uppercase + string.digits):
   return ''.join(random.choice(chars) for _ in range(size))          

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

if __name__ == '__main__':
    main = TextToBB()

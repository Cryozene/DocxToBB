# DocxToBB
Creates a text-string out of a .docx (Added with 0.3: Also accepts .doc) file, that inherits select text-fonts with BBCode and formats the text for easier reading on a computer screen.  
Each Option is fully documented within the GUI.  
Written for romane-forum.de


## Current Build: Beta 1 (Version 0.6) 

## How to use

1. Executable is compiled for Windows 64-Bit systems directly from the source .py file, users on other systems will need to download and use the source (Python 3.7)  
2. To run, simply start the executable(or .py if running form source)
*No installation needed!*
3. When starting the first time, most likely an alert will inform you, that 'DocxToBB.ini' is missing. Click 'OK' for it to be generated
4. Select a file to convert, press 'Convert' - finished. If the `clipboard`-option is set to true, simply paste the resulting text at the desired place
5. As long as 'preview' is enabled, a preview of your file will present itself at the bottom. Make sure to change the settings and convert as often as you like, until you are satisfied with the result. All settings will be saved, when 'Convert' is pressed (even when no file is selected!). Possible Errors will be displayed up top.

## Troubleshooting 

- **Output looks ugly**  
Please have a look at the various options and change them to your liking. Click the buttons with "i" on it for more information about each setting. If you feel something is missing or doesn't behave as expected, refer to **Everything is broken!**
- **The tool deletes linebreaks!**  
If you're using the standard notepad editor from Windows (seriously though: Why would you do that?), make sure that `pruneWhitespace` ist set to `True`.  
- **The tool still deletes linebreaks!**  
Yes. Any Empty line in the source file will be skipped by the parser, in order to create an homogenous look. If you want to turn off that behavior, set 'skip empty lines' to `True`. 
Additional line-breaks currently can be added only via the Search & Replace field.
- **My headings are gone!**  
Please do not format the source with special options, such as lists, but as standard Paragraphs. Currently, any changes in textsize won't be inherited. This may change in a later version though. 
- **Everything is broken!**  
Please open an issue and describe you problem(s) as detailed as possible or contact me through other means. I'll try to fix any bugs as fast as possible and add common problems to this section. 




<a rel="license" href="http://creativecommons.org/licenses/by-sa/4.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by-sa/4.0/80x15.png" /></a><br />This work is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by-sa/4.0/">Creative Commons Attribution-ShareAlike 4.0 International License</a>.

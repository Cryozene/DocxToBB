# DocxToBB
Creates a string out of a .docx file, that inherits select text-fonts with BBCode
Written for romane-forum.de


## Current Build: Beta 1 (Version 0.2) 

## How to use

1. Executable is compiled for Windows 64-Bit systems, users on other systems will need to downlaod and use the source  
2. Python-file is python 2.7, all requirements are in imports  
3. To run, simply place DocxToBB.ini next to the excutable (or .py if running from source) No Installation needed!
4. Before starting for the first time, make sure to open DocxToBB.ini with the text-editor of your choice (I recommend notepad++), take a glance at the options and change the to your liking  
5. Run the Executable, select a .docx file to convert - finished. If the `clipboard`-option is set to true, simply paste the resulting text at the desired place

## Troubleshooting 

- **Output looks ugly**  
Please have a look at the various options in `DocxToBB.ini` and change them to your liking. Everything should be explained in the file.
If you feel something is missing or doesn't behave as you expected, refer to **Everything is broken!**
- **The Tool deletes linebreaks!**  
Make sure that you have set the correct endline-characters in the .ini file. If you're using the standard notepad editor from Windows (seriously though: Why would you do that?), make sure that `pruneWhitespace` ist set to `True`  
- **The Tool still deletes linebreaks!**  
Yes. Any Empty line in the source file will be skipped by the parser, in order to create an homogenous look. Extra line-breaks currently can be added via the `searchFor` and `replaceWith` options. If you wnat to turn off the behavior, set `skipemptylines` to `False`
- **My headings are gone!**  
Please do not format the with special options, such as chapters, but as standard Paragraphs. Currently, any changes in size won't be inherited. This may change in a later version though. 
- **Everything is broken!**  
Please open an issue and describe the issue as detailed as possible or contact me by other means. I'll try to fix any bugs as fast as possible and add common problems to this section. 




<a rel="license" href="http://creativecommons.org/licenses/by-sa/4.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by-sa/4.0/80x15.png" /></a><br />This work is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by-sa/4.0/">Creative Commons Attribution-ShareAlike 4.0 International License</a>.

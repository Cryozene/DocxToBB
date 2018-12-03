#Documentation for GUI
linebreakInfo = \
r'''
[c][b][size=120]Linebreaks[/size][/b][/c]
To create linebreaks within [i]preamble[/i], [i]postamble[/i] or [i]header[/i], use \[/br]

[i]Line 1[/br]Line 2[/i] will show up as:
Line 1
Line 2

Also be careful, that "\[c]","\[r]" and "\[j]" will force a linebreak, but delete the preceding one.
The closing options "\[/c]","\[/r]" and "\[/j]" will create a linebreak.
Therefore, "A [/br][/br]\[c] B \[/c][/br] C" will have an empty line between B and C, but not between A and B.

This is actually a limitation of the underlying parser from BBCode-Tags to html.


[c][b][size=120]Backslashes[/size][/b][/c]
Backslashes must be escaped with another backslash.
For example, if "a\b\c" should be written, "a\\b\\c" has to be entered


[c][b][size=120]Whitespace[/size][/b][/c]
Leading or multiple whitespaces will be ignored by the Parser. To prevent this behavior use NO_BREAK_SPACE at every part, where multiple or leading Whitespace is desired. On Windows, they can be entered by holding ALT and entering 255 via the numpad
Alternative, type "\u20A0" (explanation below)


[c][b][size=120]Special characters[/size][/b][/c]
Every character is represented by a Unicode-sequence
NO_BREAK_SPACE for "special" whitespace is an example for this: "\u20A0" will be interpreted as that special character
An Ellipsis works the same way: "\u2026" will show up as …

When clicking "Convert", the settings will be interpreted and show up as their respective single-characters.
When searching for the Unicode sequence for other characters, make sure to get the one for python
Generally, a Google-search with "*Name of the character* Unicode python" will yield a good result

Useful characters:
NO_BREAK_SPACE: \u20A0, Ellipsis: \u2026 
Double quotation marks (german) are \u201E (open) and \u201C (close)
Single quotation marks (german) are \u201A (open) and \u2018 (close)
Double quotation marks (english) are \u201C (open) and \u201D (close)
Single quotation marks (english) are \u2018 (open) and \u2019 (close)
Double quotation marks (french) are \u00AB (open) and \u00BB (close)
Single quotation marks (french) are \u2039 (open) and \u203A (close)
'''
extrainfo ={
'preamble':
r'''[c][b][size=150]Preamble[/size][/b][/c] 
[i]For info regarding linebreaks and special characters see below![/i] 

- The preamble of the whole file

- Defines General Styles, that are not already present within the source text

- Must consist of BBCode-Tags such as \[book], \[size=x] etc. and/or plain text

'''+ linebreakInfo,
'postamble':
r'''[c][b][size=150]Postamble[/size][/b][/c]
[i]For info regarding linebreaks and special characters see below![/i]

- The postamble of the whole file

- Any open Tags, such as \[book] or \[seite] should be closed here

- Must consist of BBCode-Tags such as \[/book], \[/size] etc. and/or plain text

'''+ linebreakInfo,
'header':
r'''[c][b][size=150]Header[/size][/b][/c]
[i]For info regarding linebreaks and special characters see below![/i]

- Uses this style for the first line and Copyright

- Enter Format as valid BBCode, omit any placeholder you don't want to see

- Please make sure to close any opened BBCode-Tags to prevent unwanted side-effects

- \title refers to the first line of your Textfile (usually the title). If you don't want any special formatting applied to it, enter it after at the of your header

- \cr refers to the copyright symbol "\u00A9" will yield the same result [cr=year]author[/cr] is valid, but discouraged

- \date refers to the date in the format specified below

''' + linebreakInfo,
'copyrightDateFormat': 
r'''[c][b][size=150]Date Format[/size][/b][/c]
- Defines the desired style for \date in header

- Use either fixed Value or any date format according to python's datetime.strftime(copyrightDay) specified below:

[i]Extract from official documentation:[/i]
*Output depends on the local language Settings

Directive-> Meaning (Example)
%a -> Weekday as locale’s abbreviated name. (Su, Mo, …, Sa*)
%A -> Weekday as locale’s full name. (Sunday, Monday, …, Saturday*)
%w -> Weekday as a decimal number, where 0 is Sunday and 6 is Saturday. (0, 1, …, 6)
%d -> Day of the month as a zero-padded decimal number. (01, 02, …, 31)
%b -> Month as locale’s abbreviated name. (Jan, Feb, …, Dez*)
%B -> Month as locale’s full name. (January, February, …, December*)
%m -> Month as a zero-padded decimal number. (01, 02, …, 12)
%y -> Year without century as a zero-padded decimal number. (00, 01, …, 99)
%Y -> Year with century as a decimal number. (1970, 1988, 2001, 2013)
%j -> Day of the year as a zero-padded decimal number.(001, 002, …, 366)
%U -> Week number of the year (Sunday as the first day of the week) (00, 01, …, 53)
%W -> Week number of the year (Monday as the first day of the week) (00, 01, …, 53)
%x -> Locale’s appropriate date representation. (16.08.1988*)
%% -> A literal '%' character.(%)
''',
'emptyLineAfterParagraph':
r'''[c][b][size=150]Empty Line After Paragraph[/size][/b][/c]
- Will add an extra linebreak after each paragraph to improve readability
''',
'skipEmptyLines' :
r'''[c][b][size=150]Skip Empty Lines[/size][/b][/c]
- This setting will tell the parser to skip empty lines of the source file

- This will leave the control of linebreaks completely to the actual Parser and should create a more homogenous look

- Disable This setting, if you want empty lines of the source preserved
''',
'holdTogetherSpeech':
r'''[c][b][size=150]Hold direct Speech together[/size][/b][/c]
- Only has an Influence, if Empty Line After Paragraph is enabled

- Will tell the parser, to hold short segments of direct speech together with no extra line-break in-between (Maximum of 6 Paragraphs)

- The value controls the maximum number of words in the preceding paragraph to be held together with the next

- Set to 0, to disable

Example: 
[i]"Hey"
"How are you?"
"Good"[/i]
Will be held together with a value of 3 or bigger.

With a Value of 2, the result would look like the following:
[i]"Hey"
"How are you?"

"Good"[/i]
''',
'identFirstLine':
r'''[c][b][size=150]Ident First Line[/size][/b][/c]
- Sets indentation of \{Value\} whitespace at first Line in a Paragraph, in order to create a standard "book"-feeling

- Set to 0, to disable

- Justification should be OFF when the value is greater than 0 (Otherwise, the indentation will vary wildly)#

Example:
  Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua.
  Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua.
''',
'outputPath' :
r'''[c][b][size=150]Output Path[/size][/b][/c]
- Set's the destination to save the .txt file containing the parsed text

- Don't generate an output file: This option will skip the step of saving to a file. Useful when Copy to Clipboard is enabled, to prevent junk files from being created

- Save next to source: This will save a .txt file with {SourceName}_BB.txt next to the source [b] and overwrite any existing file with that name. [/b]

- Select file destination: This will open a dialog to specify a destination and name for each conversion
''',
'clipboard':
r'''[c][b][size=150]Copy Result to Clipboard[/size][/b][/c]

This will add the conversion result directly to clipboard

Simply paste the result wherever you like
''',
'searchandreplace':
r'''[c][b][size=150]Search an Replace[/size][/b][/c]

This option will add a couple of search & replace operations in the specified order, from top to bottom
- 'Add' appends a rule to the order of execution

- 'Delete' will completely remove the current rule, so that it won't show up anymore. Keep in mind, that a rule can always be disabled instead of deleted!

- 'Move Up' and 'Move Down' Let you Change the position at which the current rule will be executed. Top goes first, then second etc. 

- 'Enable' allows you to disable /enable a rule. A disabled rule won't be executed but is saved for later.

- Both \[book] and \[seite] have a corresponding rule to "fix" the display of an ellipsis in the chosen font. It is recommend to enable the one corresponding to your chosen environment, move it to the top and disable the other one 


[c][b][size=120]Replace[/size][/b][/c]
The easy part: A literal replacement for the "Search", unicode patterns such as u\2026 (an ellipsis) are allowed and will be converted accordingly
Literal backslashes need to be escaped with another backslash (like always)!
(To reference the Search within the replace-filed use "\1")


[c][b][size=120]Search[/size][/b][/c]
The search-Term must be entered as a [i]Regular expression (or Regex)[/i](For edge-cases: python-Interpreter, linebreaks not allowed).
This is very similar to a normal, literal search, but yields a few extra options with just a few caveats:
A search term is a [i]pattern[/i].

- For literal searching patterns consisting only of letters and/or numbers, the pattern is the same as normal: searching for "where" would need the pattern "where" in the search-field
Note that unicode-characters (\uXXXX) are allowed as well

- If your pattern has one of the following literals, they need to be preceded by a backslash each:
\^$.|?*+!$(){}[]
If you want so search for "...", the pattern would be "\.\.\."

-Why Regex though? A good example is a common issue of not leaving a space before an ellipsis, or - in the case of \[book] - three dots:
Let's go with the latter one. Every time exactly three dots are not preceded by whitespace, it should be added. This would not be possible with a normal search& replace operation!

A very brief introduction into the possibilities of Regex follows
.
For a more comprehensive or better-written tutorial on Regular expressions take a search engine of your choice and pick the one to your liking.

For testing and understanding regex-expressions I recommend regex101.com (make sure to select "python, if the expression is supposed to work here as well).


Introduction:
Regex has a few useful patterns
[ab] for example matches either a or b
[^ab] matches any character that is not a or b
[a-z] matches any character between a and z (such as d, h, ... but not H, as this is uppercase)

\w matches any word letter, \d matches any digit
[ab]* matches any combination out of a and b, as long as possible: aaaabbbb would be matched, as well as a
(ab)+ groups ab together and matches any recurring sequence of ab: ababab
a{x, y} matches a between x and y times
^  = beginning of line, $ = end of line
a(?b) = lookahead: a is followed by b; a(?!b) = negative lookahead: a is not followed by b
(?<b)a = lookbehind: a is preceded by b; (?<!b)a = negative lookbehind: a is not preceded by b

For example, this solves the previous example for 3 dots not preceded by whitespace as follows:

(?<![\. ])\.\.\.(?!\.) -> \.\.\. matches 3 dots, (?<![\. ]) negative lookbehind: not preceded by another dot or a whitespace, (?!\.) negative lookahead: not followed by another dot

Another example would be The automatic formatting of a line with roman numerals, a chapter-heading like "IV.":
"^[IVXLCDM]+\.$" -> "^" asserts the beginning of the line; "[IVXLCDM]+" matches any combination of roman numerals with at least one occurrence; "\." matches the trailing dot and "$" asserts the end of the line

A possible replacement would be the following: \u20A0[c][b]\1[/b][/c]
\u20A0 makes sure, that no preceding empty line is deleted by centering that line, \1 will match the search an will be substituted by the captured group by the search
'''
}
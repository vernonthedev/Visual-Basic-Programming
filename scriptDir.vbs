Option Explicit

dim fullname, shortname

fullname = wscript.scriptfullname
shortname = wscript.scriptname

'Display the paths with a tab space
msgbox fullname & vbTab & shortname

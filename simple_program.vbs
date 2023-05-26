' For error checking
Option Explicit
' Declare the variables
dim obj, a, b, c
'set the constantly used running object to a var obj
set obj = CreateObject("wscript.shell")
' vbSystemModal Ensures that the program stays on top even after the no buttons have been pressed
' so as the next code can run
a = msgbox("Open a Program?", vbYesNo+vbQuestion+vbSystemModal)

if a = vbYes then
    obj.run "mspaint.exe"
    b = msgbox("Open a folder?", vbYesNo+vbQuestion+vbSystemModal)
else 
    b = msgbox("Open a folder?", vbYesNo+vbQuestion+vbSystemModal)
end if

if b = vbYes then
    obj.run "C:\Users\vernonthedev\Desktop\vbDevelopments"
    c = msgbox("Open a file?", vbYesNo+vbQuestion+vbSystemModal)
else 
    c = msgbox("Open a file?", vbYesNo+vbQuestion+vbSystemModal)
end if

if c = vbYes then
    obj.run """C:\Users\vernonthedev\Desktop\vbDevelopments\New Text Document.txt"""
else 
    wscript.quit
end if
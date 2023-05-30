' Error checking
Option Explicit
'define the variables
dim objfso, objfolder, objfile
'get access to object used for dealing with files and folders
set objfso = CreateObject("Scripting.FileSystemObject")
'set the path to the folder to be used
set objfolder = objfso.GetFolder("C:\Users\vernonthedev\Downloads")

'the for each loop
for each objfile in objfolder.Files
    msgbox objfile.Name
next

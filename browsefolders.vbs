Option Explicit
dim objapp : set objapp = CreateObject("Shell.Application")
dim objfolder,path

set objfolder = objapp.BrowseForFolder(0,"Select a folder: ",0,0)

'if statement to display the path of the selected folder in a message box
if objfolder is nothing then
    msgbox "Canceled!"
    wscript.quit
else
    msgbox objfolder.self.path
end if

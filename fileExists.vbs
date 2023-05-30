Option Explicit
' define the variables
dim ftExists
'create the object that will deal with working with files or folders
set ftExists = CreateObject("Scripting.FileSystemObject")

'Test to see if the file exists
if ftExists.FileExists("C:\Users\vernonthedev\Desktop\test.txt") then
    wscript.echo "Yes! File Exists."
else 
    wscript.echo "File Not Found!"
end if

'Test to see if the folder specified exists and code  put in a sub routine
Sub CheckFolderExist
    if ftExists.FolderExists("C:\Users\vernonthedev\Desktop\vbDevelopments") then
        wscript.echo "Yes! Folder Exists."
    else 
        wscript.echo "Folder Not found"
    end if
    
End Sub

'call the sub
call CheckFolderExist
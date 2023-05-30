Option Explicit

dim fso, file, data
Set fso = CreateObject("Scripting.FileSystemObject")
set file = fso.GetFile("C:\Users\vernonthedev\Desktop\vbDevelopments\New Text Document.txt")


'attributes display
data = "Attributes: " & file.Attributes & vbLf
data = data & "Date Created: " & file.DateCreated & vbLf
data = data & "Date Last Modified: " & file.DateLastModified & vbLf
data = data & "Drive : " & file.Drive & vbLf
data = data & "Name: " & file.Name & vbLf
data = data & "Parent Folder: " & file.ParentFolder & vbLf
data = data & "Path: " & file.Path & vbLf
data = data & "Size: " & file.Size & vbLf
data = data & "Type: " & file.Type & vbLf


'display the attributes in a message box
msgbox data
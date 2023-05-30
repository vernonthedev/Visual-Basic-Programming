Option Explicit

dim fso

set fso = CreateObject("Scripting.FileSystemObject")

'copying one file to another
fso.CopyFile "LOCATION","NEW LOCATION"
fso.CopyFolder "LOCATION","NEW LOCATION"

'moving files of folders to another location
fso.MoveFile "LOCATION", "NEW LOCATION"
fso.MoveFolder "LOCATION", "NEW LOCATION"
Option Explicit

' running a program or a file in vbscript
CreateObject("wscript.shell").run "cmd.exe"
'or i can run many programs like this
dim obj
set obj = CreateObject("wscript.shell")
obj.run "mspaint.exe"
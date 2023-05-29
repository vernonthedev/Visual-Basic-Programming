Option Explicit 
' define the variables
dim x 
' create the starter shell object
Set x = createobject("wscript.shell")

' open notepad and type the following words and keys
x.run "notepad.exe"
' sleep for sometime so that we can allow the program to load 
wscript.sleep 3000

' start typing the keys
x.sendkeys "Hey! Welcome to vbscript scripting!"
x.sendkeys "{enter}"
x.sendkeys "All coded by vernonthedev"
' saving the textfile
wscript.sleep 2000
x.sendkeys "%f"
wscript.sleep 2000
x.sendkeys "s"
wscript.sleep 2000
x.sendkeys "SendingKeys.txt"
wscript.sleep 1000
x.sendkeys "{enter}"
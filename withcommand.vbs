Option Explicit

'automation with the with keyword 
'replacing the repeatedly used objects

with CreateObject("wscript.shell")
    'code to be run with the object
    .run "cmd.exe"
    .run "mspaint.exe"
    .run "notepad.exe"
    
    .sendkeys  "You have reached the end"
    .sendkeys "{enter}"
    .sendkeys "coded by vernonthedev"
    .sendkeys "{tab}"
end with

'another use case
with script
    .echo "Hello this is a printing text"
    .echo "Ssecond statement"
    .sleep "3000"
    .echo "finished"
end with


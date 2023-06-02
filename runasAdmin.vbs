Option Explicit
dim objapp

set objapp = createobject("Shell.Application")
'helps to run cmd as administrator
objapp.ShellExecute "cmd.exe",,,"runas",1
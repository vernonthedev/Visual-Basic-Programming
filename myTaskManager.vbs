Option Explicit
'set a variable and assign its value to an object that mannages windows servvices
dim objWmiService : set objWmiService = GetObject("winmgmts:")
'set the other variables
dim ProcessList, process, prorun
'Query all the available running processes on the computer
' _ represents that the command continues even onto the next line 
set ProcessList = objWmiService.ExecQuery _
("SELECT * FROM Win32_Process")

'loop through the process and display all of them from the performed query set
for each process in ProcessList
    prorun = prorun & process.Name & vbTab
next

'display them in a message box
msgbox prorun
Option Explicit
'set a variable and assign its value to an object that mannages windows servvices
dim objWmiService : set objWmiService = GetObject("winmgmts:")
'set the other variables
dim ProcessList, process
'set a var to a query object that will make selection of the available started windows processes 
' _ represents that the command continues even onto the next line 
set ProcessList = objWmiService.ExecQuery _
("SELECT * FROM Win32_Process WHERE name = 'notepad.exe'")

'loop through the process and terminate the queried one
for each process in ProcessList
    process.Terminate()
next
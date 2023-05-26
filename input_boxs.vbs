'SYNTAX
' inputbox "message","title","input field", xposition,yposition

'inputbox "What is your name", "Information: ", "Name goes here!"

Option Explicit

Dim Name
Name = inputbox("Please Enter your Username?","Login")
msgbox Name, vbOKOnly, "Verify?"
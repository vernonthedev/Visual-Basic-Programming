Option Explicit

dim ie, send

' Create the internet object or start internet explorer
set ie = createobject("InternetExplorer.Application")
set send = createobject("wscript.shell")

' Other ways of formmating the look and feel of the internet explorer
' Goto this specific website
ie.Navigate "https://www.google.com/"
' remove toolbar
ie.Toolbar = 0
' remove status bottom bar
ie.Statusbar = 0
' define the height and width of the ie
ie.Height = 800
ie.Width = 1000
' define from the margin settings
ie.Top = 100
ie.Left = 500
' display the internet explorer after all the settings have been applied
ie.Visible = 1



'if the internet connection is slow/poor...Wait abit and then run the next commands
'after the loop and put in a subroutine so as its called whenever its needed
Sub WaitForLoad
    do while ie.Busy
        wscript.sleep 2000
    loop
End Sub

'call the delay subroutine
call WaitForLoad
'type the following into the google input so as you can search for the text
send.sendkeys "vbscript programming"
send.sendkeys "{enter}"
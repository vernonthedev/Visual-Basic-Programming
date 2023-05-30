Option Explicit
dim objapp, nIE, oIE, Window

'get access to object that deals with software applications
set objapp = CreateObject("Shell.Application")
set oIE = Nothing

for each Window in objapp.Windows
    if Window = "Windows Internet Explorer" then    
        set oIE = Window
    end if
next

if oIE is Nothing then
    Call NewIE
else
    Call OpenIE
end if 

'create a  new ie if internet explorer is closed
Sub NewIE
    set nIE = CreateObject("InternetExplorer.Application")
    nIE.navigate "https://www.youtube.com", 2048
    nIE.visible = True
End Sub


'incase the internet explorer is already open
Sub OpenIE
    oIE.navigate "https://www.youtube.com"
    oIE.visible = True
End Sub



'2048 navOpenNewTab
'4096 navOpenInBackgroundTab
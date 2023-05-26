Option Explicit
'This is for declaring a variable
Dim birth
' Setting the variable
birth = Msgbox("Is it your birthday?", vbYesNo+vbQuestion,"Tell me")
' If and elseif statement i vbs
if birth = vbYes  then
    Msgbox "Happy Birthday!", vbinformation
elseif birth = vbNo then 
    Msgbox "Oops, my Bad!", vbCritical
end if
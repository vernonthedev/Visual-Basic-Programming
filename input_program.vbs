Option Explicit

Dim age

age = inputbox("What is your age?", "AgeTester")

if IsNumeric(age) then
    msgbox "Correct you passed the age gap!"
elseif age = "" then
    msgbox "Please Enter your age"
else 
    msgbox "Please Enter your verified age"
end if
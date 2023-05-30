'functions can be used the same way that subroutines are used
Option Explicit
dim fn, ln

fn = inputbox("Firstname: ")
ln = inputbox("Lastname: ")

'show the inputs to the user and also calling/ using the function at the same time
msgbox "Fullname is: " & fullname(fn, ln)

'function to print the fullnames in a nice format
function fullname(value1, value2)
    dim completeName
    completeName = value1 & " " & value2
    'what to return when the function is run
    fullname = completeName
end function
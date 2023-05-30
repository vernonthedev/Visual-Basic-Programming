Option Explicit
dim a, b, c, pass

a="ben"
b="pizza"

'this is compared with a switch statement in other programming languages
Select case a
    case "pizza"
        msgbox "You have selected piiza"
    case "ben"
        msgbox "You have selected your name"
end select


'another use case
pass = inputbox("what is your master password")
select case pass
    case ""
        msgbox "Please enter your password please!"
    case "codezilla"
        msgbox "Welcome phreak101"
end select

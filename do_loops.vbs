' use of do loop, do while loop, do until loop
Option Explicit

dim average, sum


average = 0
' start the loop
do until average = 5
    ' Whatever you want to loop
    average = average + 1
    msgbox "Average: " & average
loop
' other code to run after the loop

' the normal do while
sum =0
do while sum < 9
    sum = sum + 1
    msgbox "sum is: " & sum
loop

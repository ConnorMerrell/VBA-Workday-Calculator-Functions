Attribute VB_Name = "WorkdayCalc"
Option Compare Database

Public Function Addworkdays(currdate As Date, daystotal As Long) As Date

Do While daystotal > 0
    'if weekend
    If Weekday(currdate) = 1 Or Weekday(currdate) = 7 Then
    currdate = currdate + 1
    ' if weekday
    Else
    currdate = currdate + 1
    daystotal = daystotal - 1
    End If
Loop

Do While daystotal = 0 And (Weekday(currdate) = 1 Or Weekday(currdate) = 7)
    currdate = currdate + 1
Loop
Addworkdays = currdate
End Function

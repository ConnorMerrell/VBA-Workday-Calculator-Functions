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
    'if finishing day is weekend, skip over til weekday
    currdate = currdate + 1
Loop
Addworkdays = currdate
End Function


Public Function DifferenceWorkdays(firstdate As Date, seconddate As Date) As String

Dim a As Long

a = 0

Select Case firstdate - seconddate
    Case Is < 0
    Do While firstdate < seconddate
        firstdate = firstdate + 1
        If Weekday(firstdate) = 1 Or Weekday(firstdate) = 7 Then
        Else
        a = a + 1
        End If
    Loop
    DifferenceWorkdays = a * -1
    
    
    Case Is = 0
    DifferenceWorkdays = "0"
    
    
    Case Is > 0
    Do While firstdate > seconddate
        firstdate = firstdate - 1
        If Weekday(firstdate) = 1 Or Weekday(firstdate) = 7 Then
        Else
        a = a + 1
        End If
    Loop
    
    DifferenceWorkdays = a
    
End Select

'a = firstdate - seconddate
'DifferenceWorkdays = a

End Function

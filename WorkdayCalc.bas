Attribute VB_Name = "WorkdayCalc"
Option Compare Database
Private Function daytest(testeddate As Date) As Boolean
Select Case testeddate
Case "06/05/2024", "27/05/2024"
    daytest = True
Case Else
    Select Case Weekday(testeddate)
    Case 1, 7
        daytest = True
    Case Else
        daytest = False
    End Select
End Select
    
End Function


Public Function Addworkdays(currdate As Date, daystotal As Long) As Date

Do While daystotal > 0
    'if weekend
    If daytest(currdate) = True Then
    currdate = currdate + 1
    ' if weekday
    Else
    currdate = currdate + 1
    daystotal = daystotal - 1
    End If
Loop

Do While daystotal = 0 And daytest(currdate) = True
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
        If daytest(firstdate) = True Then
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
        If daytest(firstdate) Then
        Else
        a = a + 1
        End If
    Loop
    
    DifferenceWorkdays = a
    
End Select

'a = firstdate - seconddate
'DifferenceWorkdays = a

End Function

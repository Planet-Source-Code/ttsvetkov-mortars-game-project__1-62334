Attribute VB_Name = "TickTime"
Option Explicit
'===============================================================================
'Used to calculate the time between two calls of this function (in millisecods).
'Very precise - using GetTickCount api to find the windows time.
'Run ONLY ONCE in a loop!
'===============================================================================
Private Declare Function GetTickCount Lib "kernel32" () As Long

'the last measured time
Private Last As Long

'the differnce
Private Now As Long
Public Function CalculateTime() As Long
If Last > 0 Then
    Now = GetTickCount - Last
Else
    Now = 0
End If

Last = GetTickCount

CalculateTime = Now

End Function

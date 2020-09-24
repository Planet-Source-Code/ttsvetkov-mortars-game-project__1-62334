Attribute VB_Name = "SubMain"
Option Explicit

'the APPATH
Public sApPath As String
Private Sub Main()
'================
'main entry point
'================

'------------------
'finding the appath
'------------------
If Right(App.Path, 1) = "\" Then
    sApPath = App.Path
Else
    sApPath = App.Path & "\"
End If
'----------------------------------------
'checking for another instance of the app
'----------------------------------------
If App.PrevInstance = True Then
    MsgBox App.Title & " is already running.", vbCritical, App.Title
    
    Exit Sub
End If

'----------
'show form1
'----------
On Error Resume Next
Load Form1

End Sub

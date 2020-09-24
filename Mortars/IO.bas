Attribute VB_Name = "IO"
Option Explicit

'APIs
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private PntApi As POINTAPI

'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const KEY_DOWN As Integer = &H1000
Public Function IsKeyDown(KeyCode As Long) As Boolean

If (GetKeyState(KeyCode) And KEY_DOWN) Then
    IsKeyDown = True
Else
    IsKeyDown = False
End If

'If GetAsyncKeyState(KeyCode) <> 0 Then
'    IsKeyDown = True
'Else
'    IsKeyDown = False
'End If

End Function
Public Function GetMouseX() As Integer
If GetCursorPos(PntApi) <> 0 Then
    GetMouseX = PntApi.X
Else
    GetMouseX = 0
End If

End Function
Public Function GetMouseY() As Integer
If GetCursorPos(PntApi) <> 0 Then
    GetMouseY = PntApi.Y
Else
    GetMouseY = 0
End If

End Function

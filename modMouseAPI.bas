Attribute VB_Name = "modMouseAPI"
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type
Public Const MOUSEEVENTF_MOVE = &H1
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_ABSOLUTE = &H8000
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2

'Declare the API-Functions
Private Declare Function GetCursorPos Lib "USER32" (lppoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "USER32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "USER32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub keybd_event Lib "USER32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Sub GetMousePos(x As Long, y As Long)
Dim lppoint As POINTAPI
GetCursorPos lppoint
x = lppoint.x
y = lppoint.y
End Sub
Public Sub SetMousePos(x As Long, y As Long)
SetCursorPos x, y
End Sub

Public Sub MouseDown(x As Long, y As Long, button As Long)
Dim dwFlags As Long
dwFlags = 0
' dwFlags = MOUSEEVENTF_ABSOLUTE
' + MOUSEEVENTF_MOVE
SetMousePos x, y
If button = 1 Then dwFlags = dwFlags Or MOUSEEVENTF_LEFTDOWN
If button = 2 Then dwFlags = dwFlags Or MOUSEEVENTF_RIGHTDOWN
If button = 3 Then dwFlags = dwFlags Or MOUSEEVENTF_MIDDLEDOWN
mouse_event dwFlags, x, y, 0&, 0&
End Sub

Public Sub MouseUp(x As Long, y As Long, button As Long)
Dim dwFlags As Long
dwFlags = 0
SetMousePos x, y
' dwFlags = MOUSEEVENTF_ABSOLUTE
' + MOUSEEVENTF_MOVE
If button = 1 Then dwFlags = dwFlags Or MOUSEEVENTF_LEFTUP
If button = 2 Then dwFlags = dwFlags Or MOUSEEVENTF_RIGHTUP
If button = 3 Then dwFlags = dwFlags Or MOUSEEVENTF_MIDDLEUP
mouse_event dwFlags, x, y, 0&, 0&
End Sub

Public Sub SetKeyDown(keyascii As Long, shift As Integer)
Dim byt As Byte
byt = Asc(Chr$(keyascii))
If shift Then
    keybd_event byt, CByte(0), KEYEVENTF_EXTENDEDKEY, 0
    Else
    keybd_event byt, 0, 0, 0
    End If
End Sub

Public Sub SetKeyUp(keyascii As Long, shift As Integer)
Dim byt As Byte
byt = Asc(Chr$(keyascii))
If shift Then
    keybd_event byt, CByte(0), KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    Else
    keybd_event byt, CByte(0), KEYEVENTF_KEYUP, 0
    End If
End Sub

Public Sub Setkeypress(keyascii As Byte, shift As Integer)
keybd_event CByte(keyascii), CByte(0), CLng(0), CLng(0)
keybd_event CByte(keyascii), CByte(0), KEYEVENTF_KEYUP, CLng(0)
End Sub

Public Sub MouseClick(x As Long, y As Long, button As Long)
Select Case button
    Case 1
     ' Removed MOUSEEVENTF_MOVE + MOUSEEVENTF_ABSOLUTE +
        mouse_event MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP, x, y, 0, 0
    Case 2
        mouse_event MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, x, y, 0, 0
    Case 3
        mouse_event MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MIDDLEUP, x, y, 0, 0
    End Select
End Sub

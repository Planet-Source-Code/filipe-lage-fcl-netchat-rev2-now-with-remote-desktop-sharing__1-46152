VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTerminalWindow 
   AutoRedraw      =   -1  'True
   Caption         =   "Remote desktop window"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      ScaleHeight     =   210
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   5430
      Width           =   9015
      Begin VB.CommandButton Command2 
         Caption         =   "Req. control"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   540
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   1125
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   1770
         TabIndex        =   2
         Top             =   30
         Width           =   45
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7590
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   27977
   End
   Begin VB.Image mousecursor 
      Enabled         =   0   'False
      Height          =   480
      Left            =   5310
      Picture         =   "frmTerminalWindow.frx":0000
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmTerminalWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MousePressing As Integer
Private InControl As Boolean
Private mousex As Long, mousey As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private OutBuffer
Private RemoteHeight As Long
Private RemoteWidth As Long
Public User As String
Public ParentChat As frmChat
Dim buffer As String
Dim CurrentBufferSize As Long

Private Sub Command1_Click()
Winsock1.Close
Unload Me
End Sub

Private Sub Command2_Click()
If ParentChat.MyIP = Winsock1.RemoteHostIP Then
    If MsgBox("Are you sure you wish to request control over your own desktop ?" & vbCrLf & "This may be confusing due to different mouse positions" & vbCrLf & "Answer YES only if you are sure of what you're doing", vbYesNo Or vbDefaultButton2 Or vbExclamation, "Request own control ?") = vbNo Then
        Exit Sub
        End If
    End If
ParentChat.SendPackage "REQUEST_CONTROL" & vbTab & ParentChat.MyIP & vbVerticalTab & ParentChat.Username & vbCrLf, Winsock1.RemoteHostIP
End Sub

Private Sub Form_Click()
If InControl Then
    Winsock1.SendData "MCL=" & mousex & "," & mousey & vbTab
    Else
    Debug.Print "Mouse click on " & mousex & "," & mousey
    End If
End Sub

Private Sub Form_DblClick()
If InControl Then
    Winsock1.SendData "MDBL=" & mousex & "," & mousey & vbTab
    Else
    Debug.Print "Mouse double-click on " & mousex & "," & mousey
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
If InControl Then
    Winsock1.SendData "KDN=" & KeyCode & "," & shift & vbTab
    Else
    Debug.Print "Key down, keycode= " & KeyCode & "," & sjift
    End If
KeyCode = 0: shift = 0
End Sub

Private Sub Form_KeyPress(keyascii As Integer)
' KeyPressed
If InControl Then
    Winsock1.SendData "KPR=" & keyascii & vbTab
    Else
    Debug.Print "Key press, ascii=" & keyascii
    End If
keyascii = 0 ' Clears the keyascii so that this form doesn't do anything with this key
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, shift As Integer)
' Key has been depressed
If InControl Then
    Winsock1.SendData "KUP=" & KeyCode & "," & shift & vbTab
    Else
    Debug.Print "Key up, keycode = " & KeyCode
    End If
KeyCode = 0: shift = 0 ' ditto
End Sub

Private Sub Form_Load()
SetWindowTop Me.hwnd
Me.Caption = "Please wait ... - Connecting to remote computer's desktop"
End Sub

Public Sub ConnectTO(remoteip As String)
Me.Show
Label1 = "Connecting to " & remoteip
Winsock1.Connect remoteip, 27977
End Sub

Private Sub Form_MouseDown(button As Integer, shift As Integer, x As Single, y As Single)
mousex = Me.ScaleX(x, Me.ScaleMode, vbPixels) & ",": mousey = Me.ScaleY(y, Me.ScaleMode, vbPixels)
If InControl Then
    Winsock1.SendData "MDN=" & button & "," & shift & "," & mousex & "," & mousey & vbTab
    Else
    Debug.Print "Mouse down on " & mousex & "," & mousey
    End If
MousePressing = button

End Sub

Private Sub Form_MouseMove(button As Integer, shift As Integer, x As Single, y As Single)
mousex = Me.ScaleX(x, Me.ScaleMode, vbPixels) & ",": mousey = Me.ScaleY(y, Me.ScaleMode, vbPixels)
If InControl Then
    Winsock1.SendData "MMV=" & button & "," & shift & "," & mousex & "," & mousey & vbTab
    Else
    Debug.Print "Mouse@" & mousex & "," & mousey & " (" & button & ")"
    End If

' Mouse position
End Sub

Private Sub Form_MouseUp(button As Integer, shift As Integer, x As Single, y As Single)
mousex = Me.ScaleX(x, Me.ScaleMode, vbPixels) & ",": mousey = Me.ScaleY(y, Me.ScaleMode, vbPixels)
If InControl Then
    Winsock1.SendData "MUP=" & button & "," & shift & "," & mousex & "," & mousey & vbTab
    Else
    Debug.Print "Mouse up on " & mousex & "," & mousey & " (" & button & ")"
    End If
End Sub

Private Sub Winsock1_Close()
Me.Hide
buffer = ""
OutBuffer = ""
MsgBox "Remote side has canceled sharing of desktop", vbExclamation, "Desktop sharing"
Unload Me
End Sub

Private Sub Winsock1_Connect()
Label1.Caption = "Connected to remote desktop. Please wait for frame !"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim a As String
Dim t As Long
Winsock1.GetData a
If CurrentBufferSize = 0 Then
    t = InStr(a, JPGSeparator)
    If t <= 0 Then
        ' What ? No Separator and no size ?... Ignore packet
        Exit Sub
        End If
    GetPicInfo Left(a, t - 1)
    buffer = Mid(a, t + Len(JPGSeparator))
    Label1.Caption = "Receiving " & CurrentBufferSize & " bytes from " & Winsock1.RemoteHost
    Else
    buffer = buffer & a
    If Len(buffer) >= CurrentBufferSize Then
        HandleBuffer
        End If
    End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Me.Hide
If Number = 10061 Then
    MsgBox "Remote computer isn't sharing the desktop", vbExclamation, "Couldn't connect"
    Else
    MsgBox "Winsock error: " & Number & vbCrLf & Description, vbExclamation, "Winsock error"
    End If
Unload Me
End Sub

Private Sub HandleBuffer()
Dim imgdata(0) As String
Dim t As Double
Label1.Caption = "Rendering desktop from " & Winsock1.RemoteHost
t = Timer

Dim dib As New cdibSection
a = FreeFile

Dim jpgbuffer() As Byte
jpgbuffer() = StrConv(buffer, vbFromUnicode)

If LoadJPGFromPtr(dib, VarPtr(jpgbuffer(0)), CurrentBufferSize) Then
    Me.Cls
    dib.PaintPicture Me.hdc
    
    dib.ClearUp
    End If
ReDim jpgbuffer(0)
Erase jpgbuffer()
Me.Caption = "Desktop of " & User

Label1.Caption = "In comunication with remote computer (frame took " & Format(Timer - t, "0.0") & " ms to render )"

CurrentBufferSize = 0

buffer = Mid(buffer, CurrentBufferSize + 1)
tt = InStr(buffer, JPGSeparator)

If tt <= 0 Then
    ' What ? No Separator and no size information ?... Clear the buffer... this shouldn't happen
    buffer = ""
    Exit Sub
    Else
    GetPicInfo Left(buffer, tt - 1)
    buffer = Mid(buffer, tt + Len(JPGSeparator))
    End If
End Sub

Private Sub GetPicInfo(header)
' CurrentBufferSize = Val(Left(buffer, t - 1))
Dim n() As String
n = Split(header, vbTab)
Me.KeyPreview = True

If UBound(n) = 5 Then
    RemoteWidth = Val(n(0))
    RemoteHeight = Val(n(1))
    mbw = Val(n(0)) * 15 + (Me.Width - Me.ScaleWidth)
    mbh = Val(n(1)) * 15 + Picture2.Height + (Me.Height - Me.ScaleHeight)
    If Me.Width < mbw Or Me.Height < mbh Then
        On Error Resume Next ' If maximized, please don't crash :)
        Me.Move Me.Left, Me.Top, mbw, mbh
        End If
    PutMousePointerAt Val(n(2)), Val(n(3))
    InControl = (n(4) = "1")
    mousecursor.Visible = Not InControl
    CurrentBufferSize = Val(n(5))
    End If
End Sub


Private Sub PutMousePointerAt(x As Long, y As Long)
mousecursor.Move Me.ScaleX(x, vbPixels, Me.ScaleMode), Me.ScaleY(y, vbPixels, Me.ScaleMode)
End Sub

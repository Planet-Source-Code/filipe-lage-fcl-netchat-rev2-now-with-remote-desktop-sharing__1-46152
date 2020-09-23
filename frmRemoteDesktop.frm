VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmDesktop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desktop Sharing"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox allowcontrol 
      Caption         =   "Allow users to take control of my desktop (you'll be asked first)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Value           =   1  'Checked
      Width           =   5235
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4710
      TabIndex        =   3
      Top             =   60
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3210
      TabIndex        =   2
      Top             =   60
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock listener 
      Left            =   1590
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   27977
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2130
      Top             =   30
   End
   Begin MSWinsockLib.Winsock client 
      Index           =   0
      Left            =   1110
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Press 'CLOSE' to disable sharing."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   30
      TabIndex        =   1
      Top             =   270
      Width           =   3165
   End
   Begin VB.Label Label1 
      Caption         =   "Currently sharing desktop with other users."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3165
   End
End
Attribute VB_Name = "frmDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const FieldSeparator = vbVerticalTab ' Separator for fields...
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GetDC Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "USER32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private dib As New cdibSection
Public ParentChat As frmChat
Public AllowedIps As String

Public Sub RedrawDesktop()
' Curious fact:
' With the following code:
' SendKeys "{PRTSC}" ' with the wait statement true or false
' me.Picture = Clipboard.GetData() ' with any format (bitmap, wmf, etc)
'
' doesn't work... maybe a windows protection ?
' in fact the clipboard is empty after the sendkeys
' Can someone tell me why this doesn't work ?

Timer1.Enabled = False
t = Timer
Dim xwidth As Long
Dim xheight As Long
Dim XSH As OLE_XSIZE_HIMETRIC

DeskDC& = GetDC(GetDesktopWindow()) ' Get's the DC of the desktop

xwidth = Screen.Width / Screen.TwipsPerPixelX ' Gets the width and height of the screen (in pixels)
xheight = Screen.Height / Screen.TwipsPerPixelY

t2 = Timer
dib.Create xwidth, xheight
BitBlt dib.hdc, 0&, 0&, xwidth, xheight, DeskDC&, 0, 0, vbSrcCopy
' Call's bitblt to store the image in the Dib class
' I could use StrechBlt but it doesn't work in many OS's

Dim bufsize As Long
bufsize = 512000
ReDim buffer(bufsize) As Byte ' Reserve 512k RAM
SaveJPGToPtr dib, VarPtr(buffer(0)), bufsize, 70

t3 = Timer
'
Dim mousex As Long
Dim mousey As Long
GetMousePos mousex, mousey ' Gets the mouse position

' Sends the JPG data to the clients
SendDatatoClients xwidth & vbTab & xheight & vbTab & mousex & vbTab & mousey, StrConv(buffer(), vbUnicode), bufsize

dib.ClearUp ' Clears the DIB... Recover resources
ReDim buffer(0)
Erase buffer()

TimeTook = Timer - t
newinterval = 3 * (TimeTook) * 1000
    ' Adjust the timer to 3 times the rendering time of the desktop picture...
If newinterval > 30000 Then newinterval = 30000
    ' Cant take more than 30 seconds

Timer1.Interval = newinterval
Timer1.Enabled = True
Me.Caption = "Sharing desktop - " & client.UBound & " clients. FPS = " & Format(1 / (newinterval / 1000), "0.00") & " (jpg took " & t3 - t2 & " ms)"
End Sub


Private Sub client_Close(Index As Integer)
' Client has disconnected... Yay... I don't have to send bytes there anymore! :)
client(Index).Tag = ""
client(Index).Close
For o = client.LBound To client.UBound
If client(o).Tag > "" Then Exit Sub
Next
Timer1.Enabled = False
End Sub

Private Sub client_Connect(Index As Integer)
' Got a new client... He's a spy... He wants my desktop!!! ARGHH! Damn you!
Timer1.Enabled = True
client(Index).Tag = "OK"
End Sub

Private Sub client_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Timer1.Enabled = False
client(Index).GetData a$
If InStr(AllowedIps, client(Index).RemoteHostIP & vbTab) Then
    HandleInput a$
    Else
    ' Ignore packet
    End If
Timer1.Enabled = True
End Sub

Private Sub client_SendComplete(Index As Integer)
client(Index).Tag = "OK" ' I'm ready!
' Only after all data is sent, I can send the next frame.
' This way, you can have multiple users seeing your desktop
' and one can have 1 fps and the other 2 fps acording to the remote computer specs
End Sub

Private Sub Command1_Click()
ParentChat.SharingDesktop = False

Me.Caption = "Terminating existing connections"
Timer1.Enabled = False ' Disable the timer. We don't want to send data while we're unloading, right ?

' Closes all open connections (if any)
For o = client.UBound To client.LBound Step -1
client(o).Close
Next
ParentChat.SendPackage "ALERT" & vbTab & ParentChat.MyIP & FieldSeparator & ParentChat.Username & FieldSeparator & "has stopped sharing his desktop" & vbCrLf

Unload Me ' errr... Yep... it means that I'm unloading myself... duh!

End Sub

Private Sub Command2_Click()
Me.Hide ' Also known as Hide Me!
End Sub

Private Sub Form_Load()

listener.Listen ' Activates the winsock to listen to the specified por 27977

Timer1.Enabled = False

SetWindowTop Me.hwnd
PutXPDropShadow Me

End Sub

Private Sub listener_ConnectionRequest(ByVal requestID As Long)
GFC = GetFreeClient
If Timer1.Enabled = False Then
    ' Activate the timer
    Timer1.Enabled = True
    End If
client(GFC).Close
client(GFC).Accept requestID
client(GFC).Tag = "OK"
End Sub

Private Sub Timer1_Timer()
ParentChat.SharingDesktop = True
RedrawDesktop
End Sub

Private Function GetFreeClient()
On Error Resume Next
GetFreeClient = -1
For o = client.LBound To client.UBound
If client(o).State = 0 Then
    GetFreeClient = o
    Exit For
    End If
Next
If GetFreeClient = -1 Then
    GetFreeClient = client.UBound + 1
    Load GetFreeClient
    End If
End Function

Private Sub SendDatatoClients(header As String, data As String, datsize As Long)
For o = client.LBound To client.UBound
If client(o).State = 7 Then ' Connected
    If client(o).Tag = "OK" Then ' It's free to send data
        
        client(o).Tag = "SENDING" ' Marks as 'SENDING'. It will be OK once all data is transmited
        If InStr(AllowedIps, client(o).RemoteHostIP & vbTab) Then
            client(o).SendData header & vbTab & "1" & vbTab & datsize & JPGSeparator & Left(data, datsize)
            Else
            client(o).SendData header & vbTab & "0" & vbTab & datsize & JPGSeparator & Left(data, datsize)
            End If
        ' Sends the datsize along with the data spaced by a null character
        ' The buffer can handle it :)
        End If
    End If
Next
End Sub


Private Function ReadFullFile(f) As String ' Reads ALL bytes from a file in HD to a string in memory
' Fast method... Yep... Many MBytes / s without api call...
a = FreeFile
Open f For Binary As #a
ReadFullFile = Space(LOF(a))
Get #a, , ReadFullFile
Close #a
End Function

Private Sub HandleInput(buf As String)
Dim c() As String
c = Split(buf, vbTab)
Dim d() As String
kpr = ""
For o = LBound(c) To UBound(c)
d = Split(c(o), "=")
If UBound(d) < 1 Then Exit For
e = Split(d(1) & ",,,,,", ",") ' Make sure it has many args so the following cmds don't crash in case of under-buffer
Select Case d(0)
    Case "MCL"
        ' Mouse Click
        modMouseAPI.MouseClick Val(e(0)), Val(e(1)), 1
    Case "MDBL"
        ' Mouse double click
        modMouseAPI.MouseClick Val(e(0)), Val(e(1)), 1
        ' The click event is also raised in double-click... so the following line
        ' isn't needed
'        modMouseAPI.MouseClick Val(e(0)), Val(e(1)), 1
    Case "MDN"
        ' Mouse button pressed
        ' modMouseAPI.MouseDown Val(e(2)), Val(e(3)), Val(e(0))
        ' gotmouse = False ' No need to move the mouse to the pos
    Case "MUP"
        ' Mouse button depressed
        ' modMouseAPI.MouseUp Val(e(2)), Val(e(3)), Val(e(0))
        ' gotmouse = False
    Case "MMV"
        ' Mouse moved
        cmx = Val(e(2))
        cmy = Val(e(3))
        gotmouse = True
    Case "KDN"
        ' Key down
        ' SetKeyDown Val(e(0)), Val(e(1))
    Case "KUP"
        ' SetKeyUp Val(e(0)), Val(e(1))
    Case "KPR"
'        kpr = kpr & Chr$(Val(e(0)))
        ' SetKeyDown Val(e(0)), Val(e(1))
        SendKeys Chr$(e(0))
        ' SetKeyDown Val(e(0)), Val(e(1))
        ' SetKeyUp Val(e(0)), Val(e(1))
    End Select
Next
If gotmouse = True Then
    modMouseAPI.SetMousePos CLng(cmx), CLng(cmy)
    End If
End Sub

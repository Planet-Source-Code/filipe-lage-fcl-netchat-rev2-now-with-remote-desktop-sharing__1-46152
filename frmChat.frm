VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Begin VB.Form frmChat 
   Caption         =   "Network Chat - (c)2003 Filipe Camizão Lage"
   ClientHeight    =   5295
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8640
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8160
      Top             =   690
   End
   Begin MSWinsockLib.Winsock filesock 
      Left            =   8130
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   8130
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   8160
      Top             =   1140
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5040
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   9208
            Text            =   "In chat"
            TextSave        =   "In chat"
            Key             =   "status"
            Object.ToolTipText     =   "Current status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Users online:"
            TextSave        =   "Users online:"
            Key             =   "users_online"
            Object.ToolTipText     =   "Number of users currently in chat"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "#"
            TextSave        =   "#"
            Key             =   "ip"
            Object.ToolTipText     =   "This computer's IP"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   159
            MinWidth        =   2
            Key             =   "upload"
            Object.ToolTipText     =   "Upload activity"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   159
            MinWidth        =   2
            Key             =   "download"
            Object.ToolTipText     =   "Download activity"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock UDPListen 
      Left            =   8130
      Top             =   2970
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   27976
   End
   Begin MSWinsockLib.Winsock broadcast 
      Left            =   8130
      Top             =   3420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "0.0.0.0"
      RemotePort      =   27976
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8070
      Top             =   3990
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":058A
            Key             =   "user"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":0B24
            Key             =   "op"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":0E3E
            Key             =   "u1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1718
            Key             =   "u2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1FF2
            Key             =   "u3"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":28CC
            Key             =   "u4"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":31A6
            Key             =   "idle2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3740
            Key             =   "idle"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox sendtext 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   4650
      Width           =   6315
   End
   Begin RichTextLib.RichTextBox chat 
      Height          =   4335
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   7646
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmChat.frx":3CDA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView users 
      Height          =   6585
      Left            =   6300
      TabIndex        =   0
      Top             =   0
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   11615
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   12763842
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      Picture         =   "frmChat.frx":3D5E
   End
   Begin MSWinsockLib.Winsock filesock2 
      Left            =   8130
      Top             =   2070
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnu_chat 
      Caption         =   "Chat"
      Begin VB.Menu mnu_showwindow 
         Caption         =   "Show window"
      End
      Begin VB.Menu mnu_hidewindow 
         Caption         =   "Hide to system tray"
      End
      Begin VB.Menu mnu_blankk 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_sharedesktop 
         Caption         =   "Share my desktop"
      End
      Begin VB.Menu mnu_blank0A 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_refreshusers 
         Caption         =   "Refresh users"
      End
      Begin VB.Menu mnu_actionmessage 
         Caption         =   "Send communication message..."
      End
      Begin VB.Menu mnu_selectnick 
         Caption         =   "Change my nick name..."
      End
      Begin VB.Menu mnu_wake 
         Caption         =   "If one user tries to call me"
         Begin VB.Menu mnu_waketype 
            Caption         =   "Ignore him"
            Index           =   0
         End
         Begin VB.Menu mnu_waketype 
            Caption         =   "Popup this window"
            Index           =   1
         End
      End
      Begin VB.Menu mnu_blank00 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_About 
         Caption         =   "About"
      End
      Begin VB.Menu mnu_blank01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnu_User 
      Caption         =   "UserMenu"
      Begin VB.Menu mnu_Info 
         Caption         =   "Information about this user"
      End
      Begin VB.Menu mnu_blank0 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_slap 
         Caption         =   "Slap"
         Index           =   0
      End
      Begin VB.Menu mnu_slap 
         Caption         =   "Slap without him knowing it"
         Index           =   1
      End
      Begin VB.Menu mnu_blank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_sendfile 
         Caption         =   "Send file"
      End
      Begin VB.Menu mnu_viewdesktop 
         Caption         =   "View this user's desktop"
      End
      Begin VB.Menu mnu_blank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_popremote 
         Caption         =   "Wake him up!"
      End
      Begin VB.Menu mnu_sendalert 
         Caption         =   "Send alert message"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 13/June/2003 13:05 - Begining of this project
' 13/June/2003 17:11 - First release. Yep... took 4 hours and 6 minutes of my limited life time.
' 16/June/2003 17:24 - Second release... With idle status, advanced user information and desktop viewing
'
' Hi everyone.
'
'  This is just a small, simple, and commented project for network chat, without using a server
'
'  It's a small gift and token of my appreciation for www.planet-souce-code.com, since i have consulted
' PSC many times for info and code examples, and never submited (or contributed) any of my code.
'
' So, here is my contribution. Hope it's useful for someone.
'
' Thank you PSC! :)
'                                                       The author:
'                                                   Filipe Camizão Lage
'                                                   fclage@mail.net4b.pt

' You may use this code freely in your projects, you can change it, etc, etc, yabayaba, but please
' do not remove my name from the credits.

' Small detail about this code. I use some functions as:
'           * Some API calls (window on top, XP shadow, Get Windows User, Computer Name, etc)
'           * Load data/pictures from resource file
'           * Listview for user listing
'           * RTF textbox (RichTextFormat) from Microsoft
'           * Winsock control from Microsoft
'           * How to use the IIF function
'           * Sending files across network using TCP
'           * Sending broadcast messages across network
'           * Defining a protocol and using UDP to receive broadcasts
'           * Using a buffer to store (possible) incomplete commands
'           * Virtually unlimited users in this chat
'               (altough it may be limited to network bandwidth and memory)
'           * Using RTF with different colors, fontsizes, bold, etc
'           * Using pictures in the statusbar from microsoft to show the Download/Upload status
'           * It is possible to add an external IP for the conversation, but I leave that to you ;)
'           * Many new functions can be added, but I also leave that to you. (ex: Advanced user information)

' New stuff added since last release:
'           * Advanced remote user information (screen resolution, printer, operating system, etc, etc)
'           * Multiple Remote Desktop Viewing !!!!
'               many clients can connect to a single desktop and 'view' it... Sorry, no interaction with remote desktop yet.
'           * Call sleeping user (user that has minimized the application)
'           * Other added feats

Const BlockSize = 8192
Const HideWhenMinimized = True

Private WithEvents MySysTray As cSystray
Public SendForm As frmProgress
Public ReceiveForm As frmProgress
Public CurrentBlock As Long
Public SendingFile As String
Public Username As String     ' My username (each computer will have one)
Public CommBuffer As String   ' Communication buffer to be used. See UDPListen_Dataarrival for more info
Const FieldSeparator = vbVerticalTab ' Separator for fields...
                              ' Since this char. will not be used in text exchange (messages) from users
                              ' we can use it to separate multiple fields for one command ... Example: Logon
Public SelectedUser As ListItem ' Current selected user from list
Public MyIcon As String        ' by default it will use icon 'user' except for yourself
Private CurrentlyIdle As Boolean ' See the mnu_waketype for more details
Public SharingDesktop As Boolean ' Just a boolean for functions to check if the desktop is currently shared
Public Enum SysTray_Activity
    NoActivity = 0
    ChatInProgress = 1
    Warning = 2
    End Enum
Public Enum MSGType
    msgNormal = 0
    msgInformation = 1
    msgWarning = 2
    msgAlert = 3
    msgCritical = 4
    msgFromUser = 10
    End Enum
Private HiddenActivityAction As Integer

Private Sub Form_Load()
mnu_User.Visible = False
UDPListen.LocalPort = 27976   ' Hey! It's my birthday! 27-Set-1976 :)
broadcast.RemotePort = 27976  ' Use should use the same port to send the messages to other users
broadcast.RemoteHost = "255.255.255.255"  ' Broadcast IP - All network
UDPListen.Bind                ' UDP must be bound to a port... We shouldn't command 'Listen' like a TCP connection
broadcast.Bind                ' The broadcast object must be bound also... The local port will be dinamic
                              ' (one that's free)
CurrentBlock = -1 ' Initialize the file transfer block. -1 = No file being sent

Username = modChat.GetWindowsUser ' By default, it will use the windows login

If Username = "" Then
    ' But if it's win9x w/o network, there's no username...
    ' so let's use computer name
    Username = modChat.GetWindowsUser
    End If
If Username = "" Then Username = "Anonymous" ' Still no user ? Ok...It's anonymous :P
MyIcon = "user" ' My icon visible to other users
SetActivity "U- D-" ' No download, no upload... Reset pictures
mnu_waketype_Click 1 ' Set's wake up method
Set MySysTray = New cSystray
MySysTray.Initialize Me.hWnd, Me.Icon, "FCL Network Chat"
MySysTray.ShowIcon

setSystrayActivity NoActivity

status.Panels("ip").Text = MyIP

chat.Text = ""
GetSystemInfo
Logon
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msgCallBackMessage As Long
msgCallBackMessage = X / Screen.TwipsPerPixelX
Dim WM_RBUTTONDOWN As Long
Dim WM_RBUTTONUP As Long
Dim WM_LBUTTONDBLCLK As Long

WM_LBUTTONDBLCLK = &H203
WM_RBUTTONDOWN = &H204
WM_RBUTTONUP = &H205

Select Case msgCallBackMessage
    Case WM_RBUTTONUP
        Me.PopupMenu Me.mnu_chat
    Case WM_LBUTTONDBLCLK
        mnu_showwindow_Click
    End Select
End Sub

Private Sub Form_Resize()
Dim ymax As Long

' Resizes the objects in this form
On Error Resume Next ' We don't want any errors when the user minimizes the window

If Me.WindowState = 1 Then
    Me.Caption = Replace(Me.Caption, " - (c)2003 Filipe Camizão Lage", "")
    ' If minimizing, remove the author's credits so that it doesn't take much space in the taskbar
        
    If HideWhenMinimized Then Me.Hide
        
    Else
    If InStr(Me.Caption, " - (c)2003 Filipe Camizão Lage") <= 0 Then Me.Caption = Me.Caption & " - (c)2003 Filipe Camizão Lage"
    ' If not, put them back
    If Me.Visible = False Then
        Me.Visible = True
        End If
    setSystrayActivity NoActivity
    End If

Me.mnu_showwindow.Enabled = (Me.Visible = False)
Me.mnu_hidewindow.Enabled = (Me.Visible = True)

If Me.WindowState = vbMinimized Then
    SendPackage "IDLE" & vbTab & MyIP & vbCrLf
    CurrentlyIdle = True
    MyIcon = "idle"
    Else
    If CurrentlyIdle Then
        SendPackage "NOTIDLE" & vbTab & MyIP & vbCrLf
        CurrentlyIdle = False
        MyIcon = "user"
        End If
    End If

ymax = ScaleHeight - sendtext.Height - status.Height
users.Left = Me.ScaleWidth - users.Width - 15
users.Height = ymax
chat.Width = users.Left - 15
chat.Height = ymax
sendtext.Move 0, ymax, ScaleWidth
End Sub

Public Sub Logon()
' Logon and send your 'I'm here' to the network
SendPackage "LOGON" & vbTab & MyIP & FieldSeparator & Username & FieldSeparator & MyIcon & vbCrLf

' Request a ping from all users... They'll report their nicks and ip's... Just to populate the userlist
RequestPing
End Sub

Private Sub LogOff()
SendPackage "LOGOFF" & vbTab & MyIP & vbCrLf
End Sub

Private Sub Form_Unload(Cancel As Integer)
' UNLOAD ALL WINDOWS EXCEPT ME... I'll unload myself thank you
MySysTray.HideIcon
Set MySysTray = Nothing

Dim n As Form
For Each n In VB.Forms
If n.Name <> "frmChat" Then
    Unload n
    End If
Next
LogOff ' Logoff (other users will se you leaving)
Set SelectedUser = Nothing
End Sub

Public Sub HandleCommand(exec As String, dat As String)
Dim n As ListItem
On Error Resume Next
Dim d() As String
d = Split(dat, FieldSeparator)
Select Case exec
    Case "LOGON", "PTOING"
        On Error Resume Next ' If error exists , then that's because the user is already there... So no problem!
        Set n = Nothing
        Set n = users.ListItems("ip" & d(0)) ' Retrieve the user from that IP
        If n Is Nothing Then
            If exec = "LOGON" Then
                ' Add information that this user has entered chat (logon).
                AddLine "User " & d(1) & " has entered this conversation", msgAlert
                ' No need to show in PTOING reply
                End If
            Set n = users.ListItems.Add(, "ip" & d(0)) ' I haven't got it in my list... adding it
            End If
        n.Text = d(1)
        If d(0) = MyIP Then ' Hey... it's me... let's put a diff. icon just to distinguish myself :)
            n.SmallIcon = "u1"
            Else
            n.SmallIcon = d(2)
            End If
        If n.SmallIcon = "" Then n.SmallIcon = "user" ' That icon was not found... using default
        RefreshUserCount
    Case "PING"
        ' Since the other user is loggin in, i'll just send the information that I'm here and ready to chat with him :)
        SendPackage "PTOING" & vbTab & MyIP & FieldSeparator & Username & FieldSeparator & MyIcon & vbCrLf, d(0)
        On Error Resume Next
        Set n = Nothing ' Just checking if user changed nick
        Set n = users.ListItems("ip" & d(0)) ' Retrieve the user from that IP
        If n Is Nothing Then
            Else
            If n.Text <> d(1) Then
                AddLine "User " & n.Text & " is now " & d(1), msgWarning
                n.Text = d(1)
                End If
            End If
        ' Just a variant of PING-PONG... It's now PING, PTOING... It sounds more funny :)
    Case "MSG"
        Add2Chat d()
        ' We got a message! Let's put it's handling in a sub called Add2Chat :)
    Case "IDLE"
        users.ListItems("ip" & d(0)).Ghosted = True
        users.ListItems("ip" & d(0)).Tag = users.ListItems("ip" & d(0)).SmallIcon
        users.ListItems("ip" & d(0)).SmallIcon = "idle"
    Case "NOTIDLE"
        users.ListItems("ip" & d(0)).Ghosted = False
        users.ListItems("ip" & d(0)).SmallIcon = users.ListItems("ip" & d(0)).Tag
    Case "POPIT"
        AddLine "User " & d(0) & " requires your attention", msgCritical
        Select Case HiddenActivityAction
            Case 0
                ' Do nothing... I don't want to be disturbed
            Case 1
                Me.WindowState = 0
                Me.Show
            Case 2
                Me.SetFocus
            End Select
    Case "ALERT"
        AddLine d(1) & " " & d(2), msgWarning, " "
    Case "MSGBOX"
        MsgBox d(2), vbInformation Or vbOKOnly, "Message from: " & d(1) & " [" & d(0) & "]"
    Case "SLAP"
        ' SLAP SLAP SLAP! Take that you nasty little rat! :)
        If MyIP = d(2) Then
            If MyIP = d(0) Then
                AddLine "You're slapping yourself... damn you're mean!", msgAlert
                Else
                AddLine d(1) & " is slapping YOU!", msgAlert
                End If
            Else
            AddLine d(1) & " is slapping " & IIf(d(0) = d(2), "HIMSELF! LoL!", d(3)) & ".", msgAlert
            End If
    Case "ANONYMOUS_SLAP"
        ' Yep! It's anonymous! You don't know who slapped you... There's no IP, no user, no nothing...
        ' Of course, you know who is if
        ' a) You're slapping yourself
        ' b) You're slapping the other user in the chat when there's just you and him.
        AddLine "Someone is anonymously slapping " & d(1) & ".", msgCritical
    Case "REQUEST_INFO"
        AddLine "User " & d(1) & " requested information about your computer", msgInformation, " "
        SendPackage "MYINFO" & vbTab & MyIP & FieldSeparator & Username & FieldSeparator & GetSystemInfo & vbCrLf, d(0)
    Case "MYINFO"
        Dim kk As New frmSysinfo
        kk.HandleInformation d()
        kk.Show
    Case "SENDFILE"
        If MsgBox("The user " & d(1) & " wishes to send you one file: " & vbCrLf & d(3) & vbCrLf & d(2) & " bytes" & vbCrLf & vbCrLf & "Do you accept this file ?", vbYesNo Or vbQuestion, "Incoming transfer") = vbNo Then
            SendPackage "REFUSEFILE" & vbTab & Username & FieldSeparator & d(2) & vbCrLf, d(0)
            Else
            Dim FileToWrite As String
            
            ' Save the file to the application's path
            FileToWrite = App.Path
            If Right(FileToWrite, 1) <> "\" Then FileToWrite = FileToWrite & "\"
            FileToWrite = FileToWrite & d(3)
            
            ' You could also use the common dialog to save the file.
            ' Just uncomment the following code
            
            ' cdlg.CancelError = False
            ' cdlg.DialogTitle = "Save incoming file from user " & d(1) & " to..."
            ' cdlg.filename = d(3)
            ' cdlg.ShowSave
            ' If cdlg.filename = "" Then
            '     SendPackage "REFUSEFILE" & vbTab & Username & FieldSeparator & d(2) & vbCrLf, d(0)
            '     Exit Sub
            '     End If
            ' FileToWrite = cdlg.filename
            
            Set ReceiveForm = New frmProgress
            ReceiveForm.Tag = ""
            ReceiveForm.SetInfo d(3), d(2)
            ReceiveForm.SetProgress 0
            ReceiveForm.Show
            filesock.Close
            filesock.Listen
            filesock.Tag = d(2) & vbTab & FileToWrite
            SendPackage "ACCEPTFILE" & vbTab & MyIP & FieldSeparator & filesock.LocalPort & FieldSeparator & d(3) & vbCrLf, d(0)
            ' Sends the IP and Port for the sender to connect
            
            On Error Resume Next
            ' Also, let's make sure there isn't a previous file here
            Kill FileToWrite
            End If
    Case "REFUSEFILE"
        AddLine "User " & d(0) & " has refused your nasty, little and insignificant file. Oh well", msgWarning, " "
        SendingFile = ""
    Case "ACCEPTFILE"
        ' File has been accepted... Let's send it!
        Set SendForm = New frmProgress
        SendForm.Tag = ""
        SendForm.SetInfo SendingFile, FileLen(SendingFile)
        SendForm.SetProgress 0
        SendForm.Show
        filesock2.Connect d(0), Val(d(1)) ' Try to connect to the IP and Port specified in the reply
    Case "REQUEST_CONTROL"
        If SharingDesktop Then
            If frmDesktop.allowcontrol.Value = 1 Then
                If MsgBox("The user " & d(1) & " (" & d(0) & ") is requesting control over your desktop." & vbCrLf & _
                    "Do you want to give him control to your desktop ?" & vbCrLf & vbCrLf & _
                    "(You can allways disconnect him at anytime disabling the desktop sharing)", _
                    vbYesNo Or vbQuestion, "Remote desktop control") = vbNo Then
                        ' Didn't accept control
                        SendPackage "CONTROL_DENIED" & vbTab & MyIP & FieldSeparator & Username & vbCrLf, d(0)
                        Else
                        ' Accepted control
                        SendPackage "CONTROL_ACCEPTED" & vbTab & MyIP & FieldSeparator & Username & vbCrLf, d(0)
                        
                        ' Check if the remote ip is authorized... If now, add it to the auth. list
                        If InStr(frmDesktop.AllowedIps, d(0) & vbTab) <= 0 Then
                            frmDesktop.AllowedIps = frmDesktop.AllowedIps & d(0) & vbTab
                            End If
                        End If
                Else
                SendPackage "CONTROL_DENIED" & vbTab & MyIP & FieldSeparator & Username & vbCrLf, d(0)
                End If
            End If
    Case "CONTROL_DENIED"
        MsgBox "User " & d(1) & " has not accepted your request to control his desktop", vbExclamation, "Request denied"
    Case "CONTROL_ACCEPTED"
        ' No need for code here...
    Case "LOGOFF"
        On Error Resume Next ' If error exists , then that's because the user is not here... So no problem!
        Set n = users.ListItems("ip" & d(0))
        If n Is Nothing Then
            Else
            AddLine "User " & n.Text & " left this conversation", msgAlert, " "
            users.ListItems.Remove n.Index
            End If
        RefreshUserCount
    End Select
End Sub
Public Sub Add2Chat(d() As String)
If d(0) = MyIP Then
    ' Hey! It's the message I've just send to the network... I know I heard it... Should I show it to myself ?
    ' if not, please un-comment the "exit sub" in the following line.
    ' Exit Sub
    End If
Dim UName As String
UName = users.ListItems("ip" & d(0)).Text
If UName = "" Then UName = "Anonymous"
AddLine d(1), msgFromUser, UName
End Sub


Private Sub mnu_About_Click()
frmAbout.Show
End Sub

Private Sub mnu_actionmessage_Click()
Dim msg As String
msg = InputBox("What is the message I should send to the other user ?", "Send action", "")
If msg = "" Then
    Exit Sub ' User has canceled
    End If
Dim remoteip As String
If SelectedUser Is Nothing Then
    remoteip = "" ' Send to all (yourself included)
    Else
    remoteip = Replace(SelectedUser, "ip", "")
    End If
SendPackage "ALERT" & vbTab & MyIP & FieldSeparator & Username & FieldSeparator & msg & vbCrLf, remoteip
End Sub

Private Sub mnu_hidewindow_Click()
Me.WindowState = 1
mnu_showwindow.Enabled = True
mnu_hidewindow.Enabled = False
End Sub

Private Sub mnu_Info_Click()
' Request information on remote computer
SendPackage "REQUEST_INFO" & vbTab & MyIP & FieldSeparator & Username & vbCrLf, Replace(SelectedUser.Key, "ip", "")
End Sub

Private Sub mnu_popremote_Click()
' Asks for the remote user to pop it's window
SendPackage "POPIT" & vbTab & Username & FieldSeparator & MyIP & vbCrLf, Replace(SelectedUser.Key, "ip", "")
End Sub

Private Sub mnu_quit_Click()
Unload Me
End Sub

Private Sub mnu_refreshusers_Click()
users.ListItems.Clear
RequestPing
AddLine "User list refreshed", msgCritical, " "
End Sub

Private Sub mnu_selectnick_Click()
Dim newnick As String
newnick = InputBox("New nickname ?", App.Title, Username)
If newnick = "" Or newnick = Username Then
    Exit Sub ' No need to change since user has canceled or selected the same nick
    End If

' If you do not want users to use names bigger than 16 characters
' you can uncomment the line below
' newnick = Left(Trim(newnick), 16)

Username = newnick
RequestPing ' Informs users of my new name

End Sub

Private Sub mnu_sendalert_Click()
Dim alertmsg As String
alertmsg = InputBox("What is the message I should send to the other user ?", "Send alert", "")
If alertmsg = "" Then
    Exit Sub ' User has canceled
    End If
SendPackage "MSGBOX" & vbTab & MyIP & FieldSeparator & Username & FieldSeparator & alertmsg & vbCrLf, Replace(SelectedUser.Key, "ip", "")
End Sub

Private Sub mnu_sendfile_Click()
On Error GoTo Canceled
cdlg.CancelError = True
cdlg.Filter = "All files|*.*"
' If you want to reset the path everytime the user tries to send a file, uncomment the following line:
' cdlg.InitDir = App.Path
cdlg.ShowOpen
If cdlg.filename > "" Then
    SendFilename Replace(SelectedUser.Key, "ip", ""), cdlg.filename
    End If
Canceled:
Exit Sub

End Sub

Private Sub mnu_sharedesktop_Click()
ShareDesktop True
End Sub

Private Sub mnu_showwindow_Click()
If Me.WindowState = 1 Then Me.WindowState = 0 ' If minimized, then restore
If Me.Visible = False Then Me.Visible = True ' If invisible, the show yourself... You can put this in the previous line
                                            ' to enable the windows restore animation
mnu_showwindow.Enabled = False
mnu_hidewindow.Enabled = True
End Sub

Private Sub mnu_slap_Click(Index As Integer)
If Index = 0 Then
    ' Slaps the user... But he'll know it's from you
    SendPackage "SLAP" & vbTab & _
        MyIP & FieldSeparator & _
        Username & FieldSeparator & _
        Replace(SelectedUser.Key, "ip", "") & FieldSeparator & _
        SelectedUser.Text & vbCrLf
    Else
    ' Anonymous slap :)
    SendPackage "ANONYMOUS_SLAP" & vbTab & _
        Replace(SelectedUser.Key, "ip", "") & FieldSeparator & _
        SelectedUser.Text & vbCrLf
    End If

End Sub

Private Sub mnu_viewdesktop_Click()

Dim ThisIP As String
Dim n As Form

' Checks if a existing window of the user's desktop is already loaded
ThisIP = Replace(SelectedUser.Key, "ip", "")
For Each n In VB.Forms
If TypeOf n Is frmTerminalWindow Then
    If n.Tag = ThisIP Then
        n.SetFocus ' Yep! I have his desktop right here! No need for another connect
        Exit Sub
        End If
    End If
Next

' Creates the terminal window and connects to the computer sharing the desktop
Dim k As New frmTerminalWindow
k.Tag = ThisIP
' k.Height = n.Height - n.ScaleHeight ' Make zero height... Until it get's the desktop image
k.User = SelectedUser.Text
Set k.ParentChat = Me
k.Show
k.ConnectTO ThisIP

End Sub

Private Sub mnu_waketype_Click(Index As Integer)
mnu_waketype(Index).Checked = True
mnu_waketype(1 - Index).Checked = False
HiddenActivityAction = Index
                                ' 0=Do nothing
                                ' 1=PopupWindow
                                ' 2=Flash in taskbar

' I know... index 2 isn't in the menu... I haven't tested it yet...
End Sub

Private Sub Timer1_Timer()
SetActivity "U- D- " ' Set upload download activity back to none
' Just a timer of some miliseconds to set Upload and Download colors back to green
Timer1.Enabled = False
End Sub

' --------------------------------------------------------------------------------------------
' -[COMMUNICATION]----------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------
Private Sub UDPListen_DataArrival(ByVal bytesTotal As Long)
Dim dat As String
SetActivity "D+"

UDPListen.GetData dat         ' Get's the data (in string) from the winsock UDP buffer

CommBuffer = CommBuffer & dat ' Since UDP packets can be fragmented, we'll add it to a buffer
                              ' Only the buffer will be handled. This way, there's no problem if
                              ' the packet is fragmented. The handler will check for package integrity

HandleBuffer                  ' Sub to handle the buffer and do the appropriate actions
End Sub

Public Sub HandleBuffer()
Dim Packets() As String
Packets = Split(CommBuffer, Chr$(255)) ' Split packets from commbuffer.

            ' There can be multiple packets in the buffer...
            ' Handle them in order of arrival

Dim o As Integer
Dim Cmd As String ' Command
Dim dat As String ' Data for that command
Dim Pnt As Long
Dim FPos As Long

Dim PacketIncomplete As Boolean ' Is the packet still incomplete ?
Dim PacketDamaged As Boolean
Dim PackSize As Integer

For o = LBound(Packets) To UBound(Packets)
If Right(Packets(o), 2) = vbCrLf Then
    PacketIncomplete = False ' Let's consider everything is ok for now
    PacketDamaged = False    ' There's nothing damaged so far
    If Left(Packets(o), 1) <> Chr$(0) Then PacketDamaged = True
    If InStr(Packets(o), Chr(0)) <= 0 Then PacketIncomplete = True
    If InStr(Packets(o), Chr(1)) <= 0 Then PacketIncomplete = True
    If InStr(Packets(o), vbTab) <= 0 Then PacketIncomplete = True
    If PacketIncomplete = False Then
        FPos = InStr(Packets(o), Chr(1))
        Pnt = InStr(Packets(o), vbTab)
        PackSize = Val(Mid$(Packets(o), 2, FPos - 2))
        Cmd = Mid(Packets(o), FPos + 1, Pnt - FPos - 1)
        dat = Mid(Packets(o), Pnt + 1, PackSize - Len(Cmd) - 3)
        HandleCommand Cmd, dat
        Packets(o) = ""
        End If
    End If
Next

' Recreate the CommBuffer from the packets splitted
CommBuffer = Join(Packets, Chr$(255))

' Removes the empty-and-already-handled packets
Do Until InStr(CommBuffer, Chr$(255) & Chr$(255)) <= 0
CommBuffer = Replace(CommBuffer, Chr$(255) & Chr$(255), Chr$(255))
Loop
If Replace(CommBuffer, Chr$(255), "") = "" Then CommBuffer = "" ' Let's empty the buffer

End Sub

Public Sub SendPackage(data_to_send As String, Optional SpecificIPs As String = "")
 ' A simple protocol is used to transmit the data, so that the remote client can check for integrity
 ' Example:
 ' [ASCII00]12[ASCII01]COMMAND[VBTAB]DATA[VBCRLF][ASCII255]

SetActivity "U+" ' Currently uploading stuff

If SpecificIPs = "" Then
    broadcast.RemoteHost = "255.255.255.255"
    broadcast.SendData CStr(Chr$(0) & Len(data_to_send) & Chr$(1) & data_to_send & Chr$(255))
    Else
    Dim ips() As String
    Dim o As Integer ' Sending to 32.767 ip's max ... That's more than enough. Usually this function will be used just for one user
    ips = Split(SpecificIPs, vbTab)
    For o = LBound(ips) To UBound(ips)
    broadcast.RemoteHost = ips(o)
    broadcast.SendData CStr(Chr$(0) & Len(data_to_send) & Chr$(1) & data_to_send & Chr$(255))
    Next
    End If
End Sub
Property Get MyIP() As String
MyIP = broadcast.LocalIP
End Property


' --------------------------------------------------------------------------------------------
' -[OBJECTS HANDLING]-------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------
Private Sub users_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set SelectedUser = users.HitTest(X, Y)
If SelectedUser Is Nothing Then
    mnu_User.Visible = False
    Exit Sub
    End If
mnu_User.Caption = "User " & SelectedUser.Text
mnu_User.Visible = (SelectedUser.Key <> "ip" & MyIP)
' Just show the menu for other users. Not yourself
' If you want to see the menu for yourself uncomment the following line
mnu_User.Visible = True

SelectedUser.Selected = True

If Button = 2 Then PopupMenu mnu_User
End Sub
Private Sub sendtext_KeyPress(keyascii As Integer)
If keyascii = 13 Then
    If Left(LCase(sendtext), 4) = "/me " Then
        SendPackage "ALERT" & vbTab & MyIP & FieldSeparator & Username & FieldSeparator & Mid(sendtext.Text, 5) & vbCrLf
        Else
        SendPackage "MSG" & vbTab & MyIP & FieldSeparator & sendtext.Text & vbCrLf
        End If
    sendtext = ""
    End If
End Sub

Public Sub AddLine(msg As String, Optional mtype As MSGType, Optional UName As String)
chat.SelStart = Len(chat.Text)
If UName > "" Then
    If Trim(UName) > "" Then
        chat.SelColor = vbBlue
        chat.SelBold = True
        chat.SelFontSize = 7
        chat.SelText = "[" & Format(Now, "hh:mm") & "]"
        chat.SelStart = Len(chat.Text)
        chat.SelColor = vbBlack
        chat.SelFontSize = 8
        chat.SelText = "<" & UName & ">"
        chat.SelBold = False
        chat.SelStart = Len(chat.Text)
        
        If Me.Visible = False Then
            ' If this chatwindow isn't visible (ex: in systray) then
            ' activates the timer show that the user knows
            ' someone is talking behind his back :)
            setSystrayActivity ChatInProgress
            End If
        End If
    Else
    chat.SelColor = vbBlue
    chat.SelFontSize = 7
    chat.SelBold = True
    chat.SelText = "[" & Format(Now, "hh:mm") & "]"
    chat.SelBold = False
    chat.SelStart = Len(chat.Text)
    End If
    
Select Case mtype
    Case msgInformation
        chat.SelBold = False
        chat.SelFontSize = 7
        chat.SelColor = RGB(128, 128, 128)
    Case msgWarning
        chat.SelBold = True
        chat.SelFontSize = 8
        chat.SelColor = RGB(0, 128, 0)
    Case msgCritical
        chat.SelBold = True
        chat.SelFontSize = 8
        chat.SelColor = RGB(128, 32, 0)
    Case msgAlert
        chat.SelBold = False
        chat.SelFontSize = 7
        chat.SelColor = RGB(0, 32, 128)
    Case msgFromUser
        chat.SelBold = False
        chat.SelFontSize = 8
        chat.SelColor = vbBlack
    Case Else
        chat.SelBold = False
        chat.SelFontSize = 8
        chat.SelColor = vbBlack
    End Select

' Adds the message to the chat
chat.SelText = msg & vbCrLf
chat.SelStart = Len(chat.Text)
End Sub

Public Sub RequestPing()
SendPackage "PING" & vbTab & MyIP & FieldSeparator & Username & vbCrLf
End Sub

Public Sub SetActivity(Activity As String)
Dim A1 As String
Dim A2 As String
A1 = ""
A2 = ""
If InStr(Activity, "U+") Then A1 = "red": Timer1.Enabled = True  ' It's uploading stuff
If InStr(Activity, "U-") Then A1 = "green"                       ' It has stopped upload
If InStr(Activity, "D+") Then A2 = "blue": Timer1.Enabled = True ' It is downloading stuff
If InStr(Activity, "D-") Then A2 = "green"                       ' It has stopped download

' Only change the pictures if necessary...
' We don't want to constantly replace the same picture over and over... no need.
If A1 <> status.Panels("upload").Tag And A1 > "" Then
    status.Panels("upload").Picture = LoadResPicture(A1, vbResBitmap)
    status.Panels("upload").Tag = A1
    End If
If A2 <> status.Panels("download").Tag And A2 > "" Then
    status.Panels("download").Picture = LoadResPicture(A2, vbResBitmap)
    status.Panels("download").Tag = A2
    End If
End Sub
Private Sub RefreshUserCount()
Dim UsersOnlineString As String
UsersOnlineString = "Users online: " & users.ListItems.Count
If status.Panels("users_online") <> UsersOnlineString Then
    status.Panels("users_online") = UsersOnlineString
    End If
End Sub



' --------------------------------------------------------------------------------------------
' -[SEND FILE SUBROTINES]---------------------------------------------------------------------
' --------------------------------------------------------------------------------------------
Public Sub SendFilename(toip As String, localfile As String)
If CurrentBlock > -1 Then
    AddLine "Can't send anonther file yet... User from previous file hasn't agreed to accept or not", msgCritical, " "
    Exit Sub
    End If
AddLine "Asking permition to send file to user " & users.ListItems("ip" & toip).Text
SendingFile = localfile
SendPackage "SENDFILE" & vbTab & MyIP & FieldSeparator & Username & FieldSeparator & FileLen(localfile) & FieldSeparator & JustFile(localfile) & vbCrLf, toip
End Sub

Public Sub SendBlock(blocknumber As Long)
Dim a As Long
Dim dat As String
a = FreeFile
' Get the corresponding block
Open SendingFile For Binary As #a
Seek #a, blocknumber * BlockSize + 1
dat = Space(BlockSize)
Get #a, , dat
Close #a
If SendForm.Tag = "CANCEL" Then
    ' User canceled
    filesock2.Close
    Unload SendForm
    Exit Sub
    End If
SendForm.SetProgress blocknumber * BlockSize
If FileLen(SendingFile) - BlockSize < blocknumber * BlockSize + 1 Then
    dat = Left(dat, FileLen(SendingFile) - blocknumber * BlockSize)
    ' Last block
    CurrentBlock = -1
    Else
    CurrentBlock = blocknumber + 1
    End If
SetActivity "U+"
filesock2.SendData dat
End Sub
Private Sub filesock_Close()
If filesock.Tag > "" Then
    filesock.Close
    AddLine "File transfer completed", msgInformation, " "
    filesock.Tag = ""
    Unload ReceiveForm
    End If
End Sub

Private Sub filesock_Connect()
AddLine "Connected. Receiving file... ", msgInformation, " "
End Sub

Private Sub filesock_ConnectionRequest(ByVal requestID As Long)
filesock.Close
filesock.Accept requestID
End Sub

Private Sub filesock_DataArrival(ByVal bytesTotal As Long)
Dim a As String
Dim d() As String
Dim totalsize As Long
Dim filename As String

filesock.GetData a
d = Split(filesock.Tag, vbTab)
totalsize = Val(d(0))
filename = CurDir
If Right(filename, 1) <> "\" Then filename = filename & "\"
filename = filename & d(1)
Dim ff As Long
ff = FreeFile
Open filename For Append As #ff
Print #ff, a;
Close #ff
ReceiveForm.SetProgress FileLen(filename)
End Sub

Private Sub filesock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
AddLine "Error while sending file: " & Number & " - " & Description, msgCritical
End Sub

Private Sub filesock2_Connect()
AddLine "Connected. Sending file... " & SendingFile, msgInformation, " "
SendBlock 0
End Sub

Private Sub filesock2_SendComplete()
If CurrentBlock = -1 Then
    ' File finished
    filesock2.Close
    Unload SendForm
    Exit Sub
    End If
' Another block can be sent
SendBlock CurrentBlock
End Sub

Private Function GetSystemInfo() As String
Dim SInfo As SYSTEM_INFO
Dim lpbuffer As MEMORYSTATUS
modSysInfo.GetSystemInfo SInfo
modSysInfo.GlobalMemoryStatus lpbuffer

GetSystemInfo = "" & vbCrLf
GetSystemInfo = GetSystemInfo & "--- [Main system info] ---" & vbCrLf
GetSystemInfo = GetSystemInfo & "PC_Name=" & modChat.GetComputerName & vbCrLf
GetSystemInfo = GetSystemInfo & "OS=" & modSysInfo.ThisSys & vbCrLf
GetSystemInfo = GetSystemInfo & "Login=" & modChat.GetWindowsUser & vbCrLf
GetSystemInfo = GetSystemInfo & "ProcessorType=" & SInfo.dwProcessorType & vbCrLf
GetSystemInfo = GetSystemInfo & "Number_of_processors=" & SInfo.dwNumberOrfProcessors & vbCrLf
GetSystemInfo = GetSystemInfo & "OEMID=" & SInfo.dwOemID & vbCrLf
GetSystemInfo = GetSystemInfo & "Internet_Explorer=" & IEVersion & vbCrLf
GetSystemInfo = GetSystemInfo & "Screen=" & Screen.Width / Screen.TwipsPerPixelX & " x " & Screen.Height / Screen.TwipsPerPixelY & vbCrLf
If Printer Is Nothing Then
    GetSystemInfo = GetSystemInfo & "Printer=" & "Not installed" & vbCrLf
    Else
    GetSystemInfo = GetSystemInfo & "Printer=" & Printer.DeviceName & " on " & Printer.Port & " using " & Printer.DriverName & vbCrLf
    End If


GetSystemInfo = GetSystemInfo & vbCrLf

GetSystemInfo = GetSystemInfo & "--- [Physical hardware info] ---" & vbCrLf
GetSystemInfo = GetSystemInfo & "PhysicalMemory=" & Format(lpbuffer.dwTotalPhys / 1024, "#,##0.0") & " Kbytes" & vbCrLf
GetSystemInfo = GetSystemInfo & "VirtualMemory=" & Format(lpbuffer.dwTotalPageFile / 1024, "#,##0.0") & " Kbytes" & vbCrLf
GetSystemInfo = GetSystemInfo & "TotalMemory=" & Format(lpbuffer.dwTotalVirtual / 1024, "#,##0.0") & " Kbytes" & vbCrLf
GetSystemInfo = GetSystemInfo & "Memory_load=" & lpbuffer.dwMemoryLoad & "%" & vbCrLf
GetSystemInfo = GetSystemInfo & vbCrLf

GetSystemInfo = GetSystemInfo & "--- [Directories] ---" & vbCrLf
GetSystemInfo = GetSystemInfo & "WindowsDirectory=" & modSysInfo.GetWindowsDir & vbCrLf
GetSystemInfo = GetSystemInfo & "Desktop_Directory=" & GetSpecialfolder(sfidDESKTOP) & vbCrLf
GetSystemInfo = GetSystemInfo & "Startup_Directory=" & GetSpecialfolder(sfidSTARTUP) & vbCrLf
GetSystemInfo = GetSystemInfo & vbCrLf
End Function

Private Sub setSystrayActivity(Activity As SysTray_Activity)
Dim TIcn As Picture
Set TIcn = Nothing
If Activity = NoActivity Then
    Set TIcn = Me.Icon
    Timer2.Enabled = False
    End If
If Activity = ChatInProgress Then
    Timer2.Tag = "ACT0"
    Timer2.Enabled = True
    Set TIcn = LoadResPicture("ACT0", vbResIcon)
    End If
If Activity = Warning Then
    ' No icon for this one yet
    End If
If TIcn Is Nothing Then
    Else
    MySysTray.IconHandle = TIcn
    MySysTray.ShowIcon
    End If
End Sub

Private Sub Timer2_Timer()
Select Case Timer2.Tag
    Case "ACT0"
        Timer2.Tag = "ACT1"
    Case "ACT1"
        Timer2.Tag = "ACT0"
    End Select
MySysTray.IconHandle = LoadResPicture(CStr(Timer2.Tag), vbResIcon)
End Sub

Public Sub ShareDesktop(showwindow As Boolean)
If frmDesktop.Visible = False Then
    Set frmDesktop.ParentChat = Me
    SendPackage "ALERT" & vbTab & MyIP & FieldSeparator & Username & FieldSeparator & "is sharing his desktop" & vbCrLf
    frmDesktop.AllowedIps = "" ' Clears the allowed ip . You can add your ip as 'auto-authorized' here...
                            ' example: frmdesktop.allowedips = "10.0.0.1" & vbtab
                            ' Make sure the IP's are spaced with a vbtab (even if it is only one)
    If showwindow Then frmDesktop.Show
    End If
End Sub

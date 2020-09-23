VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   4005
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2764.322
   ScaleMode       =   0  'User
   ScaleWidth      =   6366.771
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   168.56
      ScaleMode       =   0  'User
      ScaleWidth      =   168.56
      TabIndex        =   1
      Top             =   240
      Width           =   240
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5460
      TabIndex        =   0
      Top             =   3600
      Width           =   1260
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2730
      TabIndex        =   5
      Top             =   420
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   7500.206
      Y1              =   2018.887
      Y2              =   2018.887
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2190
      Left            =   30
      TabIndex        =   2
      Top             =   690
      Width           =   6705
   End
   Begin VB.Label lblTitle 
      Caption         =   "FCL Network Chat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   810
      TabIndex        =   4
      Top             =   30
      Width           =   4515
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   7500.206
      Y1              =   2029.24
      Y2              =   2029.24
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":058A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   3060
      Width           =   6660
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    SetWindowTop Me.hwnd
    PutXPDropShadow Me
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Dim desc As String
    desc = "Hi everyone." & vbCrLf & vbCrLf
    desc = desc & "This is just a small, simple, and commented project for network chat, without using a server" & vbCrLf
    desc = desc & "" & vbCrLf
    desc = desc & "It's a small gift and token of my appreciation for www.planet-souce-code.com, since i have consulted" & vbCrLf
    desc = desc & "PSC many times for info and code examples, and never submited (or contributed) any of my code." & vbCrLf
    desc = desc & "" & vbCrLf
    desc = desc & "So, here is my contribution. Hope it's useful for someone." & vbCrLf
    desc = desc & "" & vbCrLf
    desc = desc & "Thank you PSC! :)" & vbCrLf
    desc = desc & "                                                       The author:" & vbCrLf
    desc = desc & "                                                   Filipe Camizão Lage" & vbCrLf
    desc = desc & "                                                   fclage@mail.net4b.pt" & vbCrLf
    lblDescription.Caption = desc
End Sub


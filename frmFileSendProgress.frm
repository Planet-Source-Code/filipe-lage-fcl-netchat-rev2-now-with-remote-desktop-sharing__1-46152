VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel transfer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2310
      TabIndex        =   2
      Top             =   1290
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   150
      TabIndex        =   1
      Top             =   840
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F5F5F5&
      BackStyle       =   0  'Transparent
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
      Index           =   2
      Left            =   1110
      TabIndex        =   7
      Top             =   570
      Width           =   5355
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F5F5F5&
      BackStyle       =   0  'Transparent
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
      Index           =   1
      Left            =   1110
      TabIndex        =   6
      Top             =   420
      Width           =   5355
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   1110
      TabIndex        =   5
      Top             =   0
      Width           =   5355
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total size:"
      Height          =   165
      Index           =   2
      Left            =   90
      TabIndex        =   4
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Current block:"
      Height          =   165
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "File:"
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   1005
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load()
SetWindowTop Me.hwnd
PutXPDropShadow Me
ProgressBar1.Value = 0
End Sub

Public Sub SetInfo(file As String, totalsize As String)
Label2(0).Caption = file
Label2(1).Caption = totalsize
DoEvents
End Sub
Public Sub SetProgress(donebytes As Long)
Dim p As Double
If Val(Label2(1).Caption) > 0 Then
    p = donebytes / Val(Label2(1).Caption)
    If p > 1 Then p = 1
    ProgressBar1.Value = p * ProgressBar1.Max
    DoEvents
    End If
End Sub


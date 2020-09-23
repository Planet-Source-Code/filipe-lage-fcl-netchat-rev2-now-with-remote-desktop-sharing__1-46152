VERSION 5.00
Begin VB.Form frmSysinfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System information"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Directory information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   0
      TabIndex        =   27
      Top             =   2310
      Width           =   9465
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   10
         Left            =   2130
         TabIndex        =   33
         Tag             =   "WindowsDirectory"
         Top             =   180
         Width           =   7275
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Windows directory:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   0
         TabIndex        =   32
         Top             =   210
         Width           =   2085
      End
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   11
         Left            =   2130
         TabIndex        =   31
         Tag             =   "Desktop_Directory"
         Top             =   450
         Width           =   7275
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Desktop directory:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   0
         TabIndex        =   30
         Top             =   480
         Width           =   2085
      End
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   12
         Left            =   2130
         TabIndex        =   29
         Tag             =   "Startup_Directory"
         Top             =   720
         Width           =   7275
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Startup directory:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   0
         TabIndex        =   28
         Top             =   750
         Width           =   2085
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Hardware information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   -30
      TabIndex        =   18
      Top             =   1560
      Width           =   9495
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   6
         Left            =   2190
         TabIndex        =   26
         Tag             =   "PhysicalMemory"
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "RAM installed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   60
         TabIndex        =   25
         Top             =   150
         Width           =   2085
      End
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   7
         Left            =   6930
         TabIndex        =   24
         Tag             =   "VirtualMemory"
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Page file size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   4800
         TabIndex        =   23
         Top             =   150
         Width           =   2085
      End
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   8
         Left            =   2190
         TabIndex        =   22
         Tag             =   "TotalMemory"
         Top             =   390
         Width           =   2535
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Total virtual memory:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   60
         TabIndex        =   21
         Top             =   420
         Width           =   2085
      End
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   9
         Left            =   6930
         TabIndex        =   20
         Tag             =   "Memory_load"
         Top             =   390
         Width           =   2535
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Usage (Memory load):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   4800
         TabIndex        =   19
         Top             =   420
         Width           =   2085
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Main computer information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9465
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Screen resolution:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   14
         Left            =   30
         TabIndex        =   17
         Top             =   1260
         Width           =   2085
      End
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   14
         Left            =   2160
         TabIndex        =   16
         Tag             =   "screen"
         Top             =   1230
         Width           =   7275
      End
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   15
         Tag             =   "PC_Name"
         Top             =   150
         Width           =   2535
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "PC Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   14
         Top             =   180
         Width           =   2085
      End
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   13
         Tag             =   "os"
         Top             =   420
         Width           =   2535
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Operating system:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   12
         Top             =   450
         Width           =   2085
      End
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   11
         Tag             =   "login"
         Top             =   690
         Width           =   2535
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Windows login:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   30
         TabIndex        =   10
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   3
         Left            =   6900
         TabIndex        =   9
         Tag             =   "processortype"
         Top             =   150
         Width           =   2535
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Processor type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   4770
         TabIndex        =   8
         Top             =   180
         Width           =   2085
      End
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   4
         Left            =   6900
         TabIndex        =   7
         Tag             =   "Number_of_processors"
         Top             =   420
         Width           =   2535
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of processors:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   4770
         TabIndex        =   6
         Top             =   450
         Width           =   2085
      End
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   5
         Left            =   6900
         TabIndex        =   5
         Tag             =   "Internet_Explorer"
         Top             =   690
         Width           =   2535
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Internet Explorer:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   4770
         TabIndex        =   4
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Active printer:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   30
         TabIndex        =   3
         Top             =   990
         Width           =   2085
      End
      Begin VB.Label field_info 
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
         Height          =   255
         Index           =   13
         Left            =   2160
         TabIndex        =   2
         Tag             =   "Printer"
         Top             =   960
         Width           =   7275
      End
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
      Height          =   345
      Left            =   8220
      TabIndex        =   0
      Top             =   3480
      Width           =   1245
   End
End
Attribute VB_Name = "frmSysinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Public Sub HandleInformation(d() As String)
Me.Caption = "Information on user " & d(1) & " (IP " & d(0) & ")"
Dim h() As String
Dim i() As String

h = Split(d(2), vbCrLf)
For o = LBound(h) To UBound(h)
ipos = 0
ipos = InStr(h(o), "=")
If ipos > 0 Then
    ' Got a field information
    fieldname = Left(h(o), ipos - 1)
    fieldvalue = Mid(h(o), ipos + 1)
    For l = Me.field_info.LBound To Me.field_info.UBound
    If LCase(field_info(l).Tag) = LCase(fieldname) Then
        field_info(l).Caption = fieldvalue
        Exit For
        End If
    Next
    End If
Next
SetWindowTop Me.hwnd
PutXPDropShadow Me


End Sub


Attribute VB_Name = "modChat"
Private Declare Function GetComputerNameEx Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Private Declare Function GetUserNameW Lib "advapi32.dll" (lpbuffer As Byte, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Declare Function GetClassLong Lib "USER32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "USER32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const CS_DROPSHADOW = &H20000
Private Const GCL_STYLE = (-26)
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_FRAMECHANGED = &H20
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_SHOWME = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TRANSPARENT = &H20
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Sub Main()
If InStr(LCase(Command$), "/?") Then ' help command
    MsgBox "/hide - start minimized in system tray" & vbCrLf & _
           "/autosharedesktop - automatically share desktop" & vbCrLf & _
           "/? - this help messsage box", vbInformation, App.Title
    End
    End If

If App.PrevInstance Then
    MsgBox "Can't run two " & App.ProductName & " at the same time...", vbOKOnly Or vbExclamation, "Previous instance found"
    ' We cant have 2 apps running...
    ' They will have the same IP and communication on the same port. Remember ?
    ' It would crash anyway!
    End
    End If
WriteManifest
WriteJPGLib
                   

DoEvents
InitCommonControls ' XP style controls if XP is installed...
                   ' No crashes in W95,98,ME,NT or 2000, so let's do it without checking the OS
                   ' (The manifest file was written in the 'WriteManifest' call)

Load frmChat
SetWindowTop frmChat.hWnd
PutXPDropShadow frmChat

If InStr(LCase(Command$), "/hide") Then ' If /hide command is specified, then
    frmChat.WindowState = 1
    End If

frmChat.Show

If InStr(LCase(Command$), "/autosharedesktop") Then ' If /hide command is specified, then
    frmChat.ShareDesktop False
    End If

End Sub

Public Sub PutXPDropShadow(theform As Form) ' Put's XP shadow in a form
SetClassLong theform.hWnd, GCL_STYLE, GetClassLong(theform.hWnd, GCL_STYLE) Or CS_DROPSHADOW
End Sub
Public Sub SetWindowNormal(lngHwnd As Long) ' Remove window from top
SetWindowPos lngHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Public Sub SetWindowTop(lngHwnd As Long)    ' Set window on top of others
SetWindowPos lngHwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Public Sub SetWindowFocus(lngHwnd As Long)  ' Set focus of the window
SetWindowPos lngHwnd, 0, 0, 0, 0, 0, SWP_SHOWME
End Sub
Public Sub SetWindowTransparent(lngHwnd As Long) ' Set form transparent
SetWindowPos lngHwnd, HWND_TOPMOST, 0, 0, 0, 0, WS_EX_TRANSPARENT
End Sub

Public Function GetComputerName() As String
    GetComputerName = String(255, Chr$(0))
    GetComputerNameEx GetComputerName, 255
    GetComputerName = Left$(GetComputerName, InStr(1, GetComputerName, Chr$(0)) - 1)
End Function
Public Sub StopWav()
Dim i As Long, rs As String, cb As Long
rs = Space$(128)
i = mciSendString("stop sound", rs, 128, cb)
i = mciSendString("close sound", rs, 128, cb)
End Sub
Public Sub StartWav(WaveFile As String)
Dim i As Long, rs As String, cb As Long, w$
StopWav
rs = Space$(128)
w$ = WaveFile
i = mciSendString("open waveaudio!" & w$ & " alias sound", rs, 128, cb)
If i = 0 Then i = mciSendString("play sound", rs, 128, cb)
End Sub
Public Function SoundFree() As Boolean
Dim i As Long, rs As String, cb As Long, w$
rs = Space$(128)
i = mciSendString("info sound", rs, 128, cb)
End Function
Public Function GetWindowsUser() As String
On Error Resume Next
GetWindowsUser = Space(128)
GetUserName GetWindowsUser, 128
If InStr(GetWindowsUser, Chr$(0)) > 0 Then
    GetWindowsUser = Left(GetWindowsUser, InStr(GetWindowsUser, Chr(0)) - 1)
    End If
End Function

Private Sub WriteManifest()
On Error Resume Next
Dim a As Long
targetmanifest = App.Path
If Right(targetmanifest, 1) <> "\" Then targetmanifest = targetmanifest & "\"

targetmanifest = targetmanifest & App.EXEName & ".exe.manifest"
dat = Replace(StrConv(LoadResData("XPMANIFEST", "CUSTOM"), vbUnicode), vbCrLf & vbCrLf, vbCrLf)
' dat = Replace(dat, " " & vbCrLf, vbCrLf)
a = FreeFile
Open targetmanifest For Output As #a
Print #a, dat;
Close #a
End Sub

Public Function JustFile(filename As String)
JustFile = filename
Dim o As Integer
For o = Len(filename) To 1 Step -1
Select Case Mid(filename, o, 1)
    Case ":", "\", "/"
        JustFile = Mid(filename, o + 1)
        Exit For
    End Select
Next
End Function

Public Function JustDirectory(filename As String)
On Error Resume Next
Dim jf As String
jf = JustFile(filename)
JustDirectory = Left(filename, Len(filename) - Len(jf))
End Function

Private Sub WriteJPGLib()
On Error Resume Next
If modJPG.InstalledOK = False Then
    ' IJL15.DLL isn't installed. Wait! I have it in my resource file!
    ' I'm going to write the Intel JPG handler in the Application's path
    ' There's no need to put it in the system or system32 directory... it'll work here
    b = -1
    TargetDLL = App.Path
    If Right(TargetDLL, 1) <> "\" Then TargetDLL = TargetDLL & "\"
    TargetDLL = TargetDLL & "ijl15.dll"
    b = FileLen(TargetDLL)
    If b <= 0 Then
        Dim a As Long
        a = FreeFile
        Open TargetDLL For Output As #a
        Print #a, StrConv(LoadResData("INTEL_JPG", "CUSTOM"), vbUnicode);
        Close #a
        End If
    End If
End Sub

Public Function JPGSeparator()
JPGSeparator = vbNullChar & "THIS_IS_THE_JPG_SEPARATOR" & vbNullChar
End Function

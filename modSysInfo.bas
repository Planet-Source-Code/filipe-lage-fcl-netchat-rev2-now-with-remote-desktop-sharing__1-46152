Attribute VB_Name = "modSysInfo"
' These are many api calls... Probably some aren't used in this, but it's a copy paste from other project I have
' and I don't want to retype code... So... Don't mind the mess in this module

Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpbuffer As MEMORYSTATUS)
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpbuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpbuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpbuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Private Declare Function SetWindowPos Lib "USER32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function OSRegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Private Declare Function OSRegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Private Declare Function OSRegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String) As Long
Private Declare Function OSRegEnumKey Lib "advapi32" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal iSubKey As Long, ByVal lpszName As String, ByVal cchName As Long) As Long
Private Declare Function OSRegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Private Declare Function OSRegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, lpdwType As Long, lpbData As Any, cbData As Long) As Long
Private Declare Function OSRegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function OSRegSetValueNumEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const gintMAX_SIZE = 256
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_NO_MORE_ITEMS = 259&
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type
Public Enum SpecialFolderIDs
    sfidDESKTOP = &H0
    sfidPROGRAMS = &H2
    sfidPERSONAL = &H5
    sfidFAVORITES = &H6
    sfidSTARTUP = &H7
    sfidRECENT = &H8
    sfidSENDTO = &H9
    sfidSTARTMENU = &HB
    sfidDESKTOPDIRECTORY = &H10
    sfidNETHOOD = &H13
    sfidFONTS = &H14
    sfidTEMPLATES = &H15
    sfidCOMMON_STARTMENU = &H16
    sfidCOMMON_PROGRAMS = &H17
    sfidCOMMON_STARTUP = &H18
    sfidCOMMON_DESKTOPDIRECTORY = &H19
    sfidAPPDATA = &H1A
    sfidPRINTHOOD = &H1B
    sfidProgramFiles = &H10000
    sfidCommonFiles = &H10001
End Enum
Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Type SHITEMID
    cb As Long
    abID As Byte
End Type
Public Type ITEMIDLIST
    mkid As SHITEMID
End Type
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Function ThisSys() As String
Dim osvi As OSVERSIONINFO
osvi.dwOSVersionInfoSize = Len(osvi)
If GetVersionEx(osvi) = 0 Then
    ThisSys = "Can´t retrieve system type"
    Exit Function
    End If
Select Case osvi.dwPlatformId
Case 1
    If osvi.dwMajorVersion = 4 Then
        If osvi.dwMinorVersion = 0 Then ThisSys = "Windows 95"
        If osvi.dwMinorVersion = 10 And osvi.dwBuildNumber < 67766446 Then ThisSys = "Windows 98"
        If osvi.dwMinorVersion = 10 And osvi.dwBuildNumber >= 67766446 Then ThisSys = "Windows 98 SE"
        If osvi.dwMinorVersion = 90 Then ThisSys = "Windows ME"
        Exit Function
        Else
        ThisSys = "Unknown system! Plat.Id=" & osvi.dwPlatformId & " V" & osvi.dwMajorVersion & "." & osvi.dwMinorVersion & " Built: " & osvi.dwBuildNumber & " CSD:" & osvi.szCSDVersion
        Exit Function
        End If
Case 2
    If osvi.dwMajorVersion = 5 Then
        If osvi.dwMinorVersion = 0 Then ThisSys = "Win 2000"
        If osvi.dwMinorVersion = 1 Then ThisSys = "Windows XP"
        Else
        ThisSys = "Unknown system! Plat.Id=" & osvi.dwPlatformId & " V" & osvi.dwMajorVersion & "." & osvi.dwMinorVersion & " Built: " & osvi.dwBuildNumber & " CSD:" & osvi.szCSDVersion
        End If
End Select
End Function

Public Function GetWindowsFontDir() As String
    GetWindowsFontDir = GetSpecialfolder(sfidFONTS)
End Function

Public Function GetWindowsDir() As String
    Dim strBuf As String
    strBuf = Space$(gintMAX_SIZE)
    If GetWindowsDirectory(strBuf, gintMAX_SIZE) Then
        GetWindowsDir = StringFromBuffer(strBuf)
    End If
End Function

Private Function StringFromBuffer(duh As String) As String
Dim a As Long
a = InStr(duh, Chr$(0))
If a = 0 Then a = Len(duh) + 1
StringFromBuffer = Left(duh, a - 1)
End Function

Public Function GetProgramFilesDir() As String
    Dim hKey As Long
    Dim strresolved As String
    Const strProgramFilesKey = "ProgramFilesDir"
    
    strPathsKey = RegPathWinCurrentVersion()
    strDestDir = Trim$(strDestDir)
    r = RegOpenKey(HKEY_LOCAL_MACHINE, strPathsKey, hKey)
    If hKey Then
        RegQueryStringValue hKey, strProgramFilesKey, strresolved
        RegCloseKey hKey
        GetProgramFilesDir = strresolved
        End If

End Function

Public Function GetSpecialfolder(CSIDL As SpecialFolderIDs) As String
    
    Dim r As Long
    Dim IDL As ITEMIDLIST
    
    'Get the special folder
    r = SHGetSpecialFolderLocation(App.hInstance, CSIDL, IDL)
    If r = NOERROR Then
        'Create a buffer
        GetSpecialfolder = Space$(512)
        'Get the path from the IDList
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal GetSpecialfolder)
        'Remove the unnecessary chr$(0)'s
        If InStr(GetSpecialfolder, Chr$(0)) > 0 Then
            GetSpecialfolder = Left$(GetSpecialfolder, InStr(GetSpecialfolder, Chr$(0)) - 1)
            End If
        If Right$(GetSpecialfolder, 1) <> "\" Then GetSpecialfolder = GetSpecialfolder & "\"
        Exit Function
    End If
    GetSpecialfolder = ""
End Function

Private Function RegPathWinCurrentVersion() As String
    RegPathWinCurrentVersion = "SOFTWARE\Microsoft\Windows\CurrentVersion"
End Function
Private Function RegPathWinPrograms() As String
    RegPathWinPrograms = RegPathWinCurrentVersion() & "\Explorer\Shell Folders"
End Function

Private Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String, strData As String) As Boolean
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    
    ' Get length/data type
    lResult = OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then
            strBuf = Space$(lDataBufSize)
            lResult = OSRegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then
                RegQueryStringValue = True
                strData = StringFromBuffer(strBuf)
            End If
        End If
    End If
End Function

Public Function IEVersion() As String

    Dim hKey As Long
    Dim strresolved As String
    Dim test_old As String
    Const strProgramFilesKey = "ProgramFilesDir"
    
    strPathsKey = "SOFTWARE\Microsoft\Internet Explorer"
    r = RegOpenKey(HKEY_LOCAL_MACHINE, CStr(strPathsKey), hKey)
    If hKey Then
        RegQueryStringValue hKey, "Version", strresolved
        RegQueryStringValue hKey, "IVer", test_old
        RegCloseKey hKey
        IEVersion = strresolved
        If IEVersion = "" Then
            If test_old <> "" Then
                IEVersion = test_old
                Else
                IEVersion = "Internet Explorer not detected"
                End If
            End If
        RegCloseKey hKey
        End If
End Function

Private Function strGetPredefinedHKEYString(ByVal hKey As Long) As String
    Select Case hKey
        Case HKEY_CLASSES_ROOT
            strGetPredefinedHKEYString = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_USER
            strGetPredefinedHKEYString = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE
            strGetPredefinedHKEYString = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS
            strGetPredefinedHKEYString = "HKEY_USERS"
    End Select
End Function



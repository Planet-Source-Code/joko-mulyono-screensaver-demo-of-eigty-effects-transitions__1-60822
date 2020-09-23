Attribute VB_Name = "Module1"
Option Explicit
Public Type SHELLEXECUTEINFO
    cbSize                                  As Long
    fMask                                   As Long
    hWnd                                    As Long
    lpVerb                                  As String
    lpFile                                  As String
    lpParameters                            As String
    lpdirectory                             As String
    nShow                                   As Long
    hInstApp                                As Long
    lpIDList                                As Long
    lpClass                                 As String
    hkeyClass                               As Long
    dwHotKey                                As Long
    hIcon                                   As Long
    hProcess                                As Long
End Type
Private Type SHFILEOPSTRUCT
    hWnd                                    As Long
    wFunc                                   As Long
    pFrom                                   As String
    pTo                                     As String
    fFlags                                  As Integer
    fAnyOperationsAborted                   As Long
    hNameMappings                           As Long
    lpszProgressTitle                       As String    '  only used if FOF_SIMPLEPROGRESS
End Type
Private Const BIF_STATUSTEXT            As Long = &H4
Private Const BIF_RETURNONLYFSDIRS      As Integer = 1
Private Const BIF_DONTGOBELOWDOMAIN     As Integer = 2
Private Const MAX_PATH                  As Integer = 260
Private Const WM_USER                   As Long = &H400
Private Const BFFM_INITIALIZED          As Integer = 1
Private Const BFFM_SELCHANGED           As Integer = 2
Private Const BFFM_SETSTATUSTEXT        As Double = (WM_USER + 100) 'Make new folder
Private Const BFFM_SETSELECTION         As Double = (WM_USER + 102)
Private Type BrowseInfo
    hWndOwner                               As Long
    pIDLRoot                                As Long
    pszDisplayName                          As Long
    lpszTitle                               As Long
    ulFlags                                 As Long
    lpfnCallback                            As Long
    lParam                                  As Long
    iImage                                  As Long
End Type
Private Type OPENFILENAME
    lStructSize                             As Long
    hWndOwner                               As Long
    hInstance                               As Long
    lpstrFilter                             As String
    lpstrCustomFilter                       As String
    nMaxCustFilter                          As Long
    nFilterIndex                            As Long
    lpstrFile                               As String
    nMaxFile                                As Long
    lpstrFileTitle                          As String
    nMaxFileTitle                           As Long
    lpstrInitialDir                         As String
    lpstrTitle                              As String
    Flags                                   As Long
    nFileOffset                             As Integer
    nFileExtension                          As Integer
    lpstrDefExt                             As String
    lCustData                               As Long
    lpfnHook                                As Long
    lpTemplateName                          As String
End Type
'Public DIRx                             As String
Private m_CurrentDirectory              As String    'The current directory
Private Start                           As Integer
Private pos                             As Integer
Private directory                       As String
Private result                          As Boolean
Private Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, _
                                                                                ByVal lpszTitle As String, _
                                                                                ByVal cbBuf As Integer) As Integer
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        ByVal lParam As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, _
                                                            ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
                                                                  ByVal lpString2 As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
                                                                                    ByVal lpszShortPath As String, _
                                                                                    ByVal cchBuffer As Long) As Long
Public FolderPath   As String
Public Function AppPath() As String
Dim strPath As String
    strPath = App.Path
    If Right$(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    AppPath = strPath
End Function

Private Function BrowseCallbackProc(ByVal lnghwnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal lp As Long, _
                                    ByVal pData As Long) As Long

Dim ret     As Long
Dim sBuffer As String
    On Error Resume Next
    Select Case uMsg
    Case BFFM_INITIALIZED
        SendMessage lnghwnd, BFFM_SETSELECTION, 1, m_CurrentDirectory
    Case BFFM_SELCHANGED
        sBuffer = Space$(MAX_PATH)
        ret = SHGetPathFromIDList(lp, sBuffer)
        If ret = 1 Then
            SendMessage lnghwnd, BFFM_SETSTATUSTEXT, 0, sBuffer
        End If
    End Select
    BrowseCallbackProc = 0
    On Error GoTo 0
End Function
Public Function BrowseForFolder(Owner As Form, _
                                ByVal Title As String, _
                                ByVal StartDir As String) As String

Dim lpIDList    As Long
Dim sBuffer     As String
Dim tBrowseInfo As BrowseInfo
    m_CurrentDirectory = StartDir & vbNullChar
    With tBrowseInfo
        .hWndOwner = Owner.hWnd
        .lpszTitle = lstrcat("Select A Directory", vbNullString)
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BFFM_SETSTATUSTEXT + BFFM_SETSELECTION + BIF_STATUSTEXT
        .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If lpIDList Then
        sBuffer = Space$(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseForFolder = sBuffer
    Else
        BrowseForFolder = vbNullString
    End If
End Function

Public Function GET_DIRECTORY(frmForm As Form) As String
Dim getdir As String
Dim M      As String
Dim N      As String

    M = GetSetting("setting", "directory", "mDir", getdir)
    N = App.Path
    If M <> vbNullChar Then
        getdir = BrowseForFolder(frmForm, "Select A Directory", M) ', MediaFilename)
    ElseIf LenB(M) = 0 Then
        getdir = BrowseForFolder(frmForm, "Select A Directory", N) ', MediaFilename)
    End If
   ' DIRx = getdir
    SaveSetting "setting", "directory", "mDir", getdir
    GET_DIRECTORY = getdir
End Function
Private Function GetAddressofFunction(add As Long) As Long
    GetAddressofFunction = add
End Function
Public Function GetExtension(ByVal FPath As String) As String
Dim p As Long
    If Len(FPath) > 0 Then
        p = InStrRev(FPath, ".")
        If p > 0 Then
            If p < Len(FPath) Then
                GetExtension = Mid$(FPath, p + 1)
            End If
        End If
    End If
End Function
Public Function GetFTitle(strFileName As String) As String
Dim cbBuf As String
    On Error GoTo GFTError
    cbBuf = String$(250, vbNullChar) 'Fill buffer with null chars
    GetFileTitle strFileName, cbBuf, Len(cbBuf) 'Get file title
    GetFTitle = Left$(cbBuf, InStr(1, cbBuf, vbNullChar) - 1) 'Extract file title from buffer
GFTError:
End Function
Public Function GetParent(Key As String) As String
    GetParent = Left$(Key, InStrRev(Key, "\") - 1)
End Function
Public Function getShortPath(fullpath As String) As String
'TESTED:OK
'purpose: to shorter long path
'If you have a file on ex:"C:\WINDOWS\Folder1\Folder2\Folder3\Folder4\TheMATRIX.dat" then
'this player can not open that file because the path may be more then max_path=260, so this is the solusion.
Dim lenPath          As String * 255
Dim Tmp              As String * 255
Dim ShortPathAndFile As String
    If LenB(fullpath) Then
        lenPath = GetShortPathName(fullpath, Tmp, 255)
        ShortPathAndFile = Left$(Tmp, lenPath)
        getShortPath = ShortPathAndFile
    End If
End Function


Public Function GetWindowsDir() As String
   Dim s As String
   Dim i As Integer
   i = GetWindowsDirectoryA("", 0)
   s = Space(i)
   Call GetWindowsDirectoryA(s, i)
   GetWindowsDir = AddBackslash(Left$(s, i - 1))
End Function
Public Function AddBackslash(s As String) As String
   If Len(s) > 0 Then
      If Right$(s, 1) <> "\" Then
         AddBackslash = s + "\"
      Else
         AddBackslash = s
      End If
   Else
      AddBackslash = "\"
   End If
End Function


Attribute VB_Name = "mdlBrowseFolder"
Option Explicit

Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long


Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, _
    ByVal nFolder As Long, ppidl As Long) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
ByVal pszPath As String) As Long

Private Declare Function GetForegroundWindow Lib "USER32" () As Long

Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
    
Private Const MAX_PATH = 260

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_NEWDIALOGSTYLE = &H40

Private Const BFFM_INITIALIZED As Long = 1
Private Const BFFM_SELCHANGED As Long = 2
Private Const BFFM_VALIDATEFAILED As Long = 3

Private Const WM_USER = &H400

Private Const BFFM_SETSTATUSTEXT As Long = (WM_USER + 100)
Private Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTION As Long = (WM_USER + 102)
   
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40

Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

Private Const PAGE_READWRITE         As Long = &H4

Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Public Function GetFolder(ByVal title As String, ByVal start As String, ByVal newfolder As Boolean) As String

Dim BI          As BROWSEINFO
Dim pidl        As Long
Dim lpSelPath   As Long
Dim sPath       As String * MAX_PATH
Dim a           As Long
    'fill in the info it needs
    
    
    
    With BI
        .hOwner = GetForegroundWindow
        
        .pidlRoot = SHGetSpecialFolderLocation(.hOwner, &H11, .pidlRoot)
        If Len(start) = 0 Then start = vbNullChar
        '.pidlRoot = 0
        .lpszTitle = title
        .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
        .ulFlags = BIF_RETURNONLYFSDIRS
        If newfolder = True Then .ulFlags = BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE
        'lpSelPath = LocalAlloc(LPTR, Len(start) + 1)
        lpSelPath = CoTaskMemAlloc(Len(start) + 1)
        
        If lpSelPath Then
            VirtualProtect lpSelPath, Len(start) + 1, PAGE_READWRITE, pidl
            CopyMemory ByVal lpSelPath, ByVal start, Len(start) + 1
            .lParam = lpSelPath
        End If
    End With
    
    'get the idlist long from the returned folder
    pidl = SHBrowseForFolder(BI)
    
    'do then if they clicked ok
    If pidl Then
        sPath = String$(MAX_PATH, 0)
        If SHGetPathFromIDList(pidl, sPath) Then
            'next line is the returned folder
            a = InStr(1, sPath, vbNullChar) - 1
            If a < 0 Then a = 0
            GetFolder = Left$(sPath, a)
        End If
        Call CoTaskMemFree(pidl)
    Else
        'user clicked cancel
    End If
    
    'Call LocalFree(lpSelPath)
    Call CoTaskMemFree(lpSelPath)
End Function

'this seems to happen before the box comes up and when a folder is clicked on within it
Public Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    Dim sPath As String, bFlag As Long
                                       
    sPath = Space$(MAX_PATH)
        
    Select Case uMsg
        Case BFFM_INITIALIZED
            'browse has been initialized, set the start folder
            Call SendMessage(hWnd, BFFM_SETSELECTION, 1, ByVal lpData)
        Case BFFM_SELCHANGED
            If SHGetPathFromIDList(lParam, sPath) Then
                sPath = Left(sPath, InStr(1, sPath, Chr(0)) - 1)
            End If
    End Select
          
End Function
          
Public Function FARPROC(pfn As Long) As Long
    FARPROC = pfn
End Function



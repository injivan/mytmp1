Attribute VB_Name = "mdlMakeDirList"
Option Explicit


'nFunction = 3000 -> Main
'nFunction = 3001 -> GetNames
'nFunction = 3002 -> dir_Folder
'nFunction = 3003 -> OpenDocument
'nFunction = 3004 -> UpdateProgressBar_CopyFile
'nFunction = 3005 -> ShowProgressBar_CopyFile
'nFunction = 3006 -> ShowProgressBar_CountFile
'nFunction = 3007 -> HideProgressBar
'nFunction = 3008 -> UpdateProgressBar_CountFile
'nFunction = 3009 -> putFolderAndFile
'nFunction = 3010 -> MakeNewFiles

Private Const ModulIdString     As String = "mdlMakeDirList - "
 
 
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
 

Public Const sINI_Name          As String = "mp3_List.ini"
Public DoCancel                 As Boolean
Public DirAppl                  As String

Public cSets                    As clsSetings
Public clsForm                  As frmShowList

'!!!!!!
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5            '  access denied
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2                     '  file not found
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_PNF = 3                     '  path not found
Private Const SE_ERR_OOM = 8                     '  out of memory
Private Const SE_ERR_SHARE = 26

Private Const MAX_PATH = 260
Private Const BFFM_INITIALIZED As Long = 1
Private Const BFFM_SELCHANGED As Long = 2

Private Const WM_USER = &H400
Private Const BFFM_SETSELECTION As Long = (WM_USER + 102)
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long

Private mTime  As Double
Private cf     As cFile
Private cF_Poz As Long

Private sFldr()     As String  '- ����� �� �������� ���
Private sFiles()    As String  '- ����� �� �������

Private sFldr_Co    As Long
Private sFiles_Co   As Long


'Purpose     :  Allows the user to select a folder
'Inputs      :  sCaption                The caption text on the dialog
'               [sDefault]                The default path to return if the user presses cancel
'Outputs     :  Returns the select path


 
Private Sub Main()
Const nFunction = 3000
On Error GoTo ErH
10
    DirAppl = App.Path: DirAppl = AddDirSep(DirAppl)
100
    If Not (cSets Is Nothing) Then Set cSets = Nothing
    Set cSets = New clsSetings
    If Not cSets.GetInitData Then GoTo ErH
200
    If Not (clsForm Is Nothing) Then Set clsForm = Nothing
    Set clsForm = New frmShowList
    clsForm.Show
300
    
400
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub
Private Function dir_Folder(ByVal sFolder As String, _
                            ByRef sFiles() As String, _
                            ByRef UBound_Arr As Long, _
                            ByVal flDir As Boolean, _
                            Optional sFindDir As String = vbNullString) As Boolean

Const nFunction = 3002
On Error GoTo ErH

Dim sBuf    As String
Dim i       As Long
Dim sDir    As String

10
    '������� �� ��������� � ����� �� ��������� � ������������ � �������� ���
    sBuf = sFolder
    'FS.AddDirSep sBuf
    sBuf = sBuf & sFindDir
    If Right$(sBuf, 1) <> "\" Then sBuf = sBuf & "\"
100
    On Error Resume Next
    If flDir Then
        sDir = Dir(sBuf, vbDirectory)
    Else
        sDir = Dir(sBuf)
    End If
    On Error GoTo ErH
    
150 ReDim sFiles(10) As String
    i = -1
    Do While Len(sDir)
        If Not (sDir = "." Or sDir = "..") Then
            i = i + 1
            If i > UBound(sFiles) Then ReDim Preserve sFiles(i + 10)
            sFiles(i) = sBuf & sDir
        End If
        sDir = Dir
    Loop
200
    UBound_Arr = i
    dir_Folder = True
Exit Function
ErH:
If Err.Number Then
    ShowErrMesage Err, ModulIdString, nFunction, Erl
End If
Err.Clear
End Function
Public Function DoObr(ByVal sFold As String, ByRef z As Long) As Boolean
On Error Resume Next
 

    If Not (cf Is Nothing) Then Set cf = Nothing
    Set cf = New cFile
    
    ReDim sFldr(10) As String
    ReDim sFiles(50) As String
    sFldr_Co = -1
    sFiles_Co = -1
    
    
    
    ShowProgressBar_CountFile "Folders"
    ShowProgressBar_CopyFile "Files"
    DoCancel = False
    'DoEvents
    
    cf.OpenFile cSets.sOutDir & cSets.sOutFile & ".txt", 1
    cf.SetEOF 0
    cF_Poz = 1
    '�������� ������� ��� ����
    GetNames sFold, z
    '���� �� ����� �� �������� � ���� �����
    If DoCancel Then GoTo ErH
    
    HideProgressBar
    
    Set cf = Nothing
    '����� �� ���� ���� � ����������
    '����� �� ���� ���� � �����
    
    
    Set cf = New cFile
    cf.OpenFile cSets.sOutDir & cSets.sOutFile & "_1.txt", 1
    cf.SetEOF 0
    Call MakeNewFiles(sFldr, sFldr_Co, cf)
    cf.CloseFile
    Set cf = Nothing
    
    
    Set cf = New cFile
    cf.OpenFile cSets.sOutDir & cSets.sOutFile & "_2.txt", 1
    cf.SetEOF 0
    Call MakeNewFiles(sFiles, sFiles_Co, cf)
    cf.CloseFile
    Set cf = Nothing
    
    HideProgressBar
    DoCancel = False
    DoObr = True
Exit Function
ErH:
 
Err.Clear
End Function
Private Function GetNames(ByVal sFold As String, ByRef z As Long) As Boolean
Const nFunction = 3001
On Error GoTo ErH
Dim sFolders() As String
Dim i As Long
Dim j As Long
    
10
    If Not dir_Folder(sFold, sFolders, j, True) Then GoTo ErH
20
    For i = 0 To j
30
        UpdateProgressBar_CopyFile "Files  " & i & " From " & j, i, j
        If DoCancel Then GoTo ErH
40
        If Len(sFolders(i)) Then
            If LCase$(Right$(sFolders(i), 3)) = "mp3" Then
                z = z + 1
50              cf.PutData cF_Poz, 0, Right$(Space$(4) & z, 4) & ". " & sFolders(i) & vbCrLf
                cF_Poz = cF_Poz + Len(Right$(Space$(4) & z, 4) & ". " & sFolders(i) & vbCrLf)
                
                If Not putFolderAndFile(z, sFolders(i)) Then GoTo ErH
                
60              sFolders(i) = vbNullString
            End If
        End If
    Next i
    For i = 0 To j
        If DoCancel Then GoTo ErH
70      UpdateProgressBar_CountFile "Folder " & i & " From " & j, i, j
        If Len(sFolders(i)) Then
80          If Not GetNames(sFolders(i), z) Then GoTo ErH
        End If
    Next i
    GetNames = True
Exit Function
ErH:
If Err.Number Then
    ShowErrMesage Err, ModulIdString, nFunction, Erl
End If
Err.Clear
End Function


Private Function UpdateProgressBar_CopyFile(ByVal msg As String, _
                                            ByVal Value As Long, _
                                            ByVal max As Long) As Boolean
Const nFunction = 3004
On Error GoTo ErH
Dim p   As Long
Dim h   As Long
Dim FK  As Long
10
   If max = 0 Then max = 1
   p = Value / max * 100
   If Trim$(msg) = vbNullString Then msg = clsForm.scrCopyFile.DataMember
    With clsForm.scrCopyFile
        .Cls
        .CurrentX = (.ScaleWidth - .TextWidth(msg)) \ 2
        .CurrentY = (.ScaleHeight - .TextHeight(msg)) \ 2
        h = .ScaleHeight
        FK = .ForeColor
    End With
    clsForm.scrCopyFile.Print msg
    clsForm.scrCopyFile.Line (0, 0)-(p, h), FK, BF
    
    If mTime < Timer Then
        DoEvents
        mTime = Timer + 0.15
    End If
    
    UpdateProgressBar_CopyFile = True
Exit Function
ErH:

If Err.Number Then
    ShowErrMesage Err, ModulIdString, nFunction, Erl
End If
Err.Clear
End Function

Private Function ShowProgressBar_CopyFile(ByVal sMsg As String) As Boolean

Const nFunction = 3005
On Error GoTo ErH

10
    clsForm.scrCopyFile.Visible = True
    clsForm.cmbCancel.Visible = True
     
100
    With clsForm.scrCopyFile
        .DataMember = sMsg
        .AutoRedraw = True
        .BackColor = vbWhite
        .ForeColor = vbBlue
        '.Height = 400
        .ScaleWidth = 100
        .ScaleHeight = 20
200     .DrawMode = vbNotXorPen
        .Visible = True
        .Cls
        .CurrentX = (.ScaleWidth - .TextWidth(sMsg)) \ 2
        .CurrentY = (.ScaleHeight - .TextHeight(sMsg)) \ 2
        clsForm.scrCopyFile.Print sMsg
    End With
    
    ShowProgressBar_CopyFile = True
    
Exit Function
ErH:

If Err.Number Then
    ShowErrMesage Err, ModulIdString, nFunction, Erl
End If
Err.Clear
End Function




Private Function ShowProgressBar_CountFile(ByVal sMsg As String) As Boolean

Const nFunction = 3006
On Error GoTo ErH

10
    clsForm.scrCountFile.Visible = True
    clsForm.cmbCancel.Visible = True
     
100
    With clsForm.scrCountFile
        .DataMember = sMsg
        .AutoRedraw = True
        .BackColor = vbWhite
        .ForeColor = vbBlue
        '.Height = 400
        .ScaleWidth = 100
        .ScaleHeight = 20
200     .DrawMode = vbNotXorPen
        .Visible = True
        .Cls
        .CurrentX = (.ScaleWidth - .TextWidth(sMsg)) \ 2
        .CurrentY = (.ScaleHeight - .TextHeight(sMsg)) \ 2
        clsForm.scrCountFile.Print sMsg
    End With
    
    ShowProgressBar_CountFile = True
    
Exit Function
ErH:

If Err.Number Then
    ShowErrMesage Err, ModulIdString, nFunction, Erl
End If
Err.Clear
End Function
Private Function HideProgressBar() As Boolean
Const nFunction = 3007
On Error GoTo ErH
10
    clsForm.scrCopyFile.Visible = False
    clsForm.scrCountFile.Visible = False
    clsForm.cmbCancel.Visible = False
    
Exit Function
ErH:
If Err.Number Then
    ShowErrMesage Err, ModulIdString, nFunction, Erl
End If
Err.Clear
End Function
Private Function UpdateProgressBar_CountFile(ByVal msg As String, _
                                            ByVal Value As Long, _
                                            ByVal max As Long) As Boolean
Const nFunction = 3008
On Error GoTo ErH
Dim p As Long
Dim h As Long
Dim FK As Long
10
   If max = 0 Then max = 1
   p = Value / max * 100
   If Trim$(msg) = vbNullString Then msg = clsForm.scrCountFile.DataMember
    With clsForm.scrCountFile
        .Cls
        .CurrentX = (.ScaleWidth - .TextWidth(msg)) \ 2
        .CurrentY = (.ScaleHeight - .TextHeight(msg)) \ 2
        h = .ScaleHeight
        FK = .ForeColor
    End With
    clsForm.scrCountFile.Print msg
    clsForm.scrCountFile.Line (0, 0)-(p, h), FK, BF
    
    If mTime < Timer Then
        DoEvents
        mTime = Timer + 0.15
    End If
    UpdateProgressBar_CountFile = True
Exit Function
ErH:

If Err.Number Then
    ShowErrMesage Err, ModulIdString, nFunction, Erl
End If
Err.Clear
End Function

'''Public Sub ShowErrMesage(ByVal MyError As Object, _
'''                         ByVal ModIDstr As String, _
'''                         ByVal nFunction As Long, _
'''                         ByVal erLine As Long)
'''
'''Dim ERN As String
'''Dim ERD As String
'''
'''    ERN = Err.Number
'''    ERD = Err.Description
'''
'''    MsgBox "Error" & " " & ModIDstr & " " & nFunction & vbCrLf & _
'''           "Row :" & " " & erLine & vbCrLf & _
'''           "Number:" & " " & ERN & vbCrLf & _
'''           "Description:" & " " & ERD, vbCritical, "Error"
'''
'''
'''End Sub


Public Function OpenDocument(ByVal FileName As String) As Boolean
Const nFunction = 3003
Dim a As Long
Dim lErr As Long, sErr As String
On Error GoTo ErH
10
    a = ShellExecute(0, vbNullString, ToGetShortName(FileName), vbNullString, vbNullString, vbNormalFocus)
       
    If (a < 0) Or (a > 32) Then
        OpenDocument = True
    Else
        Select Case a
        Case 0
            lErr = 7: sErr = "Out of memory"
        Case ERROR_FILE_NOT_FOUND
            lErr = 53: sErr = "File not found"
        Case ERROR_PATH_NOT_FOUND
80          lErr = 76: sErr = "Path not found"
        Case ERROR_BAD_FORMAT
100         sErr = "The executable file is invalid or corrupt"
        Case SE_ERR_ACCESSDENIED
            lErr = 75: sErr = "Path/file access error"
        Case SE_ERR_ASSOCINCOMPLETE
200         sErr = "This file type does not have a valid file association."
        Case SE_ERR_DDEBUSY
            lErr = 285: sErr = "The file could not be opened because the target application is busy.  Please try again in a moment."
        Case SE_ERR_DDEFAIL
300         lErr = 285: sErr = "The file could not be opened because the DDE transaction failed.  Please try again in a moment."
        Case SE_ERR_DDETIMEOUT
400         lErr = 286: sErr = "The file could not be opened due to time out.  Please try again in a moment."
        Case SE_ERR_DLLNOTFOUND
            lErr = 48: sErr = "The specified dynamic-link library was not found."
        Case SE_ERR_FNF
500         lErr = 53: sErr = "File not found"
        Case SE_ERR_NOASSOC
            sErr = "No application is associated with this file type."
        Case SE_ERR_OOM
            lErr = 7: sErr = "Out of memory"
        Case SE_ERR_PNF
600         lErr = 76: sErr = "Path not found"
        Case SE_ERR_SHARE
            lErr = 75: sErr = "A sharing violation occurred."
        Case Else
800         sErr = "An error occurred occurred whilst trying to open or print the selected file."
        End Select
        '��������� � ������ ��� ���� �� �������� �� ����� ������ ����
        '����� ���� �� ���� �� ������ � �������
        Err.Raise lErr, , sErr
        OpenDocument = False
    End If
Exit Function
ErH:
If Err.Number Then
    ShowErrMesage Err, ModulIdString, nFunction, Erl
End If

Err.Clear

End Function

Public Function ToGetShortName(ByVal sLongFileName As String) As String

Dim lRetVal As Long
Dim sShortPathName As String
Dim iLen As Integer
On Error Resume Next

    'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)
    
    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    'Strip away unwanted characters.
    ToGetShortName = Left(sShortPathName, lRetVal)

End Function

 
Private Function putFolderAndFile(ByVal lPoz As Long, ByVal sFileName As String) As Boolean
Const nFunction = 3009
On Error GoTo ErH
Dim i As Long
Dim a() As String
10
   
    a = Split(Trim$(sFileName), "\")
    i = UBound(a)
    
    
    sFiles_Co = sFiles_Co + 1
    If sFiles_Co > UBound(sFiles) Then ReDim Preserve sFiles(sFiles_Co * 2) As String
    sFiles(sFiles_Co) = Right$(Space$(4) & lPoz, 4) & ". " & a(i)            '��� �� �����
    
    
    i = i - 1
    If i >= 0 Then
        a(i + 1) = vbNullString
        If sFldr_Co >= 0 Then a(i + 1) = Mid$(sFldr(sFldr_Co), 7)
        
        If Trim$(a(i)) <> a(i + 1) Then
            sFldr_Co = sFldr_Co + 1
            If sFldr_Co > UBound(sFldr) Then ReDim Preserve sFldr(sFldr_Co * 2) As String
            sFldr(sFldr_Co) = Right$(Space$(4) & lPoz, 4) & ". " & a(i)            '��� �� dir
        End If
    End If
    
    
    putFolderAndFile = True
   
   
Exit Function
ErH:

If Err.Number Then
    ShowErrMesage Err, ModulIdString, nFunction, Erl
End If
Err.Clear
End Function
 


Private Function MakeNewFiles(ByRef arrStr() As String, _
                              ByRef arrStr_Co As Long, _
                              ByRef cFile As cFile) As Boolean
Const nFunction = 3010
On Error GoTo ErH
Dim cF_Poz  As Long
Dim i       As Long
10
    'Function for Make a file form arraj
    'arrays sFldr()     As String  '- ����� �� �������� ���
    '       sFiles()    as String  '- ����� �� �������
    
    
    cF_Poz = 1
   
    For i = 0 To arrStr_Co
100     cFile.PutData cF_Poz, 0, arrStr(i) & vbCrLf
200     cF_Poz = cF_Poz + Len(arrStr(i) & vbCrLf)
    Next i
          
    MakeNewFiles = True
   
   
Exit Function
ErH:

If Err.Number Then
    ShowErrMesage Err, ModulIdString, nFunction, Erl
End If
Err.Clear
End Function
 












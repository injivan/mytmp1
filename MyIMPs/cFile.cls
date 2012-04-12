VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "File Work Clas"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Const ModulIdString   As String = " cFile - "


Option Explicit
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long) As Long


Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_ALWAYS = 4
Private Const FILE_BEGIN = 0




'nFunction = 3001 -> GetData
'nFunction = 3002 -> OpenFile

'nFunction = 3004 -> PutData

Private Const mdlName = "cFile - "

Private hFile_1         As Long 'File Number
Private mvarRecCount    As Long 'local copy
Private myRecLen        As Long 'Len of the Record - local copy


 
Public Property Get RecCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RecCount
    RecCount = mvarRecCount
End Property


Public Function GetData(ByVal Record As Long, _
                        ByVal Offset As Long, _
                        ByRef Bufer As Variant) As Boolean
Const nFunction = 3001
On Error GoTo ErH
Dim startByte As Long
Dim sB        As String
Dim lSuccess As Long
Dim bytArr() As Byte
Dim lBytesRead As Long
Dim lBytesToRead As Long

10
    'Функция за Четена на данни от файл
    
    'Пресмятам от каде да чета
    startByte = Record * myRecLen + (Offset - 1)
    If startByte < 0 Then startByte = 0
    If Record > mvarRecCount Then GoTo ErH
20
     
    
    If (Len(Bufer) + startByte) > mvarRecCount * myRecLen + 1 Then GoTo ErH
30
    
    lBytesToRead = Len(Bufer)
    ReDim bytArr(lBytesToRead) As Byte
40
    lSuccess = SetFilePointer(hFile_1, startByte, 0, FILE_BEGIN)
    lSuccess = ReadFile(hFile_1, bytArr(LBound(bytArr)), _
                       lBytesToRead, lBytesRead, ByVal 0&)
     
    
     
    
    If VarType(Bufer) And &H2000 Then
        'vbArray
        Bufer = bytArr
    Else
        Bufer = ByteArrayToString(bytArr)
    End If
50
     
    
    GetData = True
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function
Public Sub SetEOF(ByVal nRec As Long)
Dim Ret As Long

    Ret = SetFilePointer(hFile_1, nRec, 0, FILE_BEGIN)
    Ret = SetEndOfFile(hFile_1)
    
End Sub

Public Function OpenFile(ByVal FilePath As String, _
                         ByVal RecLen As Long) As Boolean
Const nFunction = 3002
On Error GoTo ErH
10
    'Функция за отваряне на файл
    
    myRecLen = RecLen 'Дължина на 1 запис
    If myRecLen <= 0 Then myRecLen = 1
    
    hFile_1 = CreateFile(FilePath, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_ALWAYS, 0, 0)
    If hFile_1 = -1 Then hFile_1 = 0: Exit Function
    
    mvarRecCount = (GetFileSize(hFile_1, 0) / myRecLen) + 1 'Брой записи
    
    OpenFile = True

Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function



Public Function PutData(ByVal Record As Long, _
                        ByVal Offset As Long, _
                        ByVal Bufer As Variant) As Boolean

Const nFunction = 3004
On Error GoTo ErH
Dim startByte As Long
Dim sB        As String
    
Dim lSuccess        As Long
Dim lBytesWritten   As Long
Dim lBytesToWrite   As Long
       
10
    'Пресмятам от каде да пиша
    startByte = Record * myRecLen + (Offset - 1)
    If startByte < 0 Then startByte = 0
    If startByte > mvarRecCount * myRecLen Then
        startByte = mvarRecCount * myRecLen
    End If
20
    sB = Bufer
    lBytesToWrite = Len(sB)
     
    lSuccess = SetFilePointer(hFile_1, startByte, 0, FILE_BEGIN)
    lSuccess = WriteFile(hFile_1, ByVal sB, _
                        lBytesToWrite, lBytesWritten, 0) <> 0
30
    If lSuccess <> 0 Then
      ' Flush the file buffers (not sure if this is necessary)
        lSuccess = FlushFileBuffers(hFile_1)
        mvarRecCount = (GetFileSize(hFile_1, 0) / myRecLen) + 1 'Брой записи
    End If
    PutData = True
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function





Public Sub CloseFile()
    If hFile_1 Then CloseHandle hFile_1: hFile_1 = 0
End Sub

 
Private Sub Class_Terminate()
    If hFile_1 Then CloseHandle hFile_1: hFile_1 = 0
End Sub
Private Function ByteArrayToString(bytArray() As Byte) As String
    Dim sAns As String
    Dim iPos As String

    sAns = StrConv(bytArray, vbUnicode)
    iPos = InStrRev(sAns, Chr(0))
    If iPos > 0 Then sAns = Left(sAns, iPos - 1)

    ByteArrayToString = sAns

 End Function
''''
''''Private Function ByteArrayToString(Bytes() As Byte) As String
''''Dim iUnicode As Long
''''Dim i        As Long
''''Dim j        As Long
''''
''''On Error Resume Next
''''    i = UBound(Bytes)
''''
''''    If (i < 1) Then
''''        'ANSI, just convert to unicode and return
''''        ByteArrayToString = StrConv(Bytes, vbUnicode)
''''        Exit Function
''''    End If
''''    i = i + 1
''''
''''    'Examine the first two bytes
''''    CopyMemory iUnicode, Bytes(0), 2
''''
''''    If iUnicode = Bytes(0) Then 'Unicode
''''        'Account for terminating null
''''        If (i Mod 2) Then i = i - 1
''''        'Set up a buffer to recieve the string
''''        ByteArrayToString = String$(i / 2, 0)
''''        'Copy to string
''''        CopyMemory ByVal StrPtr(ByteArrayToString), Bytes(0), i
''''    Else 'ANSI
''''        ByteArrayToString = StrConv(Bytes, vbUnicode)
''''    End If
''''
''''End Function


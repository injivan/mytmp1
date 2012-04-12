Attribute VB_Name = "mdlPublConst"
Option Explicit
Private Const ModulIdString   As String = "mdlPublConst - "
'nFunction = 3001 -> FileExists
'nFunction = 3002 -> DirExists
'nFunction = 3003 -> ExtractPathFile
'nFunction = 3004 -> AddDirSep
'nFunction = 3005 -> FileInUse
'nFunction = 3006 -> Read_Reg_String
'nFunction = 3007 -> Write_Reg_String
'nFunction = 3008 -> RemDirSep

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long



Public Sub ShowErrMesage(ByVal MyError As Object, _
                         ByVal ModIDstr As String, _
                         ByVal nFunction As Long, _
                         ByVal erLine As Long)

Dim ERN As String
Dim ERD As String
    
    ERN = Err.Number
    ERD = Err.Description
    
    '526  = "Грешка"
    '527  = "Номер:"
    '528  = "Описание:"
    '529  = "на ред:"
    'RD.AddToErrLog ModuleN, ERN, ERD, ModIDstr & nFunction, , erLine
    
    'MsgBox RD.GTxt(ModuleN, 526) & " " & ModIDstr & " " & nFunction & vbCrLf & _
           RD.GTxt(ModuleN, 529) & " " & erLine & vbCrLf & _
           RD.GTxt(ModuleN, 527) & " " & ERN & vbCrLf & _
           RD.GTxt(ModuleN, 528) & " " & ERD, vbCritical, RD.GTxt(ModuleN, 526)
           
    If ERN Then
        MsgBox "Error" & " " & ModIDstr & " " & nFunction & vbCrLf & _
               "Row Numberm:" & " " & erLine & vbCrLf & _
               "Err Number:" & " " & ERN & vbCrLf & _
               "Description:" & " " & ERD, vbCritical, "Error"
    End If
End Sub



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


Function FileExists(FileName As String) As Boolean
Const nFunction = 3001
On Error GoTo ErrorHandler
10
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(FileName) And vbDirectory) = 0
Exit Function
ErrorHandler:
'ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function
Function DirExists(DirName As String) As Boolean
Const nFunction = 3002
On Error GoTo ErrorHandler
10
    ' test the directory attribute
    DirExists = GetAttr(DirName) And vbDirectory
    
Exit Function
ErrorHandler:
'ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function

Function ExtractPathFile(ByVal sInPath As String, sOutDir As String, sOutFName As String) As Boolean
Const nFunction = 3003
On Error GoTo ErrorHandler
Dim i As Long
Dim a() As String

10
    sOutDir = vbNullString
    sOutFName = vbNullString
    a = Split(sInPath, "\")
    
    i = UBound(a)
    If i >= 0 Then
        For i = 1 To i - 1
            a(0) = a(0) & "\" & a(i)
        Next i
        sOutDir = a(0) & "\"
        sOutFName = a(i)
    End If
    
Exit Function
ErrorHandler:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function
Public Function AddDirSep(ByVal sPath As String) As String
Const nFunction = 3004
On Error GoTo ErrorHandler

Dim sB  As String

10
    sPath = Trim$(sPath)
    sB = "\"
    If Right$(sPath, 1) = "\" Then sB = vbNullString
    AddDirSep = sPath & sB
    
Exit Function
ErrorHandler:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
    
End Function

Public Function FileInUse(ByRef strFileName As String) As Long
Const nFunction = 3005
On Error GoTo ErH
Dim FileNum As Long

10
    FileNum = FreeFile
    
    Open strFileName For Binary Access Read Write Lock Read Write As FileNum
    Close FileNum
    
    Exit Function
ErH:
Close FileNum
FileInUse = Err.Number
    
    'Select Case Err.Number
    'Case 53 'File not found
    '    MsgBox "The file does not exists.", vbCritical, "clsFSO - FileInUse Error"
    'Case 75 'Path/File Access error
    'Case 70 'Permission Denied
    'End Select
    
End Function


Public Function Read_Reg_String(ByVal sPath As String, _
                                ByRef NomPrm As Long, _
                                ByRef sOutStr As String) As Boolean
Const nFunction = 3006
On Error GoTo ErH
 
Dim sBuf As String
Dim sTmp As String
Dim z    As Long
Dim h    As Long
Dim cf   As cFile

10
    sTmp = AddDirSep(sPath) & sINI_Name
    sOutStr = vbNullString
100

    '  1 -  3 b ASCII - Номер на параметър
    '  4 -  1 b ASCII - "."
    '  5 - 59 b ASCII - Описание на параметъра
    ' 64 -  1 b ASCII - "="
    ' 65 - 64 b ASCII - Данни за параметъра
    '----------------------------------
    'Общ. 128 b
    
    If FileExists(vbNullString) Then z = z
    If FileExists(sTmp) Then
200
         
        Set cf = New cFile
        cf.OpenFile sTmp, 1
    
        z = FileLen(sTmp)
        h = 1
        sBuf = Space$(128)
        Do While z
            If Len(sBuf) > z Then sBuf = Space$(z)
            cf.GetData h, 0, sBuf
            If Val(Left$(sBuf, 3)) = NomPrm Then
                sOutStr = Trim$(Mid$(sBuf, 65))
                Exit Do
            End If
            z = z - 128 - 2 'vbcrlf
            h = h + 128 + 2 'vbcrlf
        Loop
        cf.CloseFile
        Set cf = Nothing
    End If
300
    Read_Reg_String = True

 
Exit Function
ErH:
Set cf = Nothing
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function
Public Function Write_Reg_String(ByVal sPath As String, _
                                 ByRef NomPrm As Long, _
                                 ByRef sInStr As String, _
                                 Optional ByRef sDescrStr As String) As Boolean

Const nFunction = 3007
Dim sTmp As String
Dim sBuf As String
Dim cf   As cFile
Dim z    As Long
Dim h    As Long
Dim x    As Long
10
    '  1 -  3 b ASCII - Номер на параметър
    '  4 -  1 b ASCII - "."
    '  5 - 59 b ASCII - Описание на параметъра
    ' 64 -  1 b ASCII - "="
    ' 65 - 64 b ASCII - Данни за параметъра
    '----------------------------------
    'Общ. 128 b


    sTmp = AddDirSep(sPath) & sINI_Name
    'If FileExists(sTmp) Then
        Set cf = New cFile
        cf.OpenFile sTmp, 1
        x = 0
        z = FileLen(sTmp)
        h = 1
        sBuf = Space$(128)
        Do While z
             
            If Len(sBuf) > z Then sBuf = Space$(z)
            cf.GetData h, 0, sBuf
            If Val(Left$(sBuf, 3)) = NomPrm Then Exit Do
            If x = 0 Then If Len(Trim$(sBuf)) = 0 Then x = h 'празен ред
            z = z - 128 - 2
            h = h + 128 + 2
        Loop
        
        
        Mid$(sBuf, 1, 3) = Right$(Space$(3) & NomPrm, 3)
        Mid$(sBuf, 4, 1) = "."
        Mid$(sBuf, 5, 59) = Left$(sDescrStr & Space$(59), 59)
        Mid$(sBuf, 64, 1) = "="
        Mid$(sBuf, 65) = Left$(sInStr & Space$(64), 64)
        If x Then h = x 'Записвам на празното място
        cf.PutData h, 0, sBuf & vbCrLf
    
        
        cf.CloseFile
        Set cf = Nothing
    'End If
    Write_Reg_String = True
Exit Function
ErH:
Set cf = Nothing
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function

Public Function RemDirSep(ByVal sPath As String) As String
Const nFunction = 3008
On Error GoTo ErrorHandler

Dim i  As Long

10
    sPath = Trim$(sPath)
    i = Len(sPath)
    If Right$(sPath, 1) = "\" Then i = i - 1
    RemDirSep = Mid$(sPath, 1, i)
    
Exit Function
ErrorHandler:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
    
End Function









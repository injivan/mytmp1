VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   240
   ClientTop       =   4920
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   6585
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Const ModulIdString As String = " - Form 1 "
Private Sub Form_Load()
    MyPath = AddDirSep(App.Path)
    read_file
End Sub
Private Function read_file()
Const nFunction = 3004
On Error GoTo ErH

Dim dbIn    As cFile
Dim dbOut   As cFile
Dim sBuf    As String
Set dbIn = New cFile
Dim arr_FileIndex() As Long
Dim tmp_Row_Count   As Long
Dim sOut            As String
Dim l_Out           As Long
Dim fl1             As Boolean
Dim flLF            As Boolean
Dim i               As Long
Dim j               As Long
Dim z               As Long
    Set dbIn = New cFile: Set dbOut = New cFile
    l_Out = 1
    If Not dbIn.OpenFile(MyPath & "090923.pdf", 1) Then GoTo ErH
    If Not dbOut.OpenFile(MyPath & "out.txt", 1) Then GoTo ErH
    dbOut.SetEOF 0
     
    'Функция за запис на началните позиции на редовете от файла във масив
    If Not fill_arrFileInd(arr_FileIndex, dbIn, tmp_Row_Count) Then GoTo ErH
    
    For i = 0 To tmp_Row_Count
        j = arr_FileIndex(i + 1) - arr_FileIndex(i) - 1 '-1 за да махна vbCrLf
        sBuf = Space$(j)
        If Not dbIn.GetData(arr_FileIndex(i), 0, sBuf) Then Err.Raise 1000, , "dbIn.GetData( arr_FileIndex(i),0,sBuf)"
        sBuf = Trim$(sBuf)
        j = Len(sBuf) 'Проверка за празен ред
        flLF = False: fl1 = False
        If j Then 'проверка за коментар
            For z = 1 To j
                Select Case Mid$(sBuf, z, 1)
                Case "(": fl1 = True: sOut = vbNullString
                Case ")":
                    fl1 = False
                    If Len(sOut) Then
                        If Not dbOut.PutData(l_Out, 0, sOut) Then GoTo ErH
                        l_Out = l_Out + Len(sOut)
                        sOut = vbNullString
                        flLF = True
                    End If
                Case Else: If fl1 Then sOut = sOut & Mid$(sBuf, z, 1)
                End Select
            Next z
        End If
        If flLF Then
            If Not dbOut.PutData(l_Out, 0, vbLf) Then GoTo ErH
            l_Out = l_Out + 1
        End If
    Next i
     
    Set dbIn = Nothing
    Set dbOut = Nothing
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function


Private Function fill_arrFileInd(ByRef arrFileInd() As Long, _
                                 ByRef dbBAK As cFile, _
                                 ByRef RowCount As Long) As Boolean
Const nFunction = 3005
On Error GoTo ErH
Dim sName       As String
Dim a           As Long
Dim b           As Long
Dim i           As Integer
Dim j           As Long
Dim rr()        As String



10
    'Функция за запис на началните позиции на редовете от файла във масив
    'Чета фаила и търся CRLF
    sName = Space$(1024)
    RowCount = -1
    a = 1&
    b = dbBAK.RecCount + 1 '!!! dali nqma da
    ReDim arrFileInd(100) As Long
    j = 0
    arrFileInd(j) = 1
    Do While b > 1
        If Len(sName) > b Then sName = Space$(b)
        If Not dbBAK.GetData(a, 0&, sName) Then Stop: GoTo ErH
300     i = 1
        Do
            i = InStr(i, sName, vbLf)
            If i Then
                j = j + 1
                If j > UBound(arrFileInd) Then ReDim Preserve arrFileInd(j + 100)
                arrFileInd(j) = (a - 1) + (i + 1)
                i = i + 1
            End If
        Loop While i > 0
        b = b - (Len(sName))
        a = a + Len(sName)
    Loop
    RowCount = j 'Брой редове в файла (Разделени с Lf)
400
    'Сега оправям последния ред
    j = j + 1
    If j > UBound(arrFileInd) Then ReDim Preserve arrFileInd(j + 1)
    'Последният ред да е дълък колкото най дългия
    arrFileInd(j) = (dbBAK.RecCount + 1 + 1)
    
    fill_arrFileInd = True

Exit Function
ErH:


ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function


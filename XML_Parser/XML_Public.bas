Attribute VB_Name = "XML_Public"
Public flFastRead     As Boolean 'Флаг за бързо парсване
Public flTMPFastRead  As Boolean 'Флаг за бързо парсване

Private m_varPath     As String
Private m_varFileName As String
Private LastA         As Long


Public Function SetDataDims(ByVal a As Long, _
                            ByVal sPath As String, _
                            ByVal sFile As String) As Boolean
Const nFunction = 3001
On Error GoTo ErH
10
    LastA = a
    m_varPath = sPath
    m_varFileName = sFile
    SetDataDims = True
    
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear

End Function

Public Function GetNewData(ByVal lPoz As Long, _
                           ByRef Data() As Byte, _
                           Optional ByRef FirsByte As Long) As Boolean
Const nFunction = 3002
On Error GoTo ErH

Dim cDBO        As cFile
Dim bData()     As Byte
Dim a           As Long
Dim b           As Long

10
    Set cDBO = New cFile
    If Not cDBO.OpenFile(m_varPath, m_varFileName) Then GoTo ErH
    
    a = LastA - lPoz + 1&
    b = 8 * 1024& - 1& + lPoz
    ReDim bData(b)
    LastA = a + b '+ 1
    b = cDBO.RecCount - a
    
    '552  = "Грешка при зареждане на Файл"
    If b < 0 Then Err.Raise 777, , "Грешка при зареждане на Файл" & vbCrLf & m_varPath & "\" & m_varFileName

    If UBound(bData) > b Then ReDim bData(b): LastA = a + b
    If Not cDBO.GetData(a, 0, bData) Then GoTo ErH
600
    cDBO.CloseFile
    Set cDBO = Nothing
    FirsByte = a
700
    Data = bData
    
    GetNewData = True

ErH:

ShowErrMesage Err, ModulIdString, nFunction, Erl
 
If Not (cDBO Is Nothing) Then
    cDBO.CloseFile
    Set cDBO = Nothing
End If
     
End Function

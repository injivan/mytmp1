VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


'nFunction = 3000 -> LoadData
'nFunction = 3001 -> SaveXML
'nFunction = 3002 -> m_RootElement_ParsedBytes
'nFunction = 3003 -> Set_PR_BAR_Data

Public Enum eTypeObr
    e_Obr_Load_Maket_Full = 1
    e_Obr_Save_Maket_Full = 1
End Enum

Private Const ModulIdString      As String = "c XMLDocument - "

Private WithEvents m_RootElement As CXmlElement
Attribute m_RootElement.VB_VarHelpID = -1
Private m_varPath                As String
Private m_varFileName            As String
Private m_varByteCount           As Long

Private m_varFormN               As Long
Private m_varPrBarName           As String
Private m_var_AbortOperation     As Boolean 'Флаг за спиране на работата



  
 
 
Public Sub AppendNode(ByVal Item As CXmlElement)
    Call m_RootElement.AppendNode(Item)
End Sub

'''''Public Function LoadData_Old(ByVal sFilePath_And_Name As String) As Boolean
'''''Const nFunction = 5000
'''''On Error GoTo ErH
'''''
'''''Dim bData()     As Byte
'''''Dim a           As Long
'''''Dim b           As Long
'''''Dim hFile       As Integer
'''''10
'''''
'''''    hFile = FreeFile
'''''    Open sFilePath_And_Name For Binary As hFile
'''''    ReDim bData(LOF(hFile))
'''''    Get hFile, 1, bData
'''''
'''''    Close hFile
'''''    hFile = 0
'''''    Call m_RootElement.Parse(bData(), , False)
'''''
'''''    LoadData_Old = True
'''''
'''''ErH:
''''''ShowErrMesage Err, ModulIdString, nFunction, Erl
'''''Err.Clear
'''''
'''''If hFile Then Close hFile
'''''
'''''
'''''End Function



' This routine does not check for the existance of a Document Prolog (I just don't need it yet)
' Also It only accepts ANSI Characters (Hope you speak english)
Public Property Get ByteCount() As Long
    ByteCount = m_varByteCount
End Property
 

Public Function LoadData(ByVal sFilePath_And_Name As String, _
                         ByVal flFastParse As Boolean, _
                         ByVal flShowErrMsg As Boolean) As Boolean
Const nFunction = 3000
On Error GoTo ErH

Dim cDBO        As cFile
Dim bData()     As Byte
Dim a           As Long
Dim b           As Long
Dim x           As Long
Dim d           As Double
10
    If Not FileExists(sFilePath_And_Name) Then
        '523  = "Избраният макет не съществува или нямате достъп до него"
        '!!!MsgBox ModuleN, RD.GTxt(ModuleN, 523), MBS_Exclamation
        GoTo ErH
    End If
    
100
    flFastRead = flFastParse
    flTMPFastRead = flFastRead
    Call ExtractPathFile(sFilePath_And_Name, m_varPath, m_varFileName)
    
    Set cDBO = New cFile ' DB.GetNewDBOther
200
    a = -1&
    d = Date + Time + TimeSerial(0, 0, 6)
    Do While d >= CDbl(Date + Time)
        If FileInUse(sFilePath_And_Name) = 0 Then
            a = 255&
            If Not cDBO.OpenFile(m_varPath & m_varFileName, 1) Then GoTo ErH
            Exit Do
        End If
        DoEvents
    Loop
    If a = -1 Then Err.Raise 7777
    m_varByteCount = cDBO.RecCount
    
300
    a = 1&
    b = 8 * 1024& - 1&
    x = a + b
    ReDim bData(b)
    b = cDBO.RecCount - 1
    If UBound(bData) > b Then ReDim bData(b): x = a + b
    If Not cDBO.GetData(a, 1, bData) Then GoTo ErH

700
    If m_varFormN Then
        
        If m_varFormN = 14 Then
            'Прогрес за формата frmCube
            
        Else
            'Ще връщам данни за прогрес бар
            'Трябва да заредя формата с текстове
'''            With SN(m_varFormN, ModuleN)
'''                .Itam(m_varPrBarName & "_Bar").Max = m_varByteCount
'''            End With
        End If
            
    End If
    If Not SetDataDims(x, m_varPath, m_varFileName) Then GoTo ErH
    If Not m_RootElement.Parse(bData(), , False, m_var_AbortOperation) Then GoTo ErH
    
800 'Затвярям го след парсвянето за да сработи
    'FS.FileInUse ако друг реши да работи
    cDBO.CloseFile
    Set cDBO = Nothing
    LoadData = True
    
ErH:

If Err.Number = 7777 Then
    If flShowErrMsg Then
        '530  = "Данните не може да се заредят.
        'В момента се използват от друг потребител."
        MsgBox "Данните не може да се заредят." & _
            "В момента се използват от друг потребител.", vbCritical
    End If
Else
    ShowErrMesage Err, ModulIdString, nFunction, Erl
End If
Err.Clear

If Not (cDBO Is Nothing) Then
    cDBO.CloseFile
    Set cDBO = Nothing
End If
    
     

End Function

Public Property Get NewEnum() As IUnknown
    Set NewEnum = m_RootElement.NewEnum
End Property

Public Property Get Node(ByVal Index As Long) As CXmlElement
    Set Node = m_RootElement.Node(Index)
End Property

Public Sub RemoveNode(ByVal Index As Long)
    Call m_RootElement.RemoveNode(Index)
End Sub


Public Property Get Root() As CXmlElement
    Set Root = m_RootElement
End Property
Public Property Let Root(ByVal RootElement As CXmlElement)
    Set m_RootElement = RootElement
End Property


Public Property Let FilePath(ByVal sData As String)
    m_varPath = Trim$(sData)
End Property
Public Property Get FilePath() As String
    FilePath = m_varPath
End Property

Public Property Let FileName(ByVal sData As String)
    m_varFileName = Trim$(sData)
End Property
Public Property Get FileName() As String
    FileName = m_varFileName
End Property


Public Function Serialize() As String
    Serialize = m_RootElement.Serialize
End Function

Public Function SaveXML() As Boolean
Const nFunction = 3001
 On Error GoTo ErH
 

Dim cDBO        As cFile
Dim sData       As String
Dim sBuf        As String
Dim a           As Long
Dim b           As Long
Dim d           As Double
Dim sFileTMP    As String
Dim sDirTMP     As String

Dim LenData     As Long

10
    '!!!Ако има мои файлове да ги махне
    
    If m_varPath = vbNullString Or m_varFileName = vbNullString Then
        '1011 = "Няма пътека или име на файл"
        Err.Raise 1000, , "Няма пътека или име на файл"
    End If
    'Ще Изграждам новия макет в TMP директорията
    'Накрая го копирам в дир-а с макети
    'Ако спре тока си има работещ макет
    '   Ще имам 1 ТМП файл, който да трия
    
    sDirTMP = App.Path
    sFileTMP = "XX_" & m_varFileName & "TMP"
50
    Set cDBO = New cFile
    If Not cDBO.OpenFile(sDirTMP & sFileTMP, 1) Then GoTo ErH
60
    sData = Serialize
70
    If m_varFormN Then
        If m_varFormN = 14 Then
            'Прогрес за формата frmCube
        Else
            'Функция за запълване на прогрес бар
            'при зареждане на макет
            'Функция за промянана данните на прогрес бара
'''            With SN(m_varFormN, ModuleN)
'''                '551  = "Запис на макета"
'''                .Itam(m_varPrBarName & "_Bar").Max = Len(sData)
'''                .Itam(m_varPrBarName & "_Bar").Value = 0
'''                .Itam(m_varPrBarName & "_TopLbl").Caption = RD.GTxt(ModuleN, 551)
'''            End With
        End If
    End If
    Call cDBO.SetEOF(1)
90

100
    a = 1&
    b = 8 * 1024&
    sBuf = Space$(b)
    b = Len(sData)
    LenData = b
    Do While b
        d = d + 1
        If Len(sBuf) > b Then sBuf = Space$(b)
        Mid$(sBuf, 1) = Mid$(sData, a)
        If Not cDBO.PutData(a, 0, sBuf) Then GoTo ErH
600
        b = b - Len(sBuf)
        a = a + Len(sBuf)
        
        'Ако Има флаг за показване на прогрес бар
        If m_varFormN Then
            If m_varFormN = 14 Then
'''                'Прогрес за формата frmCube
'''                If Not (frmCube Is Nothing) Then
'''                    'Не съм направил викането за този ред
'''                    frmCube.UpdateProgressBar "'!!!", a, LenData
'''                End If
            Else
                'Функция за запълване на прогрес бар
                'при зареждане на макет
                'Функция за промянана данните на прогрес бара
'''                If fl_LoadMaket Then
''''''                    With SN(m_varFormN, ModuleN)
''''''                        .Itam(m_varPrBarName & "_Bar").Value = a
''''''                        .Itam(m_varPrBarName & "_Bottlbl_Data").Caption = a
''''''                    End With
'''                    If (d Mod 2 = 0) Then DoEvents
'''
'''                Else
'''                    'Натиснат е Cancel
'''                    cDBO.CloseFile
'''                    Set cDBO = Nothing
'''                    sBuf = sDirTMP
'''                    AddDirSep sBuf
'''                    sBuf = sBuf & sFileTMP
'''                    Kill sBuf
'''                    GoTo ErH
'''                End If
            
            End If
        End If
        
    Loop
    cDBO.CloseFile
    Set cDBO = Nothing
800 'Макета е в тмп-то
    'сега да покрия стария макет в директ. с макети
    
    sBuf = m_varPath
    AddDirSep sBuf
    sBuf = sBuf & m_varFileName
    
    sData = sDirTMP
    AddDirSep sData
    sData = sData & sFileTMP
    
    
    a = -1&
    d = Date + Time + TimeSerial(0, 0, 6)
    Do While d >= CDbl(Date + Time)
        If FileInUse(sBuf) Then
            If FileExists(sBuf) Then Kill sBuf
            Name sData As sBuf
            a = 255&
            Exit Do
        End If
        DoEvents
    Loop
    If a = -1 Then Err.Raise 7777
    SaveXML = True

ErH:

If Err.Number = 7777& Then
    '531  = "Данните не може да се запишат.
    'В момента се използват от друг потребител."
    MsgBox "Данните не може да се запишат." & _
            "В момента се използват от друг потребител.", vbCritical
Else
    ShowErrMesage Err, ModulIdString, nFunction, Erl
End If
Err.Clear
If Not (cDBO Is Nothing) Then
    cDBO.CloseFile
    Set cDBO = Nothing
End If
sData = sDirTMP
AddDirSep sData
sData = sData & sFileTMP
If FileExists(sData) Then Kill sData

 
End Function

Private Sub Class_Initialize()
    Set m_RootElement = New CXmlElement
End Sub
Private Sub m_RootElement_ParsedBytes(ByVal NomByte As Long, _
                                      ByRef fl_Exit As Boolean)
Const nFunction = 3002
On Error GoTo ErH
10
''''    'Pri Zaregdane na maket'
''''    If m_varFormN Then
''''
''''        If m_varFormN = 14 Then
''''            'Прогрес за формата frmCube
''''            If Not (frmCube Is Nothing) Then
''''                frmCube.UpdateProgressBar vbNullString, NomByte, m_varByteCount
''''            End If
''''        Else
''''            'Функция за запълване на прогрес бар
''''            'при зареждане на макет
''''            'Функция за промянана данните на прогрес бара
''''            If fl_LoadMaket Then
''''                With SN(m_varFormN, ModuleN)
''''                    .Itam(m_varPrBarName & "_Bar").Value = NomByte
''''                    .Itam(m_varPrBarName & "_Bottlbl_Data").Caption = NomByte
''''                End With
''''            End If
''''            fl_Exit = Not fl_LoadMaket
''''        End If
''''    End If
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub

Public Function Set_PR_BAR_Data(ByVal FormI_1 As Long, _
                                ByVal prBarName As String) As Boolean
Const nFunction = 3003
On Error GoTo ErH
10
 
 m_varFormN = FormI_1
 m_varPrBarName = prBarName
 
 Set_PR_BAR_Data = True

Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
 
End Function







 


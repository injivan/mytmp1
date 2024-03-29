VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSetings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const ModulIdString   As String = "clsSetings - "
'nFunction = 3000 -> Get_Init_Data
'nFunction = 3001 -> Set_Init_Data


Private Enum sRedData
    RD_DirPath_1 = 1    '������ �� ����� � ����� 128 �������
    RD_DirPath_2 = 2
    
    RD_PicFile_1 = 3    '���� �� ���������� �� ��� 128 �������
    RD_PicFile_2 = 4
    
    RD_OutDir_1 = 5     '������ �� ������� ����
    RD_OutDir_2 = 6     '������ �� ������� ����
    
    RD_OutFile = 7      '��� �� ������� ����
    
    RD_OutFileView = 8  '������ �� ��� ���� 0-Full,1-Folders,2-Files
End Enum




Private MePath As String '�������� �� �����

'local variable(s) to hold property value(s)
Private mvarsDirPath    As String '�������� �������� ������
Private mvarsOutDir     As String '������ �� ���. ����
Private mvarsOutFile    As String '��� �� ���. ����
Private mvarsPicFile    As String '������ �� ���������� �� ���
Private mvarsFileView   As String '������ �� ��� ���� 0-Full,1-Folders,2-Files


Private Function Set_Init_Data(ByVal a As sRedData) As Boolean
Const nFunction = 3001
On Error GoTo ErH
Dim sBuf As String
Dim sTmp As String
Dim i    As Long
Dim z    As Long
Dim j    As Long
Dim k As Long
    '������� �� ����� �� ������� �� ������������ ���������
    '������ �� ���. ����
    
    '��� �� ���������� �� ���
    '����� �� �������� : �����, �����, ���� � �������� �� ������
    '�� ����� ������� : �����, �����, ���� � �������� �� ������
    
100
    Select Case a
    
'''    'Marckt_Message 1-2
'''    '����� �� ��������
'''    Case RD_Marckt_Message_1 To RD_Marckt_Message_2
'''        i = 64: z = 1
'''        sTmp = sMarckt_Message
'''        For j = 1 To 2
'''            If Len(sTmp) < i Then i = Len(sTmp)
'''            If i Then
'''                sBuf = Mid$(sTmp, 1, i)
'''                z = z + i
'''                sTmp = Mid$(sTmp, z)
'''            End If
'''            If j = 1 Then k = RD_Marckt_Message_1
'''            If j = 2 Then k = RD_Marckt_Message_2
'''
'''            Write_Reg_String MePath, k, sBuf, "Marckt_Message - " & j
'''            sBuf = vbNullString
'''
'''        Next j
    
    
    Case RD_PicFile_1 To RD_PicFile_2
200     '��� �� ���������� �� ���
        i = 64: z = 1
        sTmp = mvarsPicFile
        For j = 1 To 2
            If Len(sTmp) < i Then i = Len(sTmp)
            If i Then
                sBuf = Mid$(sTmp, z, i)
                z = z + i
                sTmp = Mid$(sTmp, z)
            End If
            
            If j = 1 Then k = RD_PicFile_1
            If j = 2 Then k = RD_PicFile_2
            
            Write_Reg_String MePath, k, sBuf, "PicFilePath - " & j
            sBuf = vbNullString
        Next j
       
    Case RD_DirPath_1 To RD_DirPath_2
300     '������ �� ����� � �����
        i = 64: z = 1
        sTmp = mvarsDirPath
        For j = 1 To 2
            If Len(sTmp) < i Then i = Len(sTmp)
            If i Then
                sBuf = Mid$(sTmp, z, i)
                z = z + i
                sTmp = Mid$(sTmp, z)
            End If
            
            If j = 1 Then k = RD_DirPath_1
            If j = 2 Then k = RD_DirPath_2
            
            Write_Reg_String MePath, k, sBuf, "HistFilePath - " & j
            sBuf = vbNullString
        Next j
     
    Case RD_DirPath_1 To RD_OutDir_2
400     '������ �� ���. ����
         i = 64: z = 1
        sTmp = mvarsOutDir
        For j = 1 To 2
            If Len(sTmp) < i Then i = Len(sTmp)
            If i Then
                sBuf = Mid$(sTmp, z, i)
                z = z + i
                sTmp = Mid$(sTmp, z)
            End If
            
            If j = 1 Then k = RD_OutDir_1
            If j = 2 Then k = RD_OutDir_2
            
            Write_Reg_String MePath, k, sBuf, "Output Path - " & j
            sBuf = vbNullString
        Next j
    Case RD_OutFile
500     '��� �� ��� ����
        Write_Reg_String MePath, RD_OutFile, mvarsOutFile, "Output File Name "
    
    Case RD_OutFileView
        '������ �� ��� ���� 0-Full,1-Folders,2-Files
        Write_Reg_String MePath, RD_OutFileView, mvarsFileView, "Output File View "
        
    
    Case Else
800     Err.Raise 1000, , "���� ���� " & a & " �� " & ModulIdString
    End Select
    
    Set_Init_Data = True
    
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function

Public Function GetInitData() As Boolean
Const nFunction = 3000
On Error GoTo ErH
Dim sBuf As String
Dim i    As Long
    '������� �� ������ �� ������� �� ������������ ���������
    '������ �� ���. ����
100
    
150 ''������ �� ����� � ����� 128 �������
    '����� �� ��������
    Read_Reg_String MePath, RD_DirPath_1, sBuf
    mvarsDirPath = sBuf
    Read_Reg_String MePath, RD_DirPath_2, sBuf
    mvarsDirPath = mvarsDirPath & sBuf
    If Not DirExists(mvarsDirPath) Then mvarsDirPath = vbNullString
    If Len(mvarsDirPath) = 0 Then mvarsDirPath = MePath
    
200 '��� �� ���������� �� ���
    Read_Reg_String MePath, RD_PicFile_1, sBuf
    mvarsPicFile = sBuf
    Read_Reg_String MePath, RD_PicFile_2, sBuf
    mvarsPicFile = mvarsPicFile & sBuf
    If Not FileExists(mvarsPicFile) Then mvarsPicFile = vbNullString
    
    
300 '������ �� ������� ����
    Read_Reg_String MePath, RD_OutDir_1, sBuf
    mvarsOutDir = Trim$(sBuf)
    Read_Reg_String MePath, RD_OutDir_2, sBuf
    mvarsOutDir = mvarsOutDir & Trim$(sBuf)
    
    If Not DirExists(mvarsOutDir) Then mvarsOutDir = vbNullString
    If Len(mvarsOutDir) = 0 Then mvarsOutDir = MePath

400 'Ime �� ������� ����
    Read_Reg_String MePath, RD_OutFile, sBuf
    
    i = InStr(1, sBuf, ".")
    If i Then sBuf = Mid$(sBuf, 1, i - 1)
    
    mvarsOutFile = Trim$(sBuf)
    If Len(mvarsOutFile) = 0 Then mvarsOutFile = "List_Out"
    
500
    '������ �� ��� ���� 0-Full,1-Folders,2-Files
    Read_Reg_String MePath, RD_OutFileView, sBuf
    mvarsFileView = Val(Trim$(sBuf))
    
    
    GetInitData = True
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear



End Function


Public Property Let sFileView(ByVal vData As String)
    mvarsFileView = vData
    Set_Init_Data RD_OutFileView
End Property
Public Property Get sFileView() As String
    sFileView = mvarsFileView
End Property



Public Property Let sPicFile(ByVal vData As String)
    mvarsPicFile = vData
    Set_Init_Data RD_PicFile_1
End Property
Public Property Get sPicFile() As String
    sPicFile = mvarsPicFile
End Property


Public Property Let sOutFile(ByVal vData As String)
Dim i As Long
    
    i = InStr(1, vData, ".")
    If i Then vData = Mid$(vData, 1, i - 1)
    
    mvarsOutFile = vData
    Set_Init_Data RD_OutFile
End Property
Public Property Get sOutFile() As String
    sOutFile = mvarsOutFile
End Property


Public Property Let sOutDir(ByVal vData As String)
    mvarsOutDir = vData
    Set_Init_Data RD_OutDir_1
End Property
Public Property Get sOutDir() As String
    sOutDir = AddDirSep(mvarsOutDir)
End Property


Public Property Let sDirPath(ByVal vData As String)
Attribute sDirPath.VB_Description = "��������� �������� ������"
    mvarsDirPath = vData
    Set_Init_Data RD_DirPath_1
End Property
Public Property Get sDirPath() As String
    sDirPath = AddDirSep(mvarsDirPath)
End Property



Private Sub Class_Initialize()
    MePath = DirAppl
End Sub

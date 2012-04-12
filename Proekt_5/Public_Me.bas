Attribute VB_Name = "Public_Me"
Option Explicit

Private Const ModulIdString   As String = "mdlPublicMe - "

'nFunction = 3000 -> Get_Init_Data
'nFunction = 3001 -> Set_Init_Data
'nFunction = 3002 ->
'nFunction = 3003 ->


Public MePath           As String '������ �� ���-��
Public sHistFilePath    As String '������ �� ����� � �����
Public sPicFilePath     As String '������ �� ���������� �� ���
Public sMarckt_Message  As String '����� �� ��������
Public lStartRec        As Long   '������� ����� �� ������
Public lEndRec          As Long   '�������� ����� �� ������
Public sOutFilePath     As String '������ �� ������� ����
Public sOutFileName     As String '��� �� ������� ����
Public Const sINI_Name  As String = "E_Mail.ini"

Public Enum sRedData
    RD_HistFilePath_1 = 1   '������ �� ����� � ����� 128 �������
    RD_HistFilePath_2 = 2
    
    RD_Marckt_Message_1 = 3 '����������� � 128 �������
    RD_Marckt_Message_2 = 4 '������ � � ������ �����
    
    RD_NameFon_1 = 5        '���� �� ���������� �� ��� 128 �������
    RD_NameFon_2 = 6
    
    RD_OutPutFilePath_1 = 7   '������ �� ������� ����
    RD_OutPutFilePath_2 = 11  '������ �� ������� ����
    
    RD_OutPutFileName = 8   '��� �� ������� ����
    
    RD_StartRec = 9         '������� ����� �� ������
    RD_EndRec = 10          '�������� ����� �� ������
    
End Enum



Public Function Get_Init_Data() As Boolean
Const nFunction = 3000
On Error GoTo ErH
Dim sBuf As String
    '������� �� ������ �� ������� �� ������������ ���������
    '������ �� ���. ����

100 '����� �� �������� : �����, �����, ���� � �������� �� ������
    'Marckt_Message 1-2
    '����� �� ��������
    Read_Reg_String MePath, RD_Marckt_Message_1, sBuf
    sMarckt_Message = sBuf
    Read_Reg_String MePath, RD_Marckt_Message_2, sBuf
    sMarckt_Message = sMarckt_Message & sBuf
    
150 ''������ �� ����� � ����� 128 �������
    '����� �� ��������
    Read_Reg_String MePath, RD_HistFilePath_1, sBuf
    sHistFilePath = sBuf
    Read_Reg_String MePath, RD_HistFilePath_2, sBuf
    sHistFilePath = sHistFilePath & sBuf
    
200 '��� �� ���������� �� ���
    Read_Reg_String MePath, RD_NameFon_1, sBuf
    sPicFilePath = sBuf
    Read_Reg_String MePath, RD_NameFon_2, sBuf
    sPicFilePath = sPicFilePath & sBuf
    
300 '������ �� ������� ����
    Read_Reg_String MePath, RD_OutPutFilePath_1, sBuf
    sOutFilePath = Trim$(sBuf)
    Read_Reg_String MePath, RD_OutPutFilePath_2, sBuf
    sOutFilePath = sOutFilePath & Trim$(sBuf)
    sOutFilePath = "C:\desde"
    If Not DirExists(sOutFilePath) Then sOutFilePath = vbNullString
    
    If Len(sOutFilePath) = 0 Then sOutFilePath = MePath

400 'Ime �� ������� ����
    Read_Reg_String MePath, RD_OutPutFileName, sBuf
    sOutFileName = Trim$(sBuf)
    If Len(sOutFileName) = 0 Then sOutFileName = "Pred_Out.txt"
    
500 '������� ����� �� ������
    Read_Reg_String MePath, RD_StartRec, sBuf
    lStartRec = Val(sBuf)
    If lStartRec <= 0 Then lStartRec = 1
    
600 '�������� ����� �� ������
    Read_Reg_String MePath, RD_EndRec, sBuf
    lEndRec = Val(sBuf)
    If lEndRec <= 0 Then lEndRec = 0
        
    Get_Init_Data = True
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear

End Function




Public Function Set_Init_Data(ByRef a As sRedData) As Boolean
Const nFunction = 3001
On Error GoTo ErH
Dim sBuf As String
Dim sTmp As String
Dim i    As Long
Dim z    As Long
Dim j    As Long
Dim k As Long
    '������� �� ������ �� ������� �� ������������ ���������
    '������ �� ���. ����
    
    '��� �� ���������� �� ���
    '����� �� �������� : �����, �����, ���� � �������� �� ������
    '�� ����� ������� : �����, �����, ���� � �������� �� ������
    
100
    Select Case a
    
    'Marckt_Message 1-2
    '����� �� ��������
    Case RD_Marckt_Message_1 To RD_Marckt_Message_2
        i = 64: z = 1
        sTmp = sMarckt_Message
        For j = 1 To 2
            If Len(sTmp) < i Then i = Len(sTmp)
            If i Then
                sBuf = Mid$(sTmp, 1, i)
                z = z + i
                sTmp = Mid$(sTmp, z)
            End If
            If j = 1 Then k = RD_Marckt_Message_1
            If j = 2 Then k = RD_Marckt_Message_2
            
            Write_Reg_String MePath, k, sBuf, "Marckt_Message - " & j
            sBuf = vbNullString
           
        Next j
    
    
    Case RD_NameFon_1 To RD_NameFon_2
200     '��� �� ���������� �� ���
        i = 64: z = 1
        sTmp = sPicFilePath
        For j = 1 To 2
            If Len(sTmp) < i Then i = Len(sTmp)
            If i Then
                sBuf = Mid$(sTmp, z, i)
                z = z + i
                sTmp = Mid$(sTmp, z)
            End If
            
            If j = 1 Then k = RD_NameFon_1
            If j = 2 Then k = RD_NameFon_2
            
            Write_Reg_String MePath, k, sBuf, "PicFilePath - " & j
            sBuf = vbNullString
        Next j
       
    Case RD_HistFilePath_1 To RD_HistFilePath_2
300     '������ �� ����� � �����
        i = 64: z = 1
        sTmp = sHistFilePath
        For j = 1 To 2
            If Len(sTmp) < i Then i = Len(sTmp)
            If i Then
                sBuf = Mid$(sTmp, z, i)
                z = z + i
                sTmp = Mid$(sTmp, z)
            End If
            
            If j = 1 Then k = RD_HistFilePath_1
            If j = 2 Then k = RD_HistFilePath_2
            
            Write_Reg_String MePath, k, sBuf, "HistFilePath - " & j
            sBuf = vbNullString
        Next j
     
    Case RD_OutPutFilePath_1 To RD_OutPutFilePath_2
400     '������ �� ���. ����
         i = 64: z = 1
        sTmp = sOutFilePath
        For j = 1 To 2
            If Len(sTmp) < i Then i = Len(sTmp)
            If i Then
                sBuf = Mid$(sTmp, z, i)
                z = z + i
                sTmp = Mid$(sTmp, z)
            End If
            
            If j = 1 Then k = RD_OutPutFilePath_1
            If j = 2 Then k = RD_OutPutFilePath_2
            
            Write_Reg_String MePath, k, sBuf, "Output Path - " & j
            sBuf = vbNullString
        Next j
    Case RD_OutPutFileName
500     '��� �� ��� ����
        Write_Reg_String MePath, RD_OutPutFileName, sOutFileName, "Output File Name "
    
    End Select
    
    Set_Init_Data = True
    
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear

End Function

Attribute VB_Name = "mdlResFile"
Option Explicit

'fN = 3000 -> ReadRegFile
'fN = 3001 -> AddText

'fN = 3003 -> GetsText

Private Const mdlName As String = "mdlResFile - "
Private Const ub      As Integer = 32767

Private Type strLens
    lStart  As Long
    lLong   As Long
    CharSet As Long
End Type


Private Type sMdlStrInd
    arInd() As Long
    sStr()  As strLens
    maxID   As Integer
End Type
Private inModule()  As Integer '����� � ������ �� ������
                               '� ���������� ������� �� ������ �� �������� �������
                               '��� ����� arModule() As sMdlStrInd
                               '��� ������� ����� ��� ���������
                               '� ������� ����� ���������� �� ����� ����� ���
                               '������� ������
Private arModule()  As sMdlStrInd
Dim max_arModule    As Integer
'===========================================
'������
'����� ����� 5432 �� ����� 12
'����� ����� � � = inModule(12)
'� � ������� �� ������ arModule()
'i = arModule(�).arInd(5432)
'���� i ���� �������� �� ������ ��� ���������
'sString = arModule(�).sStr(i)
'===========================================



Public Function ReadRegFile(ByVal ModuleN As Integer) As Boolean
Const fN = 3000
On Error GoTo ErH

Dim sPath       As String
Dim frFile      As Long
Dim sRow        As String
Dim flRedText   As Boolean
Dim flFindSekQuote As Boolean
Dim flFindCharSet  As Boolean
'
Dim bRead       As Long
Dim b           As Long
Dim a           As Long
'
'
Dim StartB       As Long
'
Dim i           As Long
Dim sBuf        As String
Dim iNomMod     As Integer '����� �� �����
Dim iTxtInd     As Integer '������ �� �����
Dim sText       As String  '�����

Dim iAppCarset  As Integer 'CharSet �� ������ ����������
Dim iMdlCharSet As Integer 'CharSet �� ������, ��� ���� �� ����� ���� �� ������������
Dim iTextCarset As Integer 'CharSet �� ������


Dim flMdlCarset As Boolean



10
'������� �� ������ �� ��������� ����

    If max_arModule Then
        If inModule(ModuleN) Then
            ReadRegFile = True
            Exit Function
        End If
    End If

    sPath = mdl_Const.appPath
    sPath = AddDirSep(sPath) & mdl_Const.res_file_Name
        
    '��� �� ����� ����
    If Not FileExist(sPath) Then
        MsgBox "������ ��� ������ � ����:" & vbCrLf & _
               sPath & vbCrLf & _
               "��������� ���� ������"
        GoTo ErH
    End If
    frFile = FreeFile
    
    ReDim inModule(ub) As Integer
    max_arModule = 0
100
    bRead = 0
    
    Open sPath For Input As #frFile
        Do While Not EOF(frFile)
            Line Input #frFile, sRow
            b = Len(sRow) + 2
            sBuf = LTrim$(Left$(sRow, 1))
200
            Select Case sBuf
            Case "#" '���� � ��������
            Case "[" '����� �� �����
                i = InStr(1, sRow, "]") - 1
                If i > 1 Then
                    sBuf = Mid$(Trim$(sRow), 1, i)
                    a = Val(Trim$(Mid$(sBuf, 2)))   '����� �� �����
                    If a Then iNomMod = a
                End If
            Case "<" 'CharSet
                i = InStr(1, sRow, ">") - 1
                If i > 1 Then
                    sBuf = Mid$(Trim$(sRow), 1, i)
                    a = Val(Trim$(Mid$(sBuf, 2)))   '����� �� CharSet
                    If a Then
                        If flFindCharSet Then
                            '������ � �������� (�����, ������ � �����)
                            '�������� CharSet �� ������
                            iTextCarset = iMdlCharSet
                            If a Then iTextCarset = a
                        Else '������ �� �
                            'CharSet �� ������ ���������� ���
                            'CharSet �� modula
                        
                            If flMdlCarset Then
                                'True - �������� � �� ������ ����������
                                '���� - CharSet �� modula
                                iMdlCharSet = iAppCarset
                                If a Then iMdlCharSet = a
                            Else
                                flMdlCarset = True
                                iAppCarset = a
                            End If
                        End If
                        
                    End If
                End If
                
300
            Case Else
ReadText:
                If flFindCharSet Then
                    '�������� � �� �� ����� Charset, � ���� ������ �� � �������
                    '�������� ������ � ������
                    If iTextCarset = 0 Then
                        If iMdlCharSet = 0 Then iMdlCharSet = iAppCarset
                        iTextCarset = iMdlCharSet
                    End If
                    
                    If Not AddText(iNomMod, iTxtInd, Len(sText) - 2, StartB, iTextCarset) Then GoTo ErH
                    iTextCarset = 0
                    flFindCharSet = False
                    iTxtInd = 0
                    sText = vbNullString
                End If
                If flRedText Then
                    '������ �� �����
                    If flFindSekQuote Then
                        i = 1: a = 1
                        Do '����� ��������� �������
                            i = InStr(a, sRow, Chr$(34))
                            If i Then
                                '�������� � �������
                                i = i - a + 1
                                sText = sText & Mid$(sRow, a, i)
                                a = i + a
500
                                '���� ���� � ����������� �������
                                sBuf = Right$(sText, 2)
                                If Not (Mid$(sBuf, 1, 1) = "\") Then
                                    
                                    flFindCharSet = True
                                    flFindSekQuote = False
                                    flRedText = False
                                    
                                     
                                    Exit Do
                                End If
                            Else
                                '�� � �������� ������� �� ����� ���
                                sText = sText & sRow
                                Exit Do
                            End If
                        Loop
                    Else
700                         '����� ������� �������
                        i = InStr(1, sRow, Chr$(34))
                        If i Then
                            StartB = StartB + i '�� �� �� �� ���� ������� �������
                            
                            sText = Chr$(34)
                            sRow = Mid$(sRow, i + 1)
                            flFindSekQuote = True
                            GoTo ReadText '������ �� ����
                        End If
                    End If
                Else
                    '����� �� ����� ������ � �������� �������
800                 '����� ���� "="
                    i = InStr(1, sRow, "=")
                    If i > 1 Then
                        sBuf = Trim$(Mid$(sRow, 1, i - 1))
                        a = Val(sBuf)
                        If a Then
                            iTxtInd = a: flRedText = True
                            '������� ���� "="-��
                            StartB = bRead + i + 1
                            sRow = Mid$(sRow, i + 1)
                            GoTo ReadText '������ �� ����
                        End If
                    End If
                End If
            End Select
1100        bRead = bRead + b
              
        Loop
    Close #frFile
    frFile = 0

    ReadRegFile = True

Exit Function
ErH:
If frFile Then Close #frFile
mdlErrLog.add_to_errLog mdlID, mdlName, fN, Erl, Err
Err.Clear
End Function






Private Function AddText(ByVal modN As Integer, _
                         ByVal IndTxt As Integer, _
                         ByVal lTextLen As Long, _
                         ByVal lStartByte As Long, _
                         ByVal lCharSet As Long) As Boolean
Const fN = 3001
On Error GoTo ErH
Dim i As Integer
Dim a As Integer

10
'������� �� ��������� �� ��������� ����� � ������ �� ����� �����
    
    
    
'===========================================
'������
'����� ����� 5432 �� ����� 12
'����� ����� � � = inModule(12)
'� � ������� �� ������ arModule()
'i = arModule(�).arInd(5432)
'���� i ���� �������� �� ������ ��� ���������
'sString = arModule(�).sStr(i)
'===========================================
        
        If lCharSet = 0 Then GoTo ErH
        
        
        a = inModule(modN)
        If a = 0 Then
            '�� � ������� ���� �����
            max_arModule = max_arModule + 1
100         If max_arModule = 1 Then
                ReDim arModule(10) As sMdlStrInd
            Else
                If max_arModule > UBound(arModule) Then ReDim Preserve arModule(max_arModule + 10) As sMdlStrInd
            End If
            '�������� ������ �� ����� �� �����
            inModule(modN) = max_arModule
500         ReDim arModule(max_arModule).arInd(ub)
            a = max_arModule
        End If
        With arModule(a)
            i = .arInd(IndTxt)
            If i = 0 Then
                '���� ��������� �� �������
                .maxID = .maxID + 1
                If .maxID = 1 Then
                    ReDim .sStr(.maxID + 10)
                Else
                    If .maxID > UBound(.sStr) Then ReDim Preserve .sStr(.maxID + 10)
                End If
800             i = .maxID
            End If
            With .sStr(i)
                .lStart = lStartByte
                .lLong = lTextLen
                .CharSet = lCharSet
            End With
            
            .arInd(IndTxt) = i
        End With
900
        AddText = True

Exit Function
ErH:
mdlErrLog.add_to_errLog mdlID, mdlName, fN, Erl, Err
Err.Clear
End Function




Public Function GetsText(ByVal ModulN As Integer, _
                         ByVal TextIndex As Integer, _
                         ByRef lCharSet As Long) As String
Const fN = 3003
On Error GoTo ErH
Dim sPath As String
Dim sBuf  As String
Dim a As Long
Dim i As Long
Dim x As Long
10
'������� �� ������� �� ����� �� ������

'===========================================
'������
'����� ����� 5432 �� ����� 12
'����� ����� � � = inModule(12)
'� � ������� �� ������ arModule()
'i = arModule(�).arInd(5432)
'���� i ���� �������� �� ������ ��� ���������
'sString = arModule(�).sStr(i)
'===========================================
    sBuf = "N/a - " & TextIndex
    
    a = inModule(ModulN)
    If a Then
        i = arModule(a).arInd(TextIndex)
        If i Then
            x = arModule(a).sStr(i).lLong
            lCharSet = arModule(a).sStr(i).CharSet
            sBuf = Space$(x)
            x = arModule(a).sStr(i).lStart
            sPath = mdl_Const.appPath
            sPath = AddDirSep(sPath) & mdl_Const.res_file_Name
            
            If Not clsFS.OpenFile(sPath, 1) Then GoTo ErH
            If Not clsFS.GetData(x, 0, sBuf) Then GoTo ErH
            
        End If
    End If
    
    GetsText = Trim$(sBuf)

Exit Function
ErH:
mdlErrLog.add_to_errLog mdlID, mdlName, fN, Erl, Err
Err.Clear
End Function










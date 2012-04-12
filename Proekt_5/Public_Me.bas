Attribute VB_Name = "Public_Me"
Option Explicit

Private Const ModulIdString   As String = "mdlPublicMe - "

'nFunction = 3000 -> Get_Init_Data
'nFunction = 3001 -> Set_Init_Data
'nFunction = 3002 ->
'nFunction = 3003 ->


Public MePath           As String 'Пътека на ехе-то
Public sHistFilePath    As String 'Пътека на файла с данни
Public sPicFilePath     As String 'Пътека на картинката за фон
Public sMarckt_Message  As String 'текст за послание
Public lStartRec        As Long   'Стартов запис за четене
Public lEndRec          As Long   'Последен запис за четене
Public sOutFilePath     As String 'Пътека на изходящ файл
Public sOutFileName     As String 'Име за изходящ файл
Public Const sINI_Name  As String = "E_Mail.ini"

Public Enum sRedData
    RD_HistFilePath_1 = 1   'Пътека до файла с данни 128 символа
    RD_HistFilePath_2 = 2
    
    RD_Marckt_Message_1 = 3 'съобщението е 128 символа
    RD_Marckt_Message_2 = 4 'Затова е в четири ключа
    
    RD_NameFon_1 = 5        'Файл за картинката за фон 128 символа
    RD_NameFon_2 = 6
    
    RD_OutPutFilePath_1 = 7   'Пътека за изходящ файл
    RD_OutPutFilePath_2 = 11  'Пътека за изходящ файл
    
    RD_OutPutFileName = 8   'Име за изходящ файл
    
    RD_StartRec = 9         'Стартов запис за четене
    RD_EndRec = 10          'Последен запис за четене
    
End Enum



Public Function Get_Init_Data() As Boolean
Const nFunction = 3000
On Error GoTo ErH
Dim sBuf As String
    'Функция за четене на основни за приложението параметри
    'Пътека за изх. файл

100 'текст за послание : Текст, шрифт, цвят и височина на шрифта
    'Marckt_Message 1-2
    'текст за послание
    Read_Reg_String MePath, RD_Marckt_Message_1, sBuf
    sMarckt_Message = sBuf
    Read_Reg_String MePath, RD_Marckt_Message_2, sBuf
    sMarckt_Message = sMarckt_Message & sBuf
    
150 ''Пътека до файла с данни 128 символа
    'текст за послание
    Read_Reg_String MePath, RD_HistFilePath_1, sBuf
    sHistFilePath = sBuf
    Read_Reg_String MePath, RD_HistFilePath_2, sBuf
    sHistFilePath = sHistFilePath & sBuf
    
200 'Име на картинката за фон
    Read_Reg_String MePath, RD_NameFon_1, sBuf
    sPicFilePath = sBuf
    Read_Reg_String MePath, RD_NameFon_2, sBuf
    sPicFilePath = sPicFilePath & sBuf
    
300 'Пътека за изходящ файл
    Read_Reg_String MePath, RD_OutPutFilePath_1, sBuf
    sOutFilePath = Trim$(sBuf)
    Read_Reg_String MePath, RD_OutPutFilePath_2, sBuf
    sOutFilePath = sOutFilePath & Trim$(sBuf)
    sOutFilePath = "C:\desde"
    If Not DirExists(sOutFilePath) Then sOutFilePath = vbNullString
    
    If Len(sOutFilePath) = 0 Then sOutFilePath = MePath

400 'Ime за изходящ файл
    Read_Reg_String MePath, RD_OutPutFileName, sBuf
    sOutFileName = Trim$(sBuf)
    If Len(sOutFileName) = 0 Then sOutFileName = "Pred_Out.txt"
    
500 'Стартов запис за четене
    Read_Reg_String MePath, RD_StartRec, sBuf
    lStartRec = Val(sBuf)
    If lStartRec <= 0 Then lStartRec = 1
    
600 'Последен запис за четене
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
    'Функция за четене на основни за приложението параметри
    'Пътека за изх. файл
    
    'Име на картинката за фон
    'текст за послание : Текст, шрифт, цвят и височина на шрифта
    'за трите етикета : Текст, шрифт, цвят и височина на шрифта
    
100
    Select Case a
    
    'Marckt_Message 1-2
    'текст за послание
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
200     'Име на картинката за фон
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
300     'Пътека на файла с данни
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
400     'Пътека на изх. файл
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
500     'Име на изх файл
        Write_Reg_String MePath, RD_OutPutFileName, sOutFileName, "Output File Name "
    
    End Select
    
    Set_Init_Data = True
    
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear

End Function

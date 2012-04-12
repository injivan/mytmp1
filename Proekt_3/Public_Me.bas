Attribute VB_Name = "Public_Me"
Option Explicit

Private Const ModulIdString   As String = "mdlPublicMe - "

'nFunction = 3000 -> Get_Init_Data
'nFunction = 3001 -> Set_Init_Data
'nFunction = 3002 -> Get_Font
'nFunction = 3003 -> Set_Font


Public MePath       As String 'Пътека на ехе-то
Public IzhFilePath  As String 'Пътека на Изходящия файл
Public IzhFileName  As String 'Пътека на Изходящия файл
Public PicFilePath  As String 'Пътека на картинката за фон
Public Marckt_Message As String 'текст за послание
Public Font1_Color    As Long
Public Font2_Color    As Long
Public Font1          As StdFont
Public Font2          As StdFont
Public sINI_Name      As String




Public Enum sRedData
    RD_OutFilePath_1 = 1 'Пътека за Изходящия файл 128 символа
    RD_OutFilePath_2 = 2
    
    RD_Marckt_Message_1 = 3 'съобщението е 256 символа
    RD_Marckt_Message_2 = 4 'Затова е в четири ключа
    RD_Marckt_Message_3 = 5
    RD_Marckt_Message_4 = 6
    
    RD_NameFon_1 = 7 'Файл за картинката за фон 128 символа
    RD_NameFon_2 = 8 '
    
    RD_OutFileName = 9 'Име на Изходящия файл 18 символа
    
    'Шрифт за Marckt_Message
    RD_Font1_Name = 10
    RD_Font1_Color = 11
    RD_Font1_Size = 12
    RD_Font1_Charset = 13
    RD_Font1_BIU = 14
    
    'Шрифт за drugite etiketi
    RD_Font2_Name = 15
    RD_Font2_Color = 16
    RD_Font2_Size = 17
    RD_Font2_Charset = 18
    RD_Font2_BIU = 19
End Enum



Public Function Get_Init_Data() As Boolean
Const nFunction = 3000
On Error GoTo ErH
Dim sBuf As String
    'Функция за четене на основни за приложението параметри
    'Пътека за изх. файл
    
    'Име на картинката за фон
    'текст за послание : Текст, шрифт, цвят и височина на шрифта
    'за трите етикета : Текст, шрифт, цвят и височина на шрифта
    
100 'Marckt_Message 1-4
    'текст за послание
    Read_Reg_String MePath, RD_Marckt_Message_1, sBuf
    Marckt_Message = sBuf
    Read_Reg_String MePath, RD_Marckt_Message_2, sBuf
    Marckt_Message = Marckt_Message & sBuf
    Read_Reg_String MePath, RD_Marckt_Message_3, sBuf
    Marckt_Message = Marckt_Message & sBuf
    Read_Reg_String MePath, RD_Marckt_Message_4, sBuf
    Marckt_Message = Marckt_Message & sBuf
    If Marckt_Message = vbNullString Then Marckt_Message = "Hear is my marketing message"
150 '!!!шрифт, цвят и височина на шрифта

    
200 'Име на картинката за фон
    Read_Reg_String MePath, RD_NameFon_1, sBuf
    PicFilePath = sBuf
    Read_Reg_String MePath, RD_NameFon_2, sBuf
    PicFilePath = PicFilePath & sBuf
    
300 'Пътека за изх. файл
    Read_Reg_String MePath, RD_OutFilePath_1, sBuf
    IzhFilePath = sBuf
    Read_Reg_String MePath, RD_OutFilePath_2, sBuf
    IzhFilePath = IzhFilePath & sBuf
    If IzhFilePath = vbNullString Then IzhFilePath = MePath
400 'Име на изх. файл
    Read_Reg_String MePath, RD_OutFileName, sBuf
    If sBuf = vbNullString Then sBuf = "Out.csv"
    IzhFileName = sBuf
    
500 'Цвят на шрифт 1
    Read_Reg_String MePath, RD_Font1_Color, sBuf
    Font1_Color = Val(sBuf)
    Read_Reg_String MePath, RD_Font2_Color, sBuf
    Font2_Color = Val(sBuf)
    
    
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
    
    'Marckt_Message 1-4
    'текст за послание
    Case RD_Marckt_Message_1 To RD_Marckt_Message_4
        i = 64: z = 1
        sTmp = Marckt_Message
        For j = 1 To 4
            If Len(sTmp) < i Then i = Len(sTmp)
            If i Then
                sBuf = Mid$(sTmp, 1, i)
                z = z + i
                sTmp = Mid$(sTmp, z)
            End If
            If j = 1 Then k = RD_Marckt_Message_1
            If j = 2 Then k = RD_Marckt_Message_2
            If j = 3 Then k = RD_Marckt_Message_3
            If j = 4 Then k = RD_Marckt_Message_4
            
            Write_Reg_String MePath, k, sBuf, "Marckt_Message - " & j
            sBuf = vbNullString
           
        Next j
    
    
    Case RD_NameFon_1 To RD_NameFon_2
    
200     'Име на картинката за фон
        i = 64: z = 1
        sTmp = PicFilePath
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
       
    Case RD_OutFilePath_1 To RD_OutFilePath_2
300     'Пътека за изх. файл
        i = 64: z = 1
        sTmp = IzhFilePath
        For j = 1 To 2
            If Len(sTmp) < i Then i = Len(sTmp)
            If i Then
                sBuf = Mid$(sTmp, z, i)
                z = z + i
                sTmp = Mid$(sTmp, z)
            End If
            
            If j = 1 Then k = RD_OutFilePath_1
            If j = 2 Then k = RD_OutFilePath_2
            
            Write_Reg_String MePath, k, sBuf, "OutFilePath - " & j
            sBuf = vbNullString
        Next j
    Case RD_OutFileName
400     'Име на изх. файл
        Write_Reg_String MePath, RD_OutFileName, IzhFileName, "IzhFileName"
    Case RD_Font1_Color
500     'Цвят на щрифта
        Write_Reg_String MePath, RD_Font1_Color, CStr(Font1_Color)
    Case RD_Font2_Color
        Write_Reg_String MePath, RD_Font2_Color, CStr(Font2_Color)
    End Select
    
    Set_Init_Data = True
    
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear

End Function


Public Function Get_Font(ByRef nom_1_2 As Long, ByRef mFont As StdFont) As Boolean
Const nFunction = 3002
On Error GoTo ErH
Dim sName   As String
 
Dim lSize   As Long
Dim lCharset   As Long
Dim sBIU    As String 'Болд Италик Ъндерлайн
Dim sBuf    As String
Dim z       As Long
10

    'Функция за вземане на еден щрифт
    'Наме Колор Сизе
    Set mFont = New StdFont
    If nom_1_2 = 1 Then
100     Read_Reg_String MePath, RD_Font1_Name, sBuf
        sName = sBuf
        
200     Read_Reg_String MePath, RD_Font1_Size, sBuf
        lSize = Val(sBuf)
        Read_Reg_String MePath, RD_Font1_Charset, sBuf
        lCharset = Val(sBuf)
400     Read_Reg_String MePath, RD_Font1_BIU, sBuf
        sBIU = Val(sBuf)
    ElseIf nom_1_2 = 2 Then
800     Read_Reg_String MePath, RD_Font2_Name, sBuf
        sName = sBuf
        
900     Read_Reg_String MePath, RD_Font2_Size, sBuf
        lSize = Val(sBuf)
        Read_Reg_String MePath, RD_Font2_Charset, sBuf
        lCharset = Val(sBuf)
1000    Read_Reg_String MePath, RD_Font2_BIU, sBuf
        sBIU = Val(sBuf)
    End If
    
    With mFont
1200    If Len(sName) Then .Name = sName: z = 1
1230    If lCharset Then .Charset = lCharset: z = 1
1250    If lSize Then .Size = lSize: z = 1
1300
        If Len(sBIU) Then
1400        If Val(Mid$(sBIU, 1, 1)) Then mFont.Bold = True: z = 1
1450        If Val(Mid$(sBIU, 2, 1)) Then mFont.Italic = True: z = 1
1500        If Val(Mid$(sBIU, 3, 1)) Then mFont.Underline = True: z = 1
        End If
    End With
    If z = 0 Then Set mFont = Nothing
    Get_Font = True
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear

End Function


Public Function Set_Font(ByRef nom_1_2 As Long, ByRef mFont As StdFont) As Boolean
Const nFunction = 3003
On Error GoTo ErH
Dim sName   As String
 
Dim lSize   As Long
Dim lCharset   As Long
Dim sBIU    As String 'Болд Италик Ъндерлайн
Dim sBuf    As String
10

    'Функция за zapis на еден щрифт
    
    With mFont
        sName = .Name
        lCharset = .Charset
        lSize = .Size
         
        sBIU = "000"
        If mFont.Bold Then Mid$(sBIU, 1, 1) = "1"
        If mFont.Italic Then Mid$(sBIU, 2, 1) = "1"
        If mFont.Underline Then Mid$(sBIU, 3, 1) = "1"
        If Val(sBIU) = 0 Then sBIU = vbNullString
    End With
    
    'Наме Колор Сизе
    If nom_1_2 = 1 Then
100     Write_Reg_String MePath, RD_Font1_Name, sName
        
200     Write_Reg_String MePath, RD_Font1_Size, CStr(lSize)
        Write_Reg_String MePath, RD_Font1_Charset, CStr(lCharset)
400     Write_Reg_String MePath, RD_Font1_BIU, sBIU
    ElseIf nom_1_2 = 2 Then
500     Write_Reg_String MePath, RD_Font2_Name, sName
        
600     Write_Reg_String MePath, RD_Font2_Size, CStr(lSize)
        Write_Reg_String MePath, RD_Font2_Charset, CStr(lCharset)
700     Write_Reg_String MePath, RD_Font2_BIU, sBIU
    End If
    
    

    Set_Font = True
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear

End Function




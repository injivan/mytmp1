VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmInet1 
   Caption         =   "Form3"
   ClientHeight    =   7380
   ClientLeft      =   3720
   ClientTop       =   1875
   ClientWidth     =   7770
   LinkTopic       =   "Form3"
   ScaleHeight     =   7380
   ScaleWidth      =   7770
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   4455
      Left            =   240
      ScaleHeight     =   4395
      ScaleWidth      =   5835
      TabIndex        =   5
      Top             =   1200
      Width           =   5895
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   240
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   3855
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
   End
End
Attribute VB_Name = "frmInet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DoProces As Boolean

Dim arD(3, 2, 1) As Double ' Първа дименсия 0 - БНБ; 1 - Пощенска; 2 - Райфаизен; 3 - ФИБ
                           ' Втора дименсия 0 - USD; 1 - GBP; 2 - CHF
                           ' Трета дименсия 0 - Купува ; 1 - Продава

Dim asBest(2, 1) As Long   ' Първа дименсия 0 - USD; 1 - GBP; 2 - CHF
                           ' Втора дименсия 0 - Купува ; 1 - Продава
                           ' стойността е 1,2 или 3 - Номр на банка
                          



Private Sub cmdStop_Click()
    DoProces = False
End Sub

Private Sub Command1_Click()
Dim s As String
    Text1.Text = vbNullString
    'отварям файл за запис на данните
    
    'Вземам данни от БНБ за Фиксинга на 3-те валути
    s = "http://www.bnb.bg/AboutUs/AURSS/index.htm?getRSS=1"
    s = Inet1.OpenURL(s)
    If Not prs_BNB(s) Then GoTo ErH

    'Вземам данни от postbank
    s = "http://www.postbank.bg/"
    s = Inet1.OpenURL(s)
    If Not prs_PostB(s) Then GoTo ErH

    'Райфаизен
    s = "http://rbb.bg/bg-BG/Corporate_Customers/Investing/07_risk/04_rates/"
    s = Inet1.OpenURL(s)
    If Not prs_Raif(s) Then GoTo ErH
    
    'ПИБ
    s = "http://www.fibank.bg/"
    s = Inet1.OpenURL(s)
    If Not prs_PIB(s) Then GoTo ErH
    
    
    Print_arr
    
ErH:

End Sub

 
Private Sub Command2_Click()
Dim s As String
Dim s1 As String
Dim h As String
Dim i As Long
Dim z As Long

On Error Resume Next
Command2.Visible = False
cmdStop.Visible = True
DoProces = True


Open "a.in" For Input As #1
Text1.Text = vbNullString
Do While Not EOF(1)
    Line Input #1, s
    If Not DoProces Then Exit Do
    
    s1 = vbNullString
    s1 = Inet1.OpenURL(s)
    
    s1 = LCase$(s1)
    If InStr(1, s1, "htc_found") Then
        h = h & vbCrLf & s
        Text1.Text = h
    End If
    
    i = i + 1
    Text2.Text = i
    DoEvents
Loop

Close #1
Text2.Text = "Gotowo"
Text1.Text = h

Open "b.out" For Output As #1
Print #1, h
Close #1

cmdStop.Visible = False
Command2.Visible = True
DoProces = False
End Sub

 
Private Sub Form_Click()

    'With Text1
    '    .ForeColor = QBColor(Rnd * 15)
'
'        .Text = "Ivan"
'        .ForeColor = QBColor(Rnd * 15)
'
'        .Text = .Text & " е пич"
'    End With
End Sub

Private Sub Form_Load()

'http://www.bnb.bg/Statistics/StExternalSector/StExchangeRates/StERForeignCurrencies/index.htm?download=xml&search=&lang=BG
'http://www.ecb.int/stats/eurofxref/eurofxref-daily.xml
'http://rbb.bg/bg-BG/Corporate_Customers/Investing/07_risk/04_rates/

'Text2.Text = "http://mysolutions.site90.net/s1/index.php"
'бнб фиксинг
'Text2.Text = "http://www.bnb.bg/AboutUs/AURSS/index.htm?getRSS=1"



With Picture1.Font
    .Name = "Fixedsys"
    .Size = 12
End With

End Sub

Private Function prs_BNB(ByRef strIn As String) As Boolean
'<?xml version="1.0" encoding="UTF-8" ?>
'<rss version="2.0" xmlns:dc="http://purl.org/dc/elements/1.1/"
'        xmlns:sy="http://purl.org/rss/1.0/modules/syndication/">
'    <channel>
'    <title>Обменни курсове</title>
'    <pubDate>Mon, 06 Dec 2010 00:00:00 +0200</pubDate>
'    <description>БНБ Текущи обменни курсове</description>
'    <dc:language>bg-bg</dc:language>
'    <sy:updatePeriod>daily</sy:updatePeriod>
'    <sy:updateFrequency>1</sy:updateFrequency>
'    <link>http://www.bnb.bg/Statistics/StExternalSector/StExchangeRates/StERForeignCurrencies/index.htm</link>
'
'    <item>
'         <title>Курсове за 06.12.2010</title>
'         <pubDate>Mon, 06 Dec 2010 00:00:00 +0200</pubDate>
'         <link>http://www.bnb.bg/Statistics/StExternalSector/StExchangeRates/StERForeignCurrencies/index.htm</link>
'    <description>
'            <![CDATA[<h3>
'
'    Българска народна банка</h3>
'    <ul>
'        <li>
'            <img src="http://www.bnb.bg/bnbweb/fragments/bnb_iclude_fragment/images/currency/usd.gif" alt="usd" />
'             <em>1 USD = </em> <strong>1.47276 </strong> BGN (down)
'       </li>
'        <li>
'            <img src="http://www.bnb.bg/bnbweb/fragments/bnb_iclude_fragment/images/currency/gbp.gif" alt="gbp" />
'             <em>1 GBP = </em> <strong>2.30858 </strong> BGN (up)
'       </li>
'        <li>
'            <img src="http://www.bnb.bg/bnbweb/fragments/bnb_iclude_fragment/images/currency/chf.gif" alt="chf" />
'             <em>1 CHF = </em> <strong>1.49483 </strong> BGN (up)
'       </li>
'     </ul>]]>
'        </description>
'        </item>
'</channel>
'</rss>

Const em     As String = "<em>"
Const strong As String = "<strong>"

 
Dim sO      As String


Dim aStr    As Long
Dim aLng    As Long
Dim a       As Long

Dim lStr    As Long

     
    lStr = 1
    Do
        a = InStr(lStr, strIn, em)
        If a Then
            strIn = Mid$(strIn, a)
            a = InStr(1, strIn, strong)
            
            aLng = InStr(a + 1, strIn, "<")
            aLng = aLng - a - Len(strong)
            sO = Trim$(Mid$(strIn, a + Len(strong), aLng))
            lStr = a + aLng
            
            'Намерил съм нещо
            Select Case Mid$(strIn, 7, 3)
            Case "USD": arD(0, 0, 0) = Val(sO): arD(0, 0, 1) = Val(sO)
            Case "GBP": arD(0, 1, 0) = Val(sO): arD(0, 1, 1) = Val(sO)
            Case "CHF": arD(0, 2, 0) = Val(sO): arD(0, 2, 1) = Val(sO)
            End Select
            
        End If
    Loop While a
    prs_BNB = True
    
    
End Function
Private Function prs_Raif(ByRef strIn As String) As Boolean
Const tbl       As String = "<table width=""100%"">"
Const End_tbl   As String = "</table>"
    
Dim sTabl       As String
Dim sBuf        As String
Dim a           As Long
Dim b           As Long
Dim c           As Long
Dim d           As Long
Dim s()         As String
    'lStr = 1
    Do
        a = InStr(1, strIn, tbl)
        If a Then
            strIn = Mid$(strIn, a)
            'Сега намитам края на таблицата
            b = InStr(2, strIn, End_tbl)
            a = b + Len(End_tbl)
            'взимам таблицата
            sTabl = Mid$(strIn, 1, a)
            strIn = Mid$(strIn, a + 1)
            'почиствам я от всички тагове и остават само код на валута и курс
            b = 0
            For a = 1 To Len(sTabl)
                c = Asc(Mid$(sTabl, a, 1))
                Select Case c
                Case 60: b = 0: d = 0 '"<"
                Case 62: b = 1: d = 1 '">"
                Case Else
                    If b Then
                        If c > 32 And c <= 122 Then
                            If d Then sBuf = sBuf & ";": d = 0
                            sBuf = sBuf & Mid$(sTabl, a, 1)
                        End If
                        
                    End If
                End Select
            Next a
        End If
    Loop While a
'  '0   1 2 3      4      5      6      7     8
'    ;CHF;1;-;1.4740;1.5280;1.4700;1.5320;2010-12-1008:20
'     '9  10           13    14       15     16     17
'    ;EUR;1;-;1.9500;1.9590;1.9480;1.9600;2010-12-1008:20
'  '18 19
'    ;GBP;1;-;2.2820;2.3720;2.2760;2.3780;2010-12-1008:20
'    ;USD;1;-;1.4530;1.4950;1.4520;1.4960;2010-12-1008:20
'
    s = Split(sBuf, ";")
    For a = 0 To 3
        Select Case s(a * 8 + 1)
        Case "USD": arD(2, 0, 0) = Val(s(a * 8 + 6)): arD(2, 0, 1) = Val(s(a * 8 + 7))
        Case "GBP": arD(2, 1, 0) = Val(s(a * 8 + 6)): arD(2, 1, 1) = Val(s(a * 8 + 7))
        Case "CHF": arD(2, 2, 0) = Val(s(a * 8 + 6)): arD(2, 2, 1) = Val(s(a * 8 + 7))
        End Select
    Next a
    
    
    prs_Raif = True
End Function

Private Function prs_PIB(ByRef strIn As String) As Boolean
Const tbl       As String = "<table class=""CrrDetails"""
Const End_tbl   As String = "</table>"

Dim sTabl       As String
Dim sBuf        As String
Dim a           As Long
Dim b           As Long
Dim c           As Long
Dim d           As Long
Dim s()         As String
Dim lStr As Long

    lStr = 1
    Do
        a = InStr(lStr, strIn, tbl)
        If a Then
            strIn = Mid$(strIn, a)
            'Сега намитам края на таблицата
            b = InStr(2, strIn, End_tbl)
            a = b + Len(End_tbl)
            'взимам таблицата
            sTabl = Mid$(strIn, 1, a)
            strIn = Mid$(strIn, a + 1)
            'почиствам я от всички тагове и остават само код на валута и курс
            b = 0: sBuf = sBuf & ";"
            For a = 1 To Len(sTabl)
                c = Asc(Mid$(sTabl, a, 1))
                Select Case c
                Case 60: b = 0: d = 0 '"<"
                Case 62: b = 1: d = 1 '">"
                Case Else
                    If b Then
                        If c > 32 And c <= 122 Then
                            If d Then sBuf = sBuf & ";": d = 0
                            sBuf = sBuf & Mid$(sTabl, a, 1)
                        End If
                        
                    End If
                End Select
            Next a
        End If
    Loop While a
'  '0 1 2  3    4  5    6       7       8
'    ; ; ;&nbsp; ;GBP;2.336300;2.24850;2.42070
'  '    9  10   11 12   13       14      15
'      ; ;&nbsp; ;EUR;1.955830;1.94600;1.95900
'  '    16 17   18 19   20        21     22
'      ; ;&nbsp; ;CHF;1.504720;1.44900;1.55920
'      ; ;&nbsp;;USD;1.476770;1.42870;1.52820

    
    
    s = Split(sBuf, ";")
    For a = 0 To 3
        Select Case s(a * 7 + 5)
        Case "USD": arD(3, 0, 0) = Val(s(a * 7 + 7)): arD(3, 0, 1) = Val(s(a * 7 + 8))
        Case "GBP": arD(3, 1, 0) = Val(s(a * 7 + 7)): arD(3, 1, 1) = Val(s(a * 7 + 8))
        Case "CHF": arD(3, 2, 0) = Val(s(a * 7 + 7)): arD(3, 2, 1) = Val(s(a * 7 + 8))
        End Select
    Next a
    
    
    prs_PIB = True
    
    
    
    
End Function
Private Function prs_PostB(ByRef strIn As String) As Boolean
Const tbl       As String = "<table class=""textsml"""
Const End_tbl   As String = "</table>"

Dim sTabl       As String
Dim sBuf        As String
Dim a           As Long
Dim b           As Long
Dim c           As Long

Dim lStr As Long

    lStr = 1
    Do
        a = InStr(lStr, strIn, tbl)
        If a Then
            strIn = Mid$(strIn, a)
            'Сега намитам края на таблицата
            b = InStr(2, strIn, End_tbl)
            a = b + Len(End_tbl)
            'взимам таблицата
            sTabl = Mid$(strIn, 1, a)
            strIn = Mid$(strIn, a + 1)
            'почиствам я от всички тагове и остават само код на валута и курс
            b = 0: sBuf = sBuf & ";"
            For a = 1 To Len(sTabl)
                c = Asc(Mid$(sTabl, a, 1))
                Select Case c
                Case 60: b = 0 '"<"
                Case 62: b = 1 '">"
                Case Else
                    If b Then
                        If c >= 32 And c <= 122 Then sBuf = sBuf & Mid$(sTabl, a, 1)
                    End If
                End Select
            Next a
        End If
    Loop While a
    
    
    strIn = sBuf
    'Сплитвам на ";"
    Do
        a = InStr(1, strIn, ";")
        If a Then
            sTabl = Trim$(Mid$(strIn, 1, a - 1))
            strIn = Mid$(strIn, a + 1)
            
            sBuf = Trim$(Mid$(sTabl, 4))
            b = InStr(1, sBuf, Chr$(32))
            
            'Намерил съм нещо
            Select Case Left$(sTabl, 3)
            Case "USD": arD(1, 0, 0) = Val(Trim$(Mid$(sBuf, 1, b))): arD(1, 0, 1) = Val(Trim$(Mid$(sBuf, b)))
            Case "GBP": arD(1, 1, 0) = Val(Trim$(Mid$(sBuf, 1, b))): arD(1, 1, 1) = Val(Trim$(Mid$(sBuf, b)))
            Case "CHF": arD(1, 2, 0) = Val(Trim$(Mid$(sBuf, 1, b))): arD(1, 2, 1) = Val(Trim$(Mid$(sBuf, b)))
            End Select
            
            
        End If
        
    Loop While a
    'Трябва да имам едим дълъг стринг
    
    prs_PostB = True
    
    
    
    
End Function
'''
'''
'''
'''<table class="textsml" width="100%" align="center" cellspacing="0" cellpadding="0" border="0">
'''  <tr>
'''    <td bgcolor="#EDB09E" colspan="3" height="16"> Валути - <b>07.12.2010</b></td>
'''  </tr>
'''  <tr>
'''    <td width="20%"> </td>
'''
'''    <td width="40%" align="right">Купува</td>
'''    <td width="40%" align="right">Продава</td>
'''  </tr>
'''</table>
'''<table class="textsml" width="100%" align="center" cellspacing="0" cellpadding="0" border="0">
'''  <tr class="td-white">
'''    <td nowrap="nowrap" width="20%" align="left" style="FONT-SIZE: 9px;">
'''      <b>USD</b>
'''
'''    </td>
'''    <td nowrap="nowrap" width="40%" align="right" style="FONT-SIZE: 9px;">1.4435</td>
'''    <td nowrap="nowrap" width="40%" align="right" style="FONT-SIZE: 9px;">1.4875</td>
'''  </tr>
'''</table>
'''<table class="textsml" width="100%" align="center" cellspacing="0" cellpadding="0" border="0">
'''  <tr class="td-white">
'''    <td nowrap="nowrap" width="20%" align="left" style="FONT-SIZE: 9px;">
'''      <b>EUR</b>
'''
'''    </td>
'''    <td nowrap="nowrap" width="40%" align="right" style="FONT-SIZE: 9px;">1.9505</td>
'''    <td nowrap="nowrap" width="40%" align="right" style="FONT-SIZE: 9px;">1.9595</td>
'''  </tr>
'''</table>
'''<marquee direction="up" height="10" loop="" scrollamount="1" scrolldelay="150" id="ieslider" style="BORDER-RIGHT: white 1px solid; BORDER-TOP: white 1px solid; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid; BACKGROUND-COLOR: white;">
'''  <table class="textsml" width="100%" align="center" cellspacing="0" cellpadding="0" border="0">
'''    <tr class="td-white">
'''      <td nowrap="nowrap" width="20%" align="left" style="FONT-SIZE: 9px;">
'''
'''        <b>CHF</b>
'''      </td>
'''      <td nowrap="nowrap" width="40%" align="right" style="FONT-SIZE: 9px;">1.4901</td>
'''      <td nowrap="nowrap" width="40%" align="right" style="FONT-SIZE: 9px;">1.5211</td>
'''    </tr>
'''  </table>
'''  <table class="textsml" width="100%" align="center" cellspacing="0" cellpadding="0" border="0">
'''
'''    <tr class="td-white">
'''      <td nowrap="nowrap" width="20%" align="left" style="FONT-SIZE: 9px;">
'''        <b>GBP</b>
'''      </td>
'''      <td nowrap="nowrap" width="40%" align="right" style="FONT-SIZE: 9px;">2.2742</td>
'''      <td nowrap="nowrap" width="40%" align="right" style="FONT-SIZE: 9px;">2.3435</td>
'''    </tr>
'''
'''  </table>
'''  <table class="textsml" width="100%" align="center" cellspacing="0" cellpadding="0" border="0">
'''    <tr class="td-white">
'''      <td nowrap="nowrap" width="20%" align="left" style="FONT-SIZE: 9px;">
'''        <b>SEK</b>
'''      </td>
'''      <td nowrap="nowrap" width="40%" align="right" style="FONT-SIZE: 9px;">2.1168</td>
'''      <td nowrap="nowrap" width="40%" align="right" style="FONT-SIZE: 9px;">2.1813</td>
'''
'''    </tr>
'''  </table>
'''</marquee>
'''







Private Sub Print_arr()
'Dim arD(3, 2, 1) As Double ' Първа дименсия 0 - БНБ; 1 - Пощенска; 2 - Райфаизен; 3 - ФИБ
                            ' Втора дименсия 0 - USD; 1 - GBP; 2 - CHF
                            ' Трета дименсия 0 - Купува ; 1 - Продава



Dim str1 As String
Dim str2 As String

Dim i    As Long
Dim j    As Long


'формат USD
         
' 1 - 10 - Banka
'11 - 10 - kurs Kupuwa
'21 - 10 - Prowada

        str1 = Space$(30)
        'Заглавие
        Mid$(str1, 11, 10) = "  КУПУВА  "
        Mid$(str1, 21, 10) = " ПРОДАВА  "
        str2 = str1 & vbCrLf
        
        For i = 0 To 2   'По валутите
            LSet str1 = vbNullString
            Select Case i
            Case 0: Mid$(str1, 1, 10) = "USD       "
            Case 1: Mid$(str1, 1, 10) = "GBP       "
            Case 2: Mid$(str1, 1, 10) = "CHF       "
            End Select
            str2 = str2 & str1 & vbCrLf
            For j = 0 To 3 ' по банките
                LSet str1 = vbNullString
                
                Select Case j
                Case 0: Mid$(str1, 1, 10) = "     БНБ  "
                Case 1: Mid$(str1, 1, 10) = "     ПОЩ  "
                Case 2: Mid$(str1, 1, 10) = "     РАЙФ "
                Case 3: Mid$(str1, 1, 10) = "     ПИБ  "
                End Select
                
                Mid$(str1, 11, 10) = arD(j, i, 0)   '"  КУПУВА  "
                Mid$(str1, 21, 10) = arD(j, i, 1)   '" ПРОДАВА  "
                
                str2 = str2 & str1 & vbCrLf
            Next j
        Next i
        
        
        Text1.Text = str2
        Picture1.Cls
        Picture1.Print str2
        
        Picture1.ForeColor = ColorConstants.vbGreen
        Picture1.Print " dadaa"
        Picture1.ForeColor = ColorConstants.vbBlack
        
End Sub


Private Sub FindBest()
    dim nBank0 as long
    dim nBank1 as long 
    dim dSum0 as double
    dim dSum1 as double
    
    Търсене
    For i = 0 To 2
        '"USD": 0 "GBP": 1 "CHF": 2 - на 2-ра дименсия
        'arD(1, i, 0)
        dsum0=0
        for j = 1 to 3
            'Банка
            if arD(j, i, 0) asBest(2, 1)
        
        
        next j
        arD(1, 0, 0) = Val(Trim$(Mid$(sBuf, 1, b))): arD(1, 0, 1) = Val(Trim$(Mid$(sBuf, b)))
        
        arD(1, 1, 0) = Val(Trim$(Mid$(sBuf, 1, b))): arD(1, 1, 1) = Val(Trim$(Mid$(sBuf, b)))
        
        arD(1, 2, 0) = Val(Trim$(Mid$(sBuf, 1, b))): arD(1, 2, 1) = Val(Trim$(Mid$(sBuf, b)))
        
    Next i
    asBest
End Sub


 
Private Sub Inet1_StateChanged(ByVal State As Integer)
Dim s As String
    Select Case State
        Case 0: s = "No state to report"
        Case 1: s = "The control is looking up the IP address of the specified host computer."
        Case 2: s = "The control successfully found the IP address of the specified host computer."
        Case 3: s = "The control is connecting to the host computer."
        Case 4: s = "The control successfully connected to the host computer."
        Case 5: s = "The control is sending a request to the host computer."
        Case 6: s = "The control successfully sent the request."
        Case 7: s = "The control is receiving a response from the host computer."
        Case 8: s = "The control successfully received a response from the host computer."
        Case 9: s = "The control is disconnecting from the host computer."
        Case 10: s = "The control successfully disconnected from the host computer."
        Case 11: s = "An error occurred in communicating with the host computer."
        Case 12: s = "The request has completed and all data has been received."
    End Select
    If Len(Text1.Text) > 1024 Then Text1.Text = vbNullString
    Text1.Text = Text1.Text & vbCrLf & s
    
    If Picture1.CurrentY > Picture1.Width Then Picture1.Cls
    Picture1.Print s
    
    
End Sub




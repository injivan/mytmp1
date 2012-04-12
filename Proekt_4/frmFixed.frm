VERSION 5.00
Begin VB.Form frmFixed 
   Caption         =   "Get Fixed numbers"
   ClientHeight    =   6465
   ClientLeft      =   1650
   ClientTop       =   1755
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   10680
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   360
      ScaleHeight     =   5475
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   360
      Width           =   9495
      Begin VB.CommandButton cmbExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   7200
         TabIndex        =   2
         Top             =   4920
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   7200
         TabIndex        =   1
         Top             =   4080
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmFixed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ModulIdString   As String = "frmFixed - "

 
Private Type CtrlProportions
    HeightProportions   As Single
    WidthProportions    As Single
    TopProportions      As Single
    LeftProportions     As Single
    FontSize            As Single
End Type
Private ProportionsArray()  As CtrlProportions



 

Private Sub Form_Initialize()
    MePath = App.Path: MePath = AddDirSep(MePath)
    InitResizeArray
End Sub
Sub InitResizeArray()

    Dim i As Integer
    
    On Error Resume Next
    
    ReDim ProportionsArray(0 To Controls.Count - 1)
    
    For i = 0 To Controls.Count - 1
        Select Case Controls(i).Name
        Case "a" '"Picture1", "Image1"
        Case Else
        With ProportionsArray(i)
            .HeightProportions = Controls(i).Height / ScaleHeight
            .WidthProportions = Controls(i).Width / ScaleWidth
            .TopProportions = Controls(i).TOp / ScaleHeight
            .LeftProportions = Controls(i).Left / ScaleWidth
            .FontSize = Controls(i).Font.Size / ScaleHeight
        End With
        End Select
    Next i
    
End Sub
Sub ResizeControls()

    On Error Resume Next
    
    Dim i As Integer
    
    For i = 0 To Controls.Count - 1
        With ProportionsArray(i)
            ' move and resize controls
            Controls(i).Move .LeftProportions * ScaleWidth, _
            .TopProportions * ScaleHeight, _
            .WidthProportions * ScaleWidth, _
            .HeightProportions * ScaleHeight
            
            Controls(i).Font.Size = .FontSize * ScaleHeight
            
        End With
    Next i
    
End Sub
Private Sub Form_Resize()
   
    ResizeControls
     
End Sub
Private Sub cmbExit_Click()
  Unload Me
End Sub
Private Function Obr() As Boolean
Const nFunction = 2005
On Error GoTo ErH
Dim cf              As cFile
Dim arr_FileIndex() As Long  'Начални позиции на всеки ред от файла
Dim tmp_Row_Count   As Long  'Брой редове в импортния файл
Dim sBuf            As String
Dim a()             As Long
Dim sA()            As String
Dim i               As Long
Dim j               As Long
Dim z               As Long
Dim x               As Long
Dim Y               As Long
Dim rez()           As t_RL
Dim flIn            As Boolean
Dim arRecIn()       As Long  'Масив с номерата на влизащите в критерия записи
Dim arRecIn_Count   As Long

Dim N               As Long 'брой възможности


10
     

    'отварям файла
    If Not FileExists(sHistFilePath) Then Err.Raise 1000, , "History File is not Exist"
    Set cf = New cFile
    If Not cf.OpenFile(sHistFilePath, 1) Then GoTo ErH
100
    'Функция за запис на началните позиции на редовете от файла във масив
    If Not fill_arrFileInd(arr_FileIndex, cf, tmp_Row_Count) Then GoTo ErH
300 'Масивите за броене на срещаните елементи
    ReDim a(5, 39) As Long
    
    ReDim arRecIn(5, 10) As Long
    arRecIn_Count = -1
    For i = 0 To tmp_Row_Count - 1
        j = arr_FileIndex(i + 1) - arr_FileIndex(i) - 2 '-2 за да махна vbCrLf
        sBuf = Space$(j)
        If Not cf.GetData(arr_FileIndex(i), 0, sBuf) Then Err.Raise 1000, , "cF.GetData(arr_FileIndex(i),0,sBuf)"
500     If Not RemSpase(sBuf) Then GoTo ErH
        
        sA = Split(sBuf, Chr$(32))
        'If Val(sA(0)) = 168 Then Stop
        flIn = True 'дали записа влиза в ограниченията
        For z = 1 To 5
            'Броя срещаните елементи
600         j = Val(sA(z))
            a(z, j) = a(z, j) + 1
            If l_A(z) Then If l_A(z) <> j Then flIn = False
            'Дали записа отговаря на търсените параметри
        Next z
1100
        If flIn Then
            'да си запиша номера на записа
            arRecIn_Count = arRecIn_Count + 1
            If UBound(arRecIn, 2) < arRecIn_Count Then ReDim Preserve arRecIn(5, arRecIn_Count + 5) As Long
            For z = 1 To 5
                arRecIn(z, arRecIn_Count) = Val(sA(z))
            Next z
        End If
    Next i
1200
    'Сега да видя кой числа търси
    'На кой записи ги има
    
    'За всяко число А-Е гледам има ли подадено твърдо число
    'Ако няма тогава
    
    N = 0
    If arRecIn_Count >= 0 Then
        ReDim rez(5, arRecIn_Count) As t_RL
    
        For i = 1 To 5
            j = l_A(i)
            If j = 0 Then
                'кои числа са се паднали и кое е с най голям рейтинг от тях
1500            For j = 0 To arRecIn_Count
                    z = arRecIn(i, j)           'това е числото
                    rez(i, j).Value = z
                    rez(i, j).Rang = a(i, z)    'какъв му е рейтинга
                    
                Next j
                'Сорт на колоната
                For j = 0 To arRecIn_Count - 1
                    Y = j
1600                x = rez(i, Y).Rang
                    For z = j + 1 To arRecIn_Count
                        If x < rez(i, z).Rang Then Y = z: x = rez(i, Y).Rang
                    Next z
                    If Y > j Then
                        'swap
1700                    rez(0, 0).Rang = rez(i, j).Rang
                        rez(0, 0).Value = rez(i, j).Value
                        rez(i, j).Rang = rez(i, Y).Rang
                        rez(i, j).Value = rez(i, Y).Value
                        rez(i, Y).Rang = rez(0, 0).Rang
                        rez(i, Y).Value = rez(0, 0).Value
                    End If
1800            Next j

                sBuf = "--"
                If rez(i, 0).Value Then sBuf = Right$("0" & rez(i, 0).Value, 2)
                
                Select Case i
                Case 1: lblA1.Caption = sBuf
                Case 2: lblB1.Caption = sBuf
                Case 3: lblC1.Caption = sBuf
                Case 4: lblD1.Caption = sBuf
                Case 5: lblE1.Caption = sBuf
                End Select
            End If
        Next i
        N = N + 1
        
    End If
2000    ReDim rez(5, 0) As t_RL
    'малка проверка за постоянно увеличаващи се числа
    z = 255: x = l_A(1)
    For i = 2 To 5
        If l_A(i) Then
            If x < l_A(i) Then
                x = l_A(i)
            Else
                z = 0: Exit For
            End If
        End If
    Next i
2100
    
    ReDim arRecIn(5) As Long
    arRecIn(0) = 1
    'явно няма записи
    'сега да мислим
    x = 0: Y = 39
    For i = 1 To 5
        'от каде до каде да търси
        If l_A(i) = 0 Then
            
2200
            For j = i + 1 To 5
                If l_A(j) Then Y = l_A(j): Y = Y - 1: Exit For
                Y = Y - 1
            Next j
2300
            If x = 0 Then
                For j = i - 1 To 1 Step -1
                    If l_A(j) Then x = l_A(j): x = x + 1: Exit For
                    x = x + 1
                Next j
            End If
2400
            If x > Y Then z = 0
            'FindMax
            If z = 0 Then Y = 0
            rez(i, 0).Rang = 0
2500
            For j = x To Y
                If rez(i, 0).Rang = 0 Then
                    rez(i, 0).Value = j 'това е числото
                    rez(i, 0).Rang = a(i, j)    'какъв му е рейтинга
                End If
                If rez(i, 0).Rang < a(i, j) Then
                    rez(i, 0).Value = j 'това е числото
                    rez(i, 0).Rang = a(i, j)    'какъв му е рейтинга
                End If
                
            Next j
2600
            '==================
            'вероятност да се падне
            If z Then
                arRecIn(i) = Y - x + 1
                arRecIn(0) = arRecIn(0) * arRecIn(i)
                x = rez(i, 0).Value + 1
                Y = 39
            End If
            '===============
2700
            sBuf = "--"
            If rez(i, 0).Value Then sBuf = Right$("0" & rez(i, 0).Value, 2) '& vbCrLf & Format$((1 / arRecIn(i) * 100), "#0.00") & "%"
            Select Case i
            Case 1: lblA2.Caption = sBuf
            Case 2: lblB2.Caption = sBuf
            Case 3: lblC2.Caption = sBuf
            Case 4: lblD2.Caption = sBuf
            Case 5: lblE2.Caption = sBuf
            End Select
        Else
            x = l_A(i) + 1
        End If
    Next i
2800
    N = N + 1
    lblPr2 = "Accuracy - " & Format$((1 / arRecIn(0) * 100), "#0.00") & "%"
    If N = 2 Then lblPr1.Caption = "Accuracy - " & Format$((1 / arRecIn(0) * 100), "#0.00") & "%"
    
    Obr = True
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function


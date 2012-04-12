VERSION 5.00
Begin VB.Form frmLoto 
   Caption         =   "Work"
   ClientHeight    =   6855
   ClientLeft      =   3600
   ClientTop       =   1995
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   10905
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   600
      ScaleHeight     =   5715
      ScaleWidth      =   9795
      TabIndex        =   6
      Top             =   240
      Width           =   9855
      Begin VB.Frame Frame3 
         Caption         =   "Frame1"
         Height          =   735
         Left            =   0
         TabIndex        =   32
         Top             =   2280
         Width           =   5295
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4680
            TabIndex        =   37
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   36
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   35
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   34
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   33
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame1"
         Height          =   735
         Left            =   0
         TabIndex        =   25
         Top             =   3960
         Width           =   5295
         Begin VB.Label lblPr2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblA2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   30
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblB2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   29
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblC2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   28
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblD2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   27
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblE2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4680
            TabIndex        =   26
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   735
         Left            =   0
         TabIndex        =   18
         Top             =   3120
         Width           =   5295
         Begin VB.Label lblA1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   24
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblB1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   23
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblC1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   22
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblD1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   21
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblE1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4680
            TabIndex        =   20
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblPr1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmbGetFixed 
         Caption         =   "Get Fixed"
         Height          =   495
         Left            =   7320
         TabIndex        =   16
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmbCleare 
         Caption         =   "Clear"
         Height          =   495
         Left            =   5520
         TabIndex        =   15
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmb_Exit 
         Caption         =   "&Exit"
         Height          =   495
         Left            =   7560
         TabIndex        =   8
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton cmbSettings 
         Caption         =   "&Settings"
         Height          =   495
         Left            =   5520
         TabIndex        =   7
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton cmb_GetNumb 
         Caption         =   "&Get Numbers"
         Height          =   495
         Left            =   5520
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtE 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtD 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtC 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtB 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtA 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1026
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   0
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Fixed numbers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   600
         TabIndex        =   14
         Top             =   0
         Width           =   7575
      End
      Begin VB.Label lblE 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblD 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   12
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblC 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   11
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblB 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   10
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblA 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   9
         Top             =   1080
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmLoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ModulIdString   As String = "frmLoto - "

'nFunction = 2005 -> Obr
'nFunction = 3004 -> RemSpase
'nFunction = 3005 -> fill_arrFileInd


Private Type CtrlProportions
    HeightProportions   As Single
    WidthProportions    As Single
    TopProportions      As Single
    LeftProportions     As Single
    FontSize            As Single
End Type
Private ProportionsArray()  As CtrlProportions

Private Type t_RL
    Value As Long
    Rang  As Long
End Type

Private l_A(5) As Long






Private Sub cmb_Exit_Click()
    Unload Me
End Sub

Private Sub cmb_GetNumb_Click()
    Obr
End Sub

Private Sub cmbCleare_Click()
    l_A(1) = 0
    l_A(2) = 0
    l_A(3) = 0
    l_A(4) = 0
    l_A(5) = 0
    
    txtA.Text = vbNullString
    txtB.Text = vbNullString
    txtC.Text = vbNullString
    txtD.Text = vbNullString
    txtE.Text = vbNullString
    
    lblA1.Caption = vbNullString
    lblB1.Caption = vbNullString
    lblC1.Caption = vbNullString
    lblD1.Caption = vbNullString
    lblE1.Caption = vbNullString
    
    lblA2.Caption = vbNullString
    lblB2.Caption = vbNullString
    lblC2.Caption = vbNullString
    lblD2.Caption = vbNullString
    lblE2.Caption = vbNullString
    
    lblPr2.Caption = vbNullString
    lblPr1.Caption = vbNullString
    
    txtA.SetFocus
    DoEvents
End Sub

Private Sub cmbSettings_Click()
    Unload Me
    frmSettings.Show
End Sub

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

Private Sub Form_Load()
Dim sB As String
    Get_Init_Data
    With Picture1
        .TOp = 0
        .Left = 0
        .Height = Me.Height
        .Width = Me.Width
        If FileExists(sPicFilePath) Then .Picture = LoadPicture(sPicFilePath)         ',vbLPCustom, vbLPColor,
    End With
    'History File
    'Results Count
    sB = " - Not set"
    If Len(sHistFilePath) Then sB = sHistFilePath
    Label1.Caption = "History File: " & sB & vbCrLf & vbCrLf & _
                      sMarckt_Message
    
    
End Sub

Private Sub Form_Resize()
   
    ResizeControls
    picStrech
End Sub
Sub picStrech()
Dim a As Long
    If Picture1.Picture Then
        a = Picture1.ScaleMode
        Picture1.ScaleMode = 3
        If Not Picture1.AutoRedraw Then Picture1.AutoRedraw = True
        Picture1.PaintPicture Picture1.Picture, _
            0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
            0, 0, _
            Picture1.Picture.Width / 26.46, _
            Picture1.Picture.Height / 26.46
        Picture1.Picture = Picture1.Image
        Picture1.ScaleMode = a
    End If
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
    lblA1.Caption = vbNullString
    lblB1.Caption = vbNullString
    lblC1.Caption = vbNullString
    lblD1.Caption = vbNullString
    lblE1.Caption = vbNullString
    
    lblA2.Caption = vbNullString
    lblB2.Caption = vbNullString
    lblC2.Caption = vbNullString
    lblD2.Caption = vbNullString
    lblE2.Caption = vbNullString
    
    lblPr2.Caption = vbNullString
    lblPr1.Caption = vbNullString
    
    DoEvents
11

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

            End If
        Next i
1900    ' Да проверя дали не са равни или не е сбъркана последователността
        If arRecIn_Count >= 10002 Then
            j = 1
            For i = 1 To 4
                If rez(i, j).Value Then
                    For z = i + 1 To 5
                        If rez(i, j).Value >= rez(z, j).Value Then
                            'Требва да се прави нещо
                            '1 Да видя кой е по-вероятен
                            If rez(i, j).Rang >= rez(z, j).Rang Then
                                x = z
                                For Y = j + 1 To arRecIn_Count
                                    'Вземам следващия по ред
                                    If rez(x, j).Value < rez(x, Y).Value Then
                                        'swap
                                        rez(0, 0).Value = rez(x, Y).Value
                                        rez(0, 0).Rang = rez(x, Y).Rang
                                        
                                        rez(x, Y).Value = rez(i, j).Value
                                        rez(x, Y).Rang = rez(i, j).Rang
                                        
                                        rez(i, j).Value = rez(0, 0).Value
                                        rez(i, j).Rang = rez(0, 0).Rang
                                        
                                        Y = -255
                                        Exit For
                                    End If
                                Next Y
                                
                            Else
                                x = i
                                For Y = j + 1 To arRecIn_Count
                                    'Вземам следващия по ред
                                    If rez(x, j).Value > rez(x, Y).Value Then
                                        
                                        'swap
                                        rez(0, 0).Value = rez(x, Y).Value
                                        rez(0, 0).Rang = rez(x, Y).Rang
                                        
                                        rez(x, Y).Value = rez(i, j).Value
                                        rez(x, Y).Rang = rez(i, j).Rang
                                        
                                        rez(i, j).Value = rez(0, 0).Value
                                        rez(i, j).Rang = rez(0, 0).Rang
                                        
                                        Y = -255
                                        Exit For
                                    End If
                                Next Y
                            End If
                         End If
                    Next z
                End If
            Next i
        End If

        'Запис на нещата
        For i = 1 To 5
            sBuf = "--"
            If rez(i, 1).Value Then sBuf = Right$("0" & rez(i, 1).Value, 2)
            
            Select Case i
            Case 1: lblA1.Caption = sBuf
            Case 2: lblB1.Caption = sBuf
            Case 3: lblC1.Caption = sBuf
            Case 4: lblD1.Caption = sBuf
            Case 5: lblE1.Caption = sBuf
            End Select
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

Private Function RemSpase(ByRef strIn As String) As Boolean
Const nFunction = 3004
On Error GoTo ErH
Dim sBuf As String
Dim z    As Long
Dim i    As Long
Dim fl1  As Boolean
10
    'Функция за премахване на повече от 1 шпация от средата на стринга
    For i = 1 To Len(strIn)
        z = Asc(Mid$(strIn, i, 1))
        If z > 32 Then
            If fl1 Then sBuf = sBuf & Chr$(32): fl1 = False
            sBuf = sBuf & Chr$(z)
        Else
            fl1 = True
        End If
    Next i
    strIn = Trim$(sBuf)
    RemSpase = True

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
    a = 1
    b = dbBAK.RecCount
    ReDim arrFileInd(100) As Long
    j = 0
    arrFileInd(j) = 1
    Do While b > 1
        If Len(sName) > b Then sName = Space$(b)
        If Not dbBAK.GetData(a, 0&, sName) Then GoTo ErH
300     i = 1
        Do
            i = InStr(i, sName, vbCrLf)
            If i Then
                j = j + 1
                If j > UBound(arrFileInd) Then ReDim Preserve arrFileInd(j + 100)
                arrFileInd(j) = (a - 1) + (i + 2)
                i = i + 2
            End If
        Loop While i > 0
        b = b - Len(sName)
        a = a + Len(sName)
    Loop
    RowCount = j 'Брой редове в файла (Разделени с CrLf)
400
    'Сега оправям последния ред
    j = j + 1
    If j > UBound(arrFileInd) Then ReDim Preserve arrFileInd(j + 1)
    'Последният ред да е дълък колкото най дългия
    arrFileInd(j) = (dbBAK.RecCount + 1 + 2)
    
    fill_arrFileInd = True

Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function


Private Sub txtA_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtB.SetFocus: txtA_Validate True
End Sub
Private Sub txtB_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtC.SetFocus: txtB_Validate True
End Sub
Private Sub txtC_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtD.SetFocus: txtC_Validate True
End Sub


Private Sub txtD_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtE.SetFocus: txtD_Validate True
End Sub
Private Sub txtE_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmb_GetNumb.SetFocus: txtE_Validate True
End Sub



Private Sub txtA_Validate(Cancel As Boolean)
    Call validateMe(txtA, l_A(1))
End Sub
Private Sub txtB_Validate(Cancel As Boolean)
    Call validateMe(txtB, l_A(2))
End Sub
Private Sub txtC_Validate(Cancel As Boolean)
    Call validateMe(txtC, l_A(3))
End Sub
Private Sub txtD_Validate(Cancel As Boolean)
    Call validateMe(txtD, l_A(4))
End Sub
Private Sub txtE_Validate(Cancel As Boolean)
    Call validateMe(txtE, l_A(5))
End Sub
Private Function validateMe(txt As TextBox, l As Long)
    l = Val(txt.Text)
    txt.Text = vbNullString
    If l < 1 Then l = 0
    If l > 39 Then l = 0
    If l Then txt.Text = Right$("0" & l, 2)
End Function

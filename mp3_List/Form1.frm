VERSION 5.00
Begin VB.Form frmShowList 
   Caption         =   " "
   ClientHeight    =   5010
   ClientLeft      =   2400
   ClientTop       =   1695
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   8280
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4875
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmb_Settings 
         Caption         =   "&Settings"
         Height          =   375
         Left            =   6600
         TabIndex        =   13
         ToolTipText     =   "Settings"
         Top             =   120
         Width           =   1335
      End
      Begin VB.PictureBox pH1 
         Height          =   375
         Left            =   7560
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2535
         Left            =   6240
         TabIndex        =   11
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton cmbGetList 
         Caption         =   "&Get List"
         Height          =   375
         Left            =   6600
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmbBrows 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   6600
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox txtList 
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   6
         Top             =   1200
         Width           =   6135
      End
      Begin VB.CommandButton cmbCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5160
         TabIndex        =   5
         Top             =   4080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.PictureBox scrCountFile 
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   4875
         TabIndex        =   4
         Top             =   4440
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.PictureBox scrCopyFile 
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   4875
         TabIndex        =   3
         Top             =   4080
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.CommandButton cmbOpenDoc 
         Caption         =   "&Open List"
         Height          =   375
         Left            =   6600
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmvExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   6600
         TabIndex        =   1
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label lblCapt 
         Alignment       =   2  'Center
         Caption         =   "Пътека"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmShowList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cScrolMaxCoefic As Long
Private arrFileInd()    As Long 'масив с началните байтове на редовете от файла

Private RowCount    As Long
Private Row_View    As Long
Private MaxRowLen   As Long
Private cf          As cFile

'За resize
Private Type CtrlProportions
    HeightProportions   As Double
    WidthProportions    As Double
    TopProportions      As Double
    LeftProportions     As Double
    FontSize            As Double
End Type
Dim ProportionsArray()  As CtrlProportions

 
Private Sub cmb_Settings_Click()
    Unload Me
    frmSettings.Show
    
End Sub

Private Sub cmbBrows_Click()
Dim s As String
    txtList = vbNullString
    s = RemDirSep(txtPath.Text)
    s = GetFolder("List from File", s, True)
    If Len(s) Then
        If txtPath.Text <> s Then
            txtPath.Text = s
            cSets.sDirPath = s
        End If
    End If
     
End Sub

Private Sub cmbCancel_Click()
    mdlMakeDirList.DoCancel = True
     
End Sub


Private Sub cmbGetList_Click()
Dim s As String
    
    
    txtList.Text = vbNullString
    s = Trim$(txtPath.Text)
    If Len(s) Then
        If DoObr(s, 0) Then ShowText
    End If
    
End Sub

Private Sub cmvExit_Click()

    mdlMakeDirList.DoCancel = True
    
    Set cSets = Nothing
    Set clsForm = Nothing
    
    Unload Me
End Sub

 

Private Sub Form_Initialize()
    'Me.Visible = False
    InitResizeArray
   
End Sub
Private Sub InitResizeArray()
Dim ScWidth  As Long
Dim ScHeight As Long

Dim i   As Integer
    
    On Error Resume Next
    ReDim ProportionsArray(0 To Controls.Count - 1)
    
    For i = 0 To Controls.Count - 1
       With ProportionsArray(i)
            .HeightProportions = Controls(i).Height / Me.ScaleHeight
            .WidthProportions = Controls(i).Width / Me.ScaleWidth
            .TopProportions = Controls(i).TOp / Me.ScaleHeight
            .LeftProportions = Controls(i).Left / Me.ScaleWidth
            .FontSize = Controls(i).Font.Size / Me.ScaleHeight
        End With
    Next i
End Sub

Private Sub ResizeControls()

    On Error Resume Next
    
    Dim i As Integer
    
    For i = 0 To Controls.Count - 1
        With ProportionsArray(i)
            ' move and resize controls
            Controls(i).Left = .LeftProportions * ScaleWidth
            Controls(i).TOp = .TopProportions * ScaleHeight
            Controls(i).Width = .WidthProportions * ScaleWidth
            Controls(i).Height = .HeightProportions * ScaleHeight
            
            Controls(i).Font.Size = .FontSize * ScaleHeight
        End With
    Next i
    
End Sub

Private Sub Form_Load()
    With Me
        .lblCapt.Caption = "I whant to make list of MP3 files from this folder"
        .txtPath.Text = cSets.sDirPath
        .txtPath.ToolTipText = "Select a FOLDER"
        .txtList.ToolTipText = "List of MP3 files"
        .Caption = "Make a list of MP3 files from FOLDER"
        
         With Picture1
            .TOp = 0
            .Left = 0
            .Height = Me.Height
            .Width = Me.Width
            If FileExists(cSets.sPicFile) Then .Picture = LoadPicture(cSets.sPicFile)
        End With
        
    End With
    ShowText
    
    
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



Private Sub cmbOpenDoc_Click()
On Error Resume Next
    If Not (cf Is Nothing) Then Set cf = Nothing
    Set cf = New cFile
    OpenDocument cSets.sOutDir & cSets.sOutFile & GetFileType & ".txt"
    
End Sub


Private Sub ShowText()
Dim need_hgt As Single
Dim s1 As String
Dim i As Long
    
    
    If Not (cf Is Nothing) Then Set cf = Nothing
    Set cf = New cFile
    
    cf.OpenFile cSets.sOutDir & cSets.sOutFile & GetFileType & ".txt", 1
     
    fill_arrFileInd
    
    s1 = vbNullString
    
    For i = 1 To 100
        s1 = s1 & Right$("   " & i, 3) & ". " & "Ред номер " & i & vbCrLf
    Next i
    'txtList.Text = s1
    'Да оправя верт. скрол
    pH1.Height = txtList.Height
    pH1.Font = txtList.Font
    need_hgt = pH1.TextWidth(s1) ' за 100 реда
    i = pH1.ScaleY(pH1.Height, pH1.ScaleMode, 3)
    Row_View = Int(i * 100 / need_hgt)
    'за RowCount - x
    need_hgt = need_hgt * RowCount / 100
    
    cScrolMaxCoefic = 1
    If need_hgt > 32700 Then cScrolMaxCoefic = need_hgt / 32700
    With VScroll1
        .Min = 1
        .max = need_hgt
        .SmallChange = (need_hgt / RowCount) * cScrolMaxCoefic
        .LargeChange = .SmallChange * 25
        
    End With
    VScroll1_Change
End Sub

 
Private Sub Form_Terminate()
    Set cSets = Nothing
End Sub

 

Private Sub txtList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    Select Case KeyCode
    Case vbKeyHome
        If Shift = vbCtrlMask Then VScroll1.Value = VScroll1.Min
    Case vbKeyEnd
        If Shift = vbCtrlMask Then VScroll1.Value = VScroll1.max
         
    Case vbKeyPageUp:
        'tek red
        i = (VScroll1.Value * RowCount / VScroll1.max) * cScrolMaxCoefic
        i = VScroll1.max * (i - Row_View) / RowCount
        If i < 1 Then i = 1
        If i > VScroll1.max Then i = VScroll1.max
        If VScroll1.Value > i Then VScroll1.Value = i
    Case vbKeyPageDown
        i = (VScroll1.Value * RowCount / VScroll1.max) * cScrolMaxCoefic
        i = VScroll1.max * (i + Row_View) / RowCount
        If i < 1 Then i = 1
        If i > VScroll1.max Then i = VScroll1.max
        If VScroll1.Value < i Then VScroll1.Value = i
        
    Case vbKeyDown
    
        i = (VScroll1.Value * RowCount / VScroll1.max) * cScrolMaxCoefic + 1
        i = VScroll1.max * (i + 1) / RowCount
        If i < 1 Then i = 1
        If i > VScroll1.max Then i = VScroll1.max
        
        If VScroll1.Value < i Then VScroll1.Value = i
        
    Case vbKeyUp
        i = (VScroll1.Value * RowCount / VScroll1.max) * cScrolMaxCoefic - 1
        i = VScroll1.max * (i - 1) / RowCount
        If i < 1 Then i = 1
        If i > VScroll1.max Then i = VScroll1.max
        If VScroll1.Value > i Then VScroll1.Value = i
    End Select
End Sub

Private Sub VScroll1_Change()
Dim i As Long
Dim z As Long
    z = txtList.SelStart
    i = (VScroll1.Value * RowCount / VScroll1.max) * cScrolMaxCoefic
    Call DataRecuest(i)
    txtList.SelStart = z
End Sub
Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub



Private Sub DataRecuest(ByVal iStart As Long)
    Dim sT As String
    Dim sBuf As String
    Dim i As Long
    Dim b As Long
    Dim c As Long
    
    If iStart < 0 Then iStart = 0
    If (RowCount - Row_View) > 0 Then
        If iStart > (RowCount - Row_View) Then iStart = RowCount - Row_View
    End If
    
    sT = vbNullString
    For i = iStart To iStart + Row_View
        If i > RowCount - 1 Then Exit For
        b = arrFileInd(i)      'Вземам от кои байт да чета
        c = arrFileInd(i + 1) - 2 - b 'Последен байт за реда
        If c < 0 Then c = 0 'Реда е къс
        sBuf = Space$(c)
        cf.GetData b, 0, sBuf
        If Len(sT) Then sT = sT & vbCrLf
        sBuf = Left$(sBuf & Space$(MaxRowLen), MaxRowLen)
        sT = sT & sBuf
    Next i
    
    txtList.Text = sT
    
End Sub

Private Function fill_arrFileInd() As Boolean
Const nFunction = 3006
On Error GoTo ErH
Dim sName       As String
Dim a           As Long
Dim b           As Long
Dim i           As Integer
Dim j           As Long

10
    'Функция за запис на началните позиции на редовете от файла във масив
    
    'Чета фаила и търся CRLF
    sName = Space$(1024)
    a = 1&
    b = cf.RecCount   '!!! dali nqma da
    MaxRowLen = 1
    ReDim arrFileInd(100) As Long
    j = 0
    arrFileInd(j) = 1
    Do While b > 1
        If Len(sName) > b Then sName = Space$(b)
        If Not cf.GetData(a, 0&, sName) Then GoTo ErH
300     i = 1
        Do
            i = InStr(i, sName, vbCrLf)
            If i Then
                j = j + 1
                If j > UBound(arrFileInd) Then ReDim Preserve arrFileInd(j + 10)
                arrFileInd(j) = (a - 1) + (i + 2)
                If arrFileInd(j) - arrFileInd(j - 1) > MaxRowLen Then MaxRowLen = arrFileInd(j) - arrFileInd(j - 1)
                i = i + 2
            End If
        Loop While i > 0
        b = b - Len(sName)
        a = a + Len(sName)
    Loop
400
    'Сега оправям последния ред
    j = j + 1
    If j > UBound(arrFileInd) Then ReDim Preserve arrFileInd(j + 10)
    'Последният ред да е дълък колкото най дългия
    arrFileInd(j) = (cf.RecCount + 1 + 2)
    RowCount = j
500
    
    fill_arrFileInd = True

Exit Function
ErH:
'SN.FatalException ModuleN, ModulIdString, nFunction, Erl, Err.Number, Err.Description
Err.Clear
    
End Function


Private Function GetFileType() As String
    
    Select Case cSets.sFileView
    Case 1 'Folder
        GetFileType = "_1"
    Case 2 'Files
        GetFileType = "_2"
    Case Else 'Full
        GetFileType = vbNullString
    End Select
End Function



'''
''''333333333333333333333333333333333
'''да се изгради лист от файлове
'''
'''
'''1. Избор на директория и файлови разширения от който ще се прави листът
'''2. Съставяне на листа
'''
'''3. Преглед и даване на възможност за запис на файла
'''
'''
'''
'''
'''
''''333333333333333333333333333333333
'''
















 
 

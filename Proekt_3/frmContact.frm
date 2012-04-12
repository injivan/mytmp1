VERSION 5.00
Begin VB.Form frmContact 
   BorderStyle     =   0  'None
   Caption         =   "Contact Information"
   ClientHeight    =   10500
   ClientLeft      =   2985
   ClientTop       =   1080
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   10500
      Left            =   120
      ScaleHeight     =   10440
      ScaleWidth      =   8955
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9015
      Begin VB.Timer Timer1 
         Left            =   7800
         Top             =   120
      End
      Begin VB.CommandButton cmbStart 
         Caption         =   "S&tart"
         Height          =   375
         Left            =   7200
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmbSettings 
         Caption         =   "&Settings"
         Height          =   375
         Left            =   7200
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   375
         Left            =   7200
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmbAdd 
         Caption         =   "Add Info"
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtMail 
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox txtNameLast 
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txtNameFirst 
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label lblMail 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E-Mail"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   2760
         Width           =   3015
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   1215
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   8895
      End
      Begin VB.Label lblNameLast 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Last Name"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label lblNameFirst 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Name"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   1800
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ModulIdString   As String = " frmContact - "
'Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" _
                        (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, _
                         ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long


'nFunction = 2000 -> Form_Load

'nFunction = 3001 -> Add_Data
'nFunction = 3002 -> Set_Lbl_Font



Private sName   As String
Private sName1  As String
Private sMail   As String

Public dTime As Double

Private Type CtrlProportions
    HeightProportions   As Single
    WidthProportions    As Single
    TopProportions      As Single
    LeftProportions     As Single
    FontSize            As Single
End Type
Private ProportionsArray()  As CtrlProportions

Private Sub cmbAdd_Click()
    Add_Data
End Sub

Private Sub cmbSettings_Click()
    
    frmSettings.Show
    Unload Me
End Sub

Public Sub cmbStart_Click()
    cmbSettings.Visible = False
    cmbStart.Visible = False
    frmContact.WindowState = 2
    frmContact.Show
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
Const nFunction = 2000
On Error GoTo ErH
    
    dTime = Timer
    Timer1.Interval = 1000
    Get_Init_Data
    With Picture1
        .TOp = 0
        .Left = 0
        .Height = Me.Height
        .Width = Me.Width
        If FileExists(PicFilePath) Then .Picture = LoadPicture(PicFilePath)         ',vbLPCustom, vbLPColor,
    End With
    
      
    
    With lblCaption
        If Font1_Color Then .ForeColor = Font1_Color
        .Caption = Marckt_Message
         
        
        
        
    End With
300
    If Font2_Color Then
        lblMail.ForeColor = Font2_Color
        lblNameFirst.ForeColor = Font2_Color
        lblNameLast.ForeColor = Font2_Color
    End If
    
    If Not Get_Font(1, Font1) Then GoTo ErH
    If Not (Font1 Is Nothing) Then
500     With lblCaption.Font
            .Name = Font1.Name
            .Charset = Font1.Charset
            
            .Size = Font1.Size
            .Bold = Font1.Bold
            .Italic = Font1.Italic
            .Underline = Font1.Underline
        End With
    End If
    
    If Not Get_Font(2, Font2) Then GoTo ErH
    If Not (Font2 Is Nothing) Then
        If Not Set_Lbl_Font(lblMail) Then GoTo ErH
        If Not Set_Lbl_Font(lblNameFirst) Then GoTo ErH
        If Not Set_Lbl_Font(lblNameLast) Then GoTo ErH
    End If
    
    
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub

Private Sub Form_Resize()
   
   ResizeControls
    picStrech
End Sub


Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Timer1_Timer()
    If Timer > dTime + 60 * 10 Then Unload Me
End Sub

Private Sub txtMail_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmbAdd.SetFocus
End Sub

Private Sub txtMail_Validate(Cancel As Boolean)
    txtMail.Text = MakeOneWord(txtMail.Text)
    sMail = txtMail.Text
End Sub

Private Sub txtNameFirst_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtNameLast.SetFocus
End Sub

Private Sub txtNameFirst_Validate(Cancel As Boolean)
    sName = MakeOneWord(txtNameFirst.Text)
    If Len(sName) Then Mid$(sName, 1, 1) = Chr$(Asc(Mid$(sName, 1, 1)) And Not &H20)
    txtNameFirst.Text = sName
End Sub

Private Sub txtNameLast_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtMail.SetFocus
End Sub

Private Sub txtNameLast_Validate(Cancel As Boolean)
    sName1 = MakeOneWord(txtNameLast.Text)
    If Len(sName1) Then Mid$(sName1, 1, 1) = Chr$(Asc(Mid$(sName1, 1, 1)) And Not &H20)
    txtNameLast.Text = sName1
End Sub

Private Function MakeOneWord(ByRef s1 As String) As String
Dim i As Long
    s1 = Trim$(s1)
    i = InStr(1, s1, Chr$(32))
    If i Then s1 = Mid$(s1, 1, i)
    s1 = Replace(s1, Chr$(44), Chr$(32))
    MakeOneWord = Trim$(s1)
End Function
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
 




Private Function Add_Data() As Boolean
Const nFunction = 3001
On Error GoTo ErH
Dim z As Long
Dim cf As cFile
10
    IzhFilePath = AddDirSep(IzhFilePath)
    Set cf = New cFile
    If Not cf.OpenFile(IzhFilePath & IzhFileName, 1) Then GoTo ErH
    z = cf.RecCount
    
    cf.PutData z, 0, sName & Chr$(44) & sName1 & Chr$(44) & sMail & vbCrLf
    
    cf.CloseFile

    sName = vbNullString
    sName1 = vbNullString
    sMail = vbNullString
    txtMail = vbNullString
    txtNameFirst = vbNullString
    txtNameLast = vbNullString
    
    txtNameFirst.SetFocus
    
    
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear

Set cf = Nothing
End Function

Private Function Set_Lbl_Font(ob As Control) As Boolean
Const nFunction = 3002
On Error GoTo ErH
10
    
    With ob.Font
        .Name = Font1.Name
        .Charset = Font1.Charset
        
        .Size = Font1.Size
        .Bold = Font1.Bold
        .Italic = Font1.Italic
        .Underline = Font1.Underline
    End With
    
    Set_Lbl_Font = True
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear

End Function



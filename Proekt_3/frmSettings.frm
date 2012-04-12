VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   8385
   ClientLeft      =   4455
   ClientTop       =   1515
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   11400
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   600
      ScaleHeight     =   4515
      ScaleWidth      =   10395
      TabIndex        =   10
      Top             =   240
      Width           =   10455
      Begin VB.TextBox txtTestLabels 
         Height          =   375
         Left            =   2160
         MaxLength       =   18
         TabIndex        =   17
         Text            =   "Test text"
         ToolTipText     =   "Name of the output file."
         Top             =   3000
         Width           =   2175
      End
      Begin VB.CommandButton cmb_Color_Lab 
         Caption         =   "Color  For Labels"
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton cmb_Color1 
         Caption         =   "Color"
         Height          =   375
         Left            =   8760
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmb_Font_Lab 
         Caption         =   "Font For Labels"
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         ToolTipText     =   "Font for the labels."
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton cmbExit 
         Caption         =   "Exit"
         Height          =   315
         Left            =   8760
         TabIndex        =   15
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox txtMrcMessage 
         Height          =   975
         Left            =   120
         MaxLength       =   256
         MultiLine       =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Type your marketing message. Max 256 characters."
         Top             =   480
         Width           =   8535
      End
      Begin VB.TextBox txtBackImage 
         Height          =   375
         Left            =   2160
         MaxLength       =   128
         TabIndex        =   3
         ToolTipText     =   "Path and name of the background image."
         Top             =   1560
         Width           =   6495
      End
      Begin VB.CommandButton cmb_BrImage 
         Caption         =   "Browse"
         Height          =   375
         Left            =   8760
         TabIndex        =   4
         ToolTipText     =   "Path and name of the background image."
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtOutPath 
         Height          =   375
         Left            =   2160
         MaxLength       =   128
         TabIndex        =   5
         ToolTipText     =   "Path to the output file."
         Top             =   2040
         Width           =   6495
      End
      Begin VB.CommandButton cmbBr_OutPath 
         Caption         =   "Browse"
         Height          =   375
         Left            =   8760
         TabIndex        =   6
         ToolTipText     =   "Path to the output file."
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmb_FontMess 
         Caption         =   "Font"
         Height          =   375
         Left            =   8760
         TabIndex        =   1
         ToolTipText     =   "Set the font for the marketing message."
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtOutName 
         Height          =   375
         Left            =   2160
         MaxLength       =   18
         TabIndex        =   7
         ToolTipText     =   "Name of the output file."
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Labels Font"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Marketing message"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Background image"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Out File Path"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Out File Name"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ModulIdString   As String = " frmSettings - "
'nFunction = 2000 -> Form_Load
'nFunction = 2001 -> cmb_FontMess_Click
'nFunction = 2002 -> cmb_Color1_Click
'nFunction = 2003 -> cmb_Font_Lab_Click
'nFunction = 2004 -> cmb_Color_Lab_Click
'nFunction = 2005 -> cmb_BrImage_Click
'nFunction = 2006 -> cmbBr_OutPath_Click


'nFunction = 3000 -> Set_Font_McrMessage
'nFunction = 3001 -> Set_Font_TestText

Private Type CtrlProportions
    HeightProportions   As Single
    WidthProportions    As Single
    TopProportions      As Single
    LeftProportions     As Single
    FontSize            As Single
End Type
Private ProportionsArray()  As CtrlProportions



Private Sub cmb_BrImage_Click()
Const nFunction = 2005
Dim a As String

On Error GoTo ErH

    Dim sF As clsFileDlg
    Set sF = New clsFileDlg
    
    With sF
        a = PicFilePath
        Call .VBGetOpenFileName(a, , True, , , , "BMP (*.bmp)|*.bmp|JPG (*.jpg)|*.jpg|All  (*.*)| *.*", , , "Open file")
        If Len(a) Then
            If a <> PicFilePath Then
                PicFilePath = a
                Set_Init_Data RD_NameFon_1
                txtBackImage = PicFilePath
            End If
        End If
    End With
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub
Private Sub cmbBr_OutPath_Click()
Const nFunction = 2006
Dim a As String
On Error GoTo ErH

    a = IzhFilePath
    If Len(a) = 0 Then a = MePath
    a = GetFolder("Output Folder", a, True)
    
    If Len(a) Then
        If a <> IzhFilePath Then
            
            IzhFilePath = a
            Set_Init_Data RD_OutFilePath_1
            txtOutPath = IzhFilePath
        End If
    End If
    
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub
Private Sub cmb_Color_Lab_Click()
Const nFunction = 2004
On Error GoTo ErH

Dim oC As clsFileDlg
Dim a As Long
    
    Set oC = New clsFileDlg
    a = Font2_Color
     
    If oC.VBChooseColor(a, , , , Me.hWnd) Then
        Font2_Color = a
        Write_Reg_String MePath, RD_Font2_Color, CStr(a)
        If Not Set_Font_TestText Then GoTo ErH
    End If
    Set oC = Nothing
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub

Private Sub cmb_Color1_Click()
Const nFunction = 2002
On Error GoTo ErH

Dim oC As clsFileDlg
Dim a As Long
    
    Set oC = New clsFileDlg
    a = Font1_Color
     
    If oC.VBChooseColor(a, , , , Me.hWnd) Then
        Font1_Color = a
        Write_Reg_String MePath, RD_Font1_Color, CStr(a)
        If Not Set_Font_McrMessage Then GoTo ErH
    End If
    Set oC = Nothing
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub

Private Sub cmb_Font_Lab_Click()
Const nFunction = 2003
On Error GoTo ErH
Dim oF          As clsFileDlg
Dim fFont       As Boolean
Dim oldFont     As StdFont
10
    
    If Font2 Is Nothing Then
12      Set oldFont = New StdFont
    Else
21      oldFont.Name = Font2.Name
        oldFont.Charset = Font2.Charset
        oldFont.Size = Font2.Size
        oldFont.Bold = Font2.Bold
        oldFont.Italic = Font2.Italic
        oldFont.Underline = Font2.Underline
        
    End If
    
    Set oF = New clsFileDlg
    fFont = oF.VBChooseFont(oldFont, , Me.hWnd)
    If fFont Then
        Call Set_Font(2, oldFont)
        Call Get_Font(2, Font2)
        If Not Set_Font_TestText Then GoTo ErH
    End If
    
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub

Private Sub cmb_FontMess_Click()
Const nFunction = 2001
On Error GoTo ErH
Dim oF          As clsFileDlg
Dim fFont       As Boolean
Dim oldFont     As StdFont
10
    
    If Font1 Is Nothing Then
12      Set oldFont = New StdFont
    Else
21      oldFont.Name = Font1.Name
        oldFont.Charset = Font1.Charset
        oldFont.Size = Font1.Size
        oldFont.Bold = Font1.Bold
        oldFont.Italic = Font1.Italic
        oldFont.Underline = Font1.Underline
        
    End If
    
    Set oF = New clsFileDlg
    fFont = oF.VBChooseFont(oldFont, , Me.hWnd)
    If fFont Then
        Call Set_Font(1, oldFont)
        Call Get_Font(1, Font1)
                  
        If Not Set_Font_McrMessage Then GoTo ErH
        
    End If
    
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub



Private Sub cmbExit_Click()
    Unload Me
    frmContact.Show
    frmContact.cmbStart_Click
End Sub

Private Sub Form_Initialize()
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

Private Sub Form_Load()
Const nFunction = 2000
On Error Resume Next
10
    Get_Init_Data
    
    With Picture1
        .TOp = 0
        .Left = 0
        .Height = Me.Height
        .Width = Me.Width
    End With
    
    If Not Set_Font_McrMessage Then GoTo ErH
    If Not Set_Font_TestText Then GoTo ErH
    txtBackImage.Text = PicFilePath
    txtOutPath.Text = IzhFilePath
    txtOutName.Text = IzhFileName
    
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub

Private Function Set_Font_McrMessage() As Boolean
Const nFunction = 3000
On Error GoTo ErH

10
    With txtMrcMessage
        .Text = Marckt_Message
        If Font1_Color Then .ForeColor = Font1_Color
    
        If Not (Font1 Is Nothing) Then
            With .Font
                .Name = Font1.Name
                .Charset = Font1.Charset
                
                .Size = Font1.Size
                .Bold = Font1.Bold
                .Italic = Font1.Italic
                .Underline = Font1.Underline
            End With
        End If
    End With
    
    Set_Font_McrMessage = True

Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function

Private Function Set_Font_TestText() As Boolean
Const nFunction = 3001
On Error GoTo ErH

10
    With txtTestLabels
        If Font2_Color Then .ForeColor = Font2_Color
    
        If Not (Font2 Is Nothing) Then
            With .Font
                .Name = Font2.Name
                .Charset = Font2.Charset
                
                .Size = Font2.Size
                .Bold = Font2.Bold
                .Italic = Font2.Italic
                .Underline = Font2.Underline
            End With
        End If
    End With
    
    Set_Font_TestText = True

Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function

 

Private Sub txtBackImage_Validate(Cancel As Boolean)
Dim a As String
    a = Trim$(txtBackImage.Text)
    If a <> PicFilePath Then
        PicFilePath = a
        Set_Init_Data RD_NameFon_1
        txtBackImage.Text = PicFilePath
    End If
End Sub

Private Sub txtMrcMessage_Validate(Cancel As Boolean)
Dim a As String
    a = Trim$(txtMrcMessage.Text)
    If a <> Marckt_Message Then
        Marckt_Message = a
        Set_Init_Data RD_Marckt_Message_1
        txtMrcMessage.Text = Marckt_Message
    End If
End Sub

Private Sub txtOutName_Validate(Cancel As Boolean)
Dim a As String
    a = Trim$(txtOutName.Text)
    If a <> IzhFileName Then
        IzhFileName = a
        Set_Init_Data RD_OutFileName
        txtOutName.Text = IzhFileName
    End If
End Sub

Private Sub txtOutPath_Validate(Cancel As Boolean)
Dim a As String
    a = Trim$(txtOutPath.Text)
    If a <> IzhFilePath Then
        IzhFilePath = a
        Set_Init_Data RD_OutFilePath_1
        txtOutPath = IzhFilePath
    End If
End Sub

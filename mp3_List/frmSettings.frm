VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   5820
   ClientLeft      =   5250
   ClientTop       =   3045
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   9150
   Begin VB.PictureBox Picture1 
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5235
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   240
      Width           =   8655
      Begin VB.Frame frmFileView 
         Caption         =   "View"
         Height          =   735
         Left            =   2040
         TabIndex        =   13
         Top             =   3600
         Width           =   5895
         Begin VB.OptionButton opFiles 
            Caption         =   "Files"
            Height          =   255
            Left            =   3960
            TabIndex        =   16
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton opFolder 
            Caption         =   "Foldes"
            Height          =   255
            Left            =   2100
            TabIndex        =   15
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton opFull 
            Caption         =   "Full"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.TextBox txtOutPutName 
         Height          =   375
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   11
         ToolTipText     =   "Name of the Output file."
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox txtOutFilePath 
         Height          =   375
         Left            =   2040
         MaxLength       =   128
         TabIndex        =   9
         ToolTipText     =   "Path of the Output file."
         Top             =   2640
         Width           =   4455
      End
      Begin VB.CommandButton cmb_OutPutDir 
         Caption         =   "Br&owse"
         Height          =   375
         Left            =   6720
         TabIndex        =   8
         ToolTipText     =   "Path of the Output file."
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmbHist 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   6720
         TabIndex        =   6
         ToolTipText     =   "Chose a search Directory"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtHistFile 
         Height          =   375
         Left            =   2040
         MaxLength       =   128
         TabIndex        =   5
         ToolTipText     =   "Last searched Directory"
         Top             =   1680
         Width           =   4455
      End
      Begin VB.CommandButton cmb_BrImage 
         Caption         =   "B&rowse"
         Height          =   375
         Left            =   6720
         TabIndex        =   3
         ToolTipText     =   "Path and name of the background image."
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtBackImage 
         Height          =   375
         Left            =   2040
         MaxLength       =   128
         TabIndex        =   2
         ToolTipText     =   "Path and name of the background image."
         Top             =   2160
         Width           =   4455
      End
      Begin VB.CommandButton cmbExit 
         Caption         =   "&Exit"
         Height          =   435
         Left            =   7080
         TabIndex        =   1
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Output File Name"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Output File Path"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Last Dir"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "Last searched Directory"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Background image"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   2160
         Width           =   1695
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

'nFunction = 2004 -> cmbHist_Click
'nFunction = 2005 -> cmb_BrImage_Click
'nFunction = 2006 -> cmb_OutPutDir_Click()

Private Type CtrlProportions
    HeightProportions   As Single
    WidthProportions    As Single
    TopProportions      As Single
    LeftProportions     As Single
    FontSize            As Single
End Type
Private ProportionsArray()  As CtrlProportions



Private Sub cmbExit_Click()
 
    Unload Me
    clsForm.Show
    'frmShowList.Show
    
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
     
    
    With Picture1
        .TOp = 0
        .Left = 0
        .Height = Me.Height
        .Width = Me.Width
    End With
    
    txtBackImage.Text = cSets.sPicFile
    txtHistFile.Text = cSets.sDirPath
    txtOutFilePath.Text = cSets.sOutDir
    txtOutPutName.Text = cSets.sOutFile
    
    opFiles.Value = (cSets.sFileView = 2)
    opFull.Value = (cSets.sFileView = 0)
    opFolder.Value = (cSets.sFileView = 1)
    
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub

Private Sub opFiles_Click()
    If opFiles.Value Then If cSets.sFileView <> 2 Then cSets.sFileView = 2
End Sub

Private Sub opFolder_Click()
    If opFolder.Value Then If cSets.sFileView <> 1 Then cSets.sFileView = 1
End Sub

Private Sub opFull_Click()
    If opFull.Value Then If cSets.sFileView <> 0 Then cSets.sFileView = 0
End Sub

Private Sub txtBackImage_Validate(Cancel As Boolean)
Dim a As String
    a = Trim$(txtBackImage.Text)
    If a <> cSets.sPicFile Then
        cSets.sPicFile = a
        txtBackImage.Text = a
    End If
End Sub

 
Private Sub txtHistFile_Validate(Cancel As Boolean)
Dim a As String
    a = Trim$(txtHistFile.Text)
    If a <> cSets.sDirPath Then
        cSets.sDirPath = a
        txtHistFile.Text = a
    End If
End Sub
  
Private Sub txtOutFilePath_Validate(Cancel As Boolean)
    Dim a As String
    a = Trim$(txtOutFilePath.Text)
    If a <> cSets.sOutDir Then
        cSets.sOutDir = a
        txtOutFilePath.Text = a
    End If
End Sub
Private Sub txtOutPutName_Validate(Cancel As Boolean)
    Dim a As String
    a = Trim$(txtOutPutName.Text)
    If a <> cSets.sOutFile Then
        cSets.sOutFile = a
        txtOutPutName.Text = a
    End If
End Sub



 
Private Sub cmbHist_Click()
Const nFunction = 2004
Dim a As String
On Error GoTo ErH
10
    a = RemDirSep(cSets.sDirPath)
    a = GetFolder("List from File", a, True)
    If Len(a) Then
        If a <> cSets.sDirPath Then
            cSets.sDirPath = a
            txtHistFile.Text = a
        End If
    End If
    
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub
Private Sub cmb_BrImage_Click()
Const nFunction = 2005
Dim a As String
On Error GoTo ErH
10
    Dim sF As clsFileDlg
    Set sF = New clsFileDlg
    
    With sF
        a = cSets.sPicFile
        Call .VBGetOpenFileName(a, , True, , , , "JPG (*.jpg)|*.jpg|GIF (*.gif)|*.gif|BMP (*.bmp)|*.bmp|All  (*.*)| *.*", , , "Open file")
        If Len(a) Then
            If a <> cSets.sPicFile Then
                cSets.sPicFile = a
                txtBackImage.Text = cSets.sPicFile
            End If
        End If
    End With
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub
Private Sub cmb_OutPutDir_Click()
Const nFunction = 2006
Dim a As String
On Error GoTo ErH
10
    a = RemDirSep(cSets.sDirPath)
    a = GetFolder("List from File", a, True)
    If Len(a) Then
        If a <> cSets.sOutDir Then
            cSets.sOutDir = a
            txtOutFilePath.Text = a
        End If
    End If
    
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub


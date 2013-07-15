VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   6825
   ClientLeft      =   3030
   ClientTop       =   1185
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   9975
   Begin VB.PictureBox Picture1 
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5235
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   240
      Width           =   8655
      Begin VB.TextBox txtOutFilePath 
         Height          =   375
         Left            =   2040
         MaxLength       =   128
         TabIndex        =   16
         ToolTipText     =   "Path of the Output file."
         Top             =   2640
         Width           =   4455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6720
         TabIndex        =   15
         ToolTipText     =   "Path and name of the background image."
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtEndRec 
         Height          =   405
         Left            =   5760
         TabIndex        =   13
         ToolTipText     =   "Read ot End nomber record. If not definet then End rec is the last record from the file."
         Top             =   3600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtStartRec 
         Height          =   405
         Left            =   3360
         TabIndex        =   11
         ToolTipText     =   "Read From nomber record. If not definet then star from rec 1"
         Top             =   3600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmbHist 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6720
         TabIndex        =   8
         ToolTipText     =   "Path and name of the background image."
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtHistFile 
         Height          =   375
         Left            =   2040
         MaxLength       =   128
         TabIndex        =   7
         ToolTipText     =   "Path and name of the background image."
         Top             =   1680
         Width           =   4455
      End
      Begin VB.CommandButton cmb_BrImage 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6720
         TabIndex        =   4
         ToolTipText     =   "Path and name of the background image."
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtBackImage 
         Height          =   375
         Left            =   2040
         MaxLength       =   128
         TabIndex        =   3
         ToolTipText     =   "Path and name of the background image."
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txtMrcMessage 
         Height          =   975
         Left            =   480
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Type your message. Max 128 characters."
         Top             =   600
         Width           =   6015
      End
      Begin VB.CommandButton cmbExit 
         Caption         =   "Exit"
         Height          =   435
         Left            =   7080
         TabIndex        =   1
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Output File Path"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "End Record"
         Height          =   375
         Left            =   4440
         TabIndex        =   14
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Start Record"
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Recors Read"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   3600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "History File"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Background image"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Any message"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   2415
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
    MePath = App.Path
    sINI_Name = "E_Mail.ini"
    
    Get_Init_Data
    
    With Picture1
        .TOp = 0
        .Left = 0
        .Height = Me.Height
        .Width = Me.Width
    End With
    
    txtMrcMessage.Text = sMarckt_Message
   
    txtBackImage.Text = sPicFilePath
    txtHistFile.Text = sHistFilePath
     
    
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub

Private Sub txtBackImage_Validate(Cancel As Boolean)
Dim a As String
    a = Trim$(txtBackImage.Text)
    If a <> sPicFilePath Then
        sPicFilePath = a
        Set_Init_Data RD_NameFon_1
        txtBackImage.Text = sPicFilePath
    End If
End Sub
Private Sub txtHistFile_Validate(Cancel As Boolean)
Dim a As String
    a = Trim$(txtHistFile.Text)
    If a <> sHistFilePath Then
        sHistFilePath = a
        Set_Init_Data RD_HistFilePath_1
        txtHistFile.Text = sHistFilePath
    End If
End Sub


Private Sub txtMrcMessage_Validate(Cancel As Boolean)
Dim a As String
    a = Trim$(txtMrcMessage.Text)
    If a <> sMarckt_Message Then
        sMarckt_Message = a
        Set_Init_Data RD_Marckt_Message_1
        txtMrcMessage.Text = sMarckt_Message
    End If
End Sub
Private Sub cmb_BrImage_Click()
Const nFunction = 2005
Dim a As String

On Error GoTo ErH


Dim apW         As Word.Application
Dim apW_P     As Word.Paragraph


Set apW = CreateObject("Word.Application")
apW.Documents.Open ("English.doc")
apW.Visible = True
Dim z As String
Dim i As Long
Dim j As Long
For Each apW_P In ActiveDocument.Paragraphs
    z = Trim$(apW_P.Range.Text)
    i = 1: j = 1
    Do
        j = InStr(i, z, vbLf)
        If j Then
            Debug.Print Mid$(z, i, j): i = j + 1
        Else
            Debug.Print Mid$(z, i)
        End If
    Loop While j
    
Next

apW.Documents.Close
Set apW = Nothing




''''    Dim sF As clsFileDlg
''''    Set sF = New clsFileDlg
''''
''''    With sF
''''        a = sPicFilePath
''''        Call .VBGetOpenFileName(a, , True, , , , "JPG (*.jpg)|*.jpg|GIF (*.gif)|*.gif|BMP (*.bmp)|*.bmp|All  (*.*)| *.*", , , "Open file")
''''        If Len(a) Then
''''            If a <> sPicFilePath Then
''''                sPicFilePath = a
''''                Set_Init_Data RD_NameFon_1
''''                txtBackImage.Text = sPicFilePath
''''            End If
''''        End If
''''    End With
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub
Private Sub cmbHist_Click()

Const nFunction = 2004
Dim a As String

On Error GoTo ErH

    Dim sF As clsFileDlg
    Set sF = New clsFileDlg
    
    With sF
        a = sHistFilePath
        Call .VBGetOpenFileName(a, , True, , , , "TXT (*.txt)|*.txt|All  (*.*)| *.*", , , "Open file")
        If Len(a) Then
            If a <> sHistFilePath Then
                sHistFilePath = a
                Set_Init_Data RD_HistFilePath_1
                txtHistFile.Text = sHistFilePath
            End If
        End If
    End With
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub

 

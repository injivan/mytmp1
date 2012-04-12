VERSION 5.00
Begin VB.Form frmXML_Parser 
   Caption         =   "XML Parser"
   ClientHeight    =   4005
   ClientLeft      =   4260
   ClientTop       =   1200
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   8160
   Begin VB.CommandButton cmbExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton cmbPars 
      Caption         =   "&Pars File"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "File that will be be changing"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   6255
      End
      Begin VB.CommandButton cmbBrowse 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   5520
         TabIndex        =   1
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "File"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmXML_Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ModulIdString   As String = "frmXML_Parser"
Private Const sINI_Name       As String = "Pars.ini"

'nFunction = 3001 -> cmbBrowse_Click
'nFunction = 3002 -> txtPath_Validate
'nFunction = 3003 -> cmbPars_Click
'nFunction = 3013 -> read_ini_file
'nFunction = 3014 -> SetLastFolder



'За resize
Private Type CtrlProportions
    HeightProportions   As Double
    WidthProportions    As Double
    TopProportions      As Double
    LeftProportions     As Double
    FontSize            As Double
End Type
Private ProportionsArray()  As CtrlProportions

Private sLastPath  As String
Private sLastFName As String
Private xmlDoc     As XMLDocument



Private Sub InitResizeArray()
On Error Resume Next
Dim ScWidth  As Long
Dim ScHeight As Long

Dim i   As Integer
    
    
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


 

Private Sub cmbExit_Click()
    Unload Me
End Sub



Private Sub Form_Initialize()
    InitResizeArray
End Sub

Private Sub Form_Load()
    Call read_ini_file
    txtPath.Text = sLastPath
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub
Private Sub cmbBrowse_Click()
Const nFunction = 3001
On Error GoTo ErH
Dim sF  As clsFileDlg
Dim a   As String
10
    
    Set sF = New clsFileDlg
    With sF
        Call .VBGetOpenFileName(a, , True, , , , "Text (*.txt)|*.txt|All  (*.*)| *.*", , sLastPath, "Open file for changing")
        If Len(a) Then
100         ExtractPathFile a, sLastPath, sLastFName
            txtPath.Text = sLastPath & sLastFName
200         If Not SetLastFolder(a) Then GoTo ErH
        End If
    End With

Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub
Private Sub txtPath_Validate(Cancel As Boolean)
Const nFunction = 3002
On Error GoTo ErH
Dim a As String
Dim b As String
10

    'Взема данните от полето.
    'Ако е валидна директория записва я в файла и в променливата
    'иначе записва само в променливата
    
    a = Trim$(txtPath.Text)
    ExtractPathFile a, b, a
    If DirExists(b) Then
        If Not SetLastFolder(b & a) Then GoTo ErH
        sLastPath = b
        sLastFName = a
    End If
    

Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub
Private Sub cmbPars_Click()
Const nFunction = 3003
On Error GoTo ErH
Dim sFileMaket As String
10
    sFileMaket = sLastPath & sLastFName
    Set xmlDoc = New XMLDocument
    If Not xmlDoc.LoadData(sFileMaket, False, True) Then GoTo ErH
    
Exit Sub
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Sub


Private Function read_ini_file() As Boolean
Const nFunction = 3013
On Error GoTo ErH
Dim ff   As Long
Dim sBuf As String
Dim sTmp As String
Dim z    As Long
Dim h    As Long

10
    sTmp = AddDirSep(App.Path) & sINI_Name
100

    '  1 -  3 b ASCII - Номер на параметър
    '  4 -  1 b ASCII - "."
    '  5 - 59 b ASCII - Описание на параметъра
    ' 64 -  1 b ASCII - "="
    ' 65 - 64 b ASCII - Данни за параметъра
    '----------------------------------
    'Общ. 128 b
    
    
    If Not FileExists(sTmp) Then
        ff = FreeFile
        Open sTmp For Binary As #ff
        sBuf = Space$(128)
        Mid$(sBuf, 1, 3) = "1"
        Mid$(sBuf, 4, 1) = "."
        Mid$(sBuf, 5) = "Последно ползвана пътека"
        Mid$(sBuf, 64, 1) = "="
        Mid$(sBuf, 65) = App.Path
        
        Put #ff, 1, sBuf
        Close #ff: ff = 0
    End If
200
    ff = FreeFile
    Open sTmp For Binary As #ff
     

    z = FileLen(sTmp)
    h = 1
    sBuf = Space$(128)
    Do
        If Len(sBuf) > z Then sBuf = Space$(z)
        Get #ff, h, sBuf
        If Trim$(Left$(sBuf, 3)) = "1" Then
            sLastPath = Trim$(Mid$(sBuf, 65))
            Exit Do
        End If
        z = z - 128
        h = h + 128
    Loop While z
    
300
    read_ini_file = True

 
Exit Function
ErH:
If ff Then Close #ff: ff = 0
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Function

Private Function SetLastFolder(ByVal sNewFolder As String) As Boolean
Const nFunction = 3014
On Error GoTo ErH
Dim ff   As Long
Dim sBuf As String
Dim sTmp As String
Dim z    As Long
Dim h    As Long
10
    sTmp = AddDirSep(App.Path) & sINI_Name
100
    ff = FreeFile
    Open sTmp For Binary As #ff
    z = FileLen(sTmp)
    h = 1
    sBuf = Space$(128)
    Do
        If Len(sBuf) > z Then sBuf = Space$(z)
        Get #ff, h, sBuf
        If Trim$(Left$(sBuf, 3)) = "1" Then
            Mid$(sBuf, 65) = Left$(sNewFolder & Space$(64), 64)
            Put #ff, h, sBuf
            Exit Do
        End If
        z = z - 128
        h = h + 128
    Loop While z
        
    Close #ff: ff = 0
    
300
    SetLastFolder = True

 
Exit Function
ErH:
If ff Then Close #ff: ff = 0
ShowErrMesage Err, ModulIdString, nFunction, Erl

Err.Clear
End Function













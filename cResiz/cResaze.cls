VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cResaze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type CtrlProportions
    HeightProportions   As Single
    WidthProportions    As Single
    TopProportions      As Single
    LeftProportions     As Single
    FontSize            As Single
End Type
Private ProportionsArray()  As CtrlProportions

Private Type tFlexCell
    ColWidth()  As Double
    RowHeight() As Double
    ColCo       As Long
    RowCo       As Long
End Type
Private arCell()  As tFlexCell

Private msFlex()  As MSFlexGrid
Private msFlex_Co As Long

Private mScaleHeight        As Long
Private mScaleWidth         As Long
Private mSH                 As Long
Private mSW                 As Long


Private WithEvents frmMe As Form
Attribute frmMe.VB_VarHelpID = -1
 

Public Property Let GetFlexGrid(vFlex As MSFlexGrid)
Dim i As Long
Dim m As Double
    If msFlex_Co = 0 Then
        ReDim msFlex(2) As MSFlexGrid
        ReDim arCell(2) As tFlexCell
        
    End If
    If msFlex_Co > UBound(msFlex) Then
        ReDim Preserve msFlex(msFlex_Co + 2) As MSFlexGrid
        ReDim Preserve arCell(msFlex_Co + 2) As tFlexCell
    End If
    Set msFlex(msFlex_Co) = vFlex
    With arCell(msFlex_Co)
        .ColCo = vFlex.Cols - 1
        .RowCo = vFlex.Rows - 1
        ReDim .ColWidth(.ColCo)
        ReDim .RowHeight(.RowCo)
        m = 0
        For i = 0 To .ColCo
            m = m + vFlex.ColWidth(i)
            .ColWidth(i) = vFlex.ColWidth(i) / frmMe.ScaleWidth
        Next i
        If vFlex.Width - m - (.ColCo * 60) > 0 Then
            m = (vFlex.Width - m - (.ColCo * 60)) / m
        Else
            m = 0
        End If
        
        For i = 0 To .ColCo
            .ColWidth(i) = (vFlex.ColWidth(i) + m * vFlex.ColWidth(i)) / frmMe.ScaleWidth
        Next i
        'If z - vFlex.ColWidth(i) > 0 Then vFlex.ColWidth(i) = z
        '.ColWidth(i) = vFlex.ColWidth(i) / frmMe.ScaleWidth
        
        For i = 0 To .RowCo '- 1
            .RowHeight(i) = vFlex.RowHeight(i) / frmMe.ScaleHeight
        Next i
    End With
    
    msFlex_Co = msFlex_Co + 1
    
    
End Property
Public Function GetFrm(vFrm As Form, frmScaleHeight As Long, frmScaleWidth As Long)

    Set frmMe = vFrm
    
    mSH = frmMe.ScaleHeight
    mSW = frmMe.ScaleWidth
     
    mScaleHeight = frmScaleHeight
    mScaleWidth = frmScaleWidth
    
    InitResizeArray
End Function

 

Private Sub InitResizeArray()
Dim i As Integer
Dim z As Boolean
On Error Resume Next
    
    With frmMe
        ReDim ProportionsArray(0 To .Controls.Count - 1)
        For i = 0 To .Controls.Count - 1
            z = False
            If .Controls(i).Left < 0 Then
                .Controls(i).Left = .Controls(i).Left + 75000
                z = True
            End If
            
            Select Case .Controls(i).Name
            Case "a" '"Picture1", "Image1"
            Case Else
                ProportionsArray(i).HeightProportions = .Controls(i).Height / mScaleHeight
                ProportionsArray(i).WidthProportions = .Controls(i).Width / mScaleWidth
                ProportionsArray(i).TopProportions = .Controls(i).Top / mScaleHeight
                ProportionsArray(i).LeftProportions = .Controls(i).Left / mScaleWidth
                ProportionsArray(i).FontSize = .Controls(i).Font.Size / mScaleHeight
            End Select
            
            If z Then .Controls(i).Left = .Controls(i).Left - 75000
            
        Next i
        frmMe.ScaleHeight = mSH
        frmMe.ScaleWidth = mSW
    End With
End Sub
Public Sub ResizeControls()
Dim i As Integer
Dim z As Boolean
On Error Resume Next
    
    For i = 0 To frmMe.Controls.Count - 1
        z = False
        If frmMe.Controls(i).Left < 0 Then
            frmMe.Controls(i).Left = frmMe.Controls(i).Left + 75000
            z = True
        End If
        With ProportionsArray(i)
            ' move and resize controls
            frmMe.Controls(i).Move .LeftProportions * frmMe.ScaleWidth, _
            .TopProportions * frmMe.ScaleHeight, _
            .WidthProportions * frmMe.ScaleWidth, _
            .HeightProportions * frmMe.ScaleHeight

            'Combo Boxes do not have a .Haight Property
            If Err.Number = 383 Then
                frmMe.Controls(i).Move .LeftProportions * frmMe.ScaleWidth, _
                .TopProportions * frmMe.ScaleHeight, _
                .WidthProportions * frmMe.ScaleWidth
            End If
            Err.Clear
            frmMe.Controls(i).Font.Size = .FontSize * frmMe.ScaleHeight
        End With
        If z Then frmMe.Controls(i).Left = frmMe.Controls(i).Left - 75000
    Next i
    If msFlex_Co Then rMe
End Sub

 

Private Sub Class_Terminate()
    Erase ProportionsArray
    Erase msFlex
    msFlex_Co = 0
End Sub


Private Sub frmMe_Resize()
    ResizeControls
End Sub
Private Sub rMe()
Dim i As Long
Dim j As Long

    For i = 0 To msFlex_Co - 1
        With arCell(i)
            For j = 0 To .ColCo
                msFlex(i).ColWidth(j) = arCell(i).ColWidth(j) * frmMe.ScaleWidth
            Next j
            For j = 0 To .RowCo
                msFlex(i).RowHeight(j) = arCell(i).RowHeight(j) * frmMe.ScaleHeight
            Next j
        End With
    Next i
End Sub

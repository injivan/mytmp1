VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDF_Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const ModulIdString As String = " - clsPDF_Parser "

Private cFilePDF As cFile



Public Function OpenPDF(ByVal fName As String) As Boolean
Const nFunction = 3004
On Error GoTo ErH
10
    
    If Not FileExists(fName) Then GoTo ErH
    Set cFilePDF = New cFile
    
    If Not cFilePDF.OpenFile(fName, 1) Then GoTo ehr
    OpenPDF = True
    
Exit Function
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
    Set cFilePDF = Nothing
End Function

 
Private Sub Class_Terminate()
    Set cFilePDF = Nothing
End Sub


White Spase chars
    0,9,10,12,13, 32, (13 & 10)


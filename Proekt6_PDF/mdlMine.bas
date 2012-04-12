Attribute VB_Name = "mdlMine"
Option Explicit

  Sub main()
Dim a As String

  Dim lngLocation As Long
  Dim strLine     As String
  
  Open "test.pdf" For Binary As #1   ' Open file just created.
  Open "pdfTest.txt" For Output As #2
  
  Do While lngLocation < LOF(1)   ' Loop until end of file.
     strLine = Input(1, #1)    ' Read character into variable.
     lngLocation = Loc(1)   ' Get current position within file.
     
    If strLine = vbLf Then
        strLine = Input(1, #1)    ' Read character into variable.
        lngLocation = Loc(1)   ' Get current position within file.
        strLine = vbNullString
        Print #2, a
        a = vbNullString
    End If
    a = a & strLine
     
  Loop
  Close #1   ' Close file.
  Print #2, a
  Close #2

  End

End Sub

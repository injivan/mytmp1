VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXmlElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text

'nFunction = 3000 -> GetElement

Private Const ModulIdString   As String = "c CXmlElement - "

Private m_colChildren As Collection     ' Child Elements of this Element
Private m_oTag        As CXmlTag               ' The tag defining this element, along
                                        ' with all it's associated properties
Public Body           As String                   ' The CharData section of the Element

Private mvarNodeCount As Long
Public Event ParsedBytes(ByVal NomByte As Long, ByRef fl_Exit As Boolean)

Public Sub AppendAtribute(ByVal Attrib As CXmlAttribute)
    ' Forward All Exceptions
    If m_oTag Is Nothing Then _
        Set m_oTag = New CXmlTag
    Call m_oTag.AppendAttribute(Attrib)
End Sub

Public Sub AppendNode(ByVal Element As CXmlElement, Optional ByVal Before As Long)
    'Forward All Exceptions
    If Before = 0 Then
        Call m_colChildren.Add(Element)
    Else
        Call m_colChildren.Add(Element, , Before)
    End If
    mvarNodeCount = mvarNodeCount + 1
End Sub
Public Property Get GetElement(ByVal Name As String, ByRef Element As CXmlElement) As Boolean
Const nFunction = 3000
On Error GoTo ErH

    Dim e  As CXmlElement
    
    For Each e In m_colChildren
        If UCase$(e.Name) = UCase$(Name) Then
            Set Element = e
            GetElement = True
            Exit For
        End If
    Next
    If Not GetElement Then
        '1010 = "Липсва указател"
        Err.Number = 1000
        Err.Description = " Липсва указател: < " & Name & " > "
    End If

Exit Property
ErH:
ShowErrMesage Err, ModulIdString, nFunction, Erl
Err.Clear
End Property
Public Function SetElement(ByVal Name As String, _
                           ByVal Data As Variant) As Boolean
    Dim e  As CXmlElement
    For Each e In m_colChildren
        If UCase$(e.Name) = UCase$(Name) Then
            e.Body = Data
            SetElement = True
            Exit For
        End If
    Next
    
End Function



Public Property Get AttributeCount() As Long
    If m_oTag Is Nothing Then Exit Property
    AttributeCount = m_oTag.AttributeCount
End Property

Public Property Get ElementAttribute(ByVal Index As Variant) As CXmlAttribute
    ' Forward All Excpetions
    If m_oTag Is Nothing Then Exit Property
    Set ElementAttribute = m_oTag.TagAttribute(Index)
End Property

Public Property Let Name(ByVal RHS As String)
    ' Todo . . . XML Validate RHS Data
    If m_oTag Is Nothing Then _
        Set m_oTag = New CXmlTag
    m_oTag.TagType = RHS
End Property

Public Property Get Name() As String
    If m_oTag Is Nothing Then Exit Property
    Name = m_oTag.TagType
End Property

Public Property Get NodeCount() As Long
    NodeCount = mvarNodeCount
End Property

Public Property Get NewEnum() As IUnknown
    Set NewEnum = m_colChildren.[_NewEnum]
End Property

Public Property Get Node(ByVal Index As Long) As CXmlElement
    ' Forward All Exceptions
    Set Node = m_colChildren(Index)
End Property
Public Function NodeSet(ByVal Index As Long, ByVal Data As CXmlElement) As Boolean
    ' Forward All Exceptions
      m_colChildren(Index) = Data
    NodeSet = True
End Function

Friend Function Parse(Data() As Byte, _
                 Optional Start As Long, _
                 Optional flNew As Boolean = True, _
                 Optional ByRef fl_DoExit As Boolean = False, _
                 Optional ByRef StartPoz As Long = 0) As Boolean
    
Dim lPos As Long, bStart As Boolean
Dim oTag As CXmlTag
Dim oChild As CXmlElement

Dim ub As Long

On Error GoTo Err_Trap
    
    Parse = True
    If Start < 0 Then Exit Function
    If Start = 0 Then Start = LBound(Data)
    
    If flNew Then
        ' Ensure we clear out the Current Data (if any)
        Set m_colChildren = New Collection
    End If
    
    ' Start Parsing the Element Body
    
    ub = UBound(Data)
    Do
        If ub < Start Then
            If lPos Then
                Start = Start - lPos
                lPos = Start
            Else
                Start = 0
            End If
            If Not GetNewData(lPos, Data, StartPoz) Then GoTo Err_Trap
            ub = UBound(Data)
            lPos = 0
        End If
        
        Select Case Data(Start)
            ' The first things we check for are the Start Tag Character '<'
            ' and the Amperstand '&'. These are the only characters not allowed
            ' in a character data section of an XML element. Also the sequence ']]>'
            ' is not allowed as it marks the end of a '<[CDATA[' markup section.
            Case ascTagBegin
                ' I am going to check if the last characters of the body are white
                ' space and trim them if they are.
                If bStart Then
                    Do
                        If Len(Body) = 0 Then Exit Do
                        lPos = Len(Body)
                        Select Case Asc(Right(Body, 1))
                            Case ascSpace, ascTab, ascCr, ascLf
                                Body = Left(Body, lPos - 1)
                        End Select
                        If Len(Body) = 0 Then Exit Do
                    Loop Until lPos = Len(Body) 'until we don't change the expression
                    lPos = 0
                    Body = Trim$(Body)
                Else
                    bStart = True
                End If
                
                ' Save the current seek value
                lPos = Start
                ' Parse the Tag By Itself
                Set oTag = New CXmlTag
                If Not oTag.Parse(Data(), Start, lPos, StartPoz) Then GoTo Err_Trap
                ub = UBound(Data)
                
                ' Start by getting the sTag for this element
                If m_oTag Is Nothing Then
                    Set m_oTag = oTag
                    
                    ' Technically, the First Tag could be an Empty Element Tag
                    If m_oTag.TagStyle = EmptyTag Then
                        Exit Function
                    ElseIf m_oTag.TagStyle = ETag Then
                        Set m_oTag = Nothing
                        ' I suppose it would be possible for an End Tag to exist without
                        ' an associated Start Tag, but it wouldn't be well formed if it was.
                    End If
                
                ' Check to see what Kind of Tag it is
                ElseIf oTag.TagStyle = ETag Then     ' We found our End Tag . . . Hurray!!!!
                    Exit Do
                
                ' New Tag . . . Time to Recurse
                Else
                    'Reset the Seek Value
                    Start = lPos
                    ' Create a Child Element
                    Set oChild = New CXmlElement
                    ' Tell it to Parse
                    Call oChild.Parse(Data, Start, , fl_DoExit, StartPoz)
                    ' Append it to the Collection
                    Call AppendNode(oChild)
                    ub = UBound(Data)
                    lPos = 0
                End If
                
                
            ' Escape Sequence
            Case ascAmper
                If bStart Then Body = Body & DecodeEscape(Data(), Start)
                
            Case Else
                ' Everything Else is CharData
                If bStart Then
                    If Data(Start) = 127 Then
                        Body = Body & "<"
                    ElseIf Data(Start) = 128 Then
                        Body = Body & ">"
                    Else
                        Body = Body & Chr$(Data(Start))
                    End If
                End If

        End Select
        Start = Start + 1
        If ((StartPoz + Start) Mod 100) = 0 Then
            RaiseEvent ParsedBytes(StartPoz + Start, fl_DoExit)
            DoEvents
        End If
        If fl_DoExit Then GoTo Err_Trap
    Loop
    
Exit Function
Err_Trap:
    Parse = False
    fl_DoExit = True
    Select Case Err.Number
        ' Exceptions Raised:
        Case 9
            'Unexpected End of Data [array index out of bounds]
            Err.Raise vbObjectError Or Err.Number, "Parse", "Missing closing tag </" & m_oTag.TagType & ">", vbNullString, 0
            
        ' Exceptions Forwarded:
        Case vbObjectError Or 9, vbObjectError Or 457
            'Tag Parse Exceptions
            'Element Parse Exceptions
            Err.Raise Err.Number
        
        Case Else
         
        ' ToDo . . . Log all other Errors
    End Select
End Function
Public Sub RemoveAttribute(ByVal Index As Variant)
    If m_oTag Is Nothing Then Exit Sub
    Call m_oTag.RemoveAttribute(Index)
End Sub

Public Sub RemoveNode(ByVal Index As Long)
    Call m_colChildren.Remove(Index)
    mvarNodeCount = mvarNodeCount - 1
End Sub

Public Function Serialize(Optional ByVal i As Integer = 0) As String
    Dim oElement As CXmlElement
    Dim lPos     As Long
    Dim sTab     As String
    
    Dim Y        As Long
    Dim sBod1    As String
    
    i = i + 1
    If i > 0 Then sTab = String$(i, Chr$(9))
    
    If m_oTag Is Nothing Then Exit Function
    
    Serialize = m_oTag.Serialize
    
    If m_colChildren.Count Then Serialize = Serialize & vbCrLf
    For Each oElement In m_colChildren
        Serialize = Serialize & sTab & oElement.Serialize(i)
    Next
    ' ToDo ... XML Encode Body Data
    lPos = InStr(Body, "]]>")
    If lPos > 0 Then
        Body = Left(Body, lPos - 1) & "]]&gt;" & Right(Body, Len(Body) - (lPos - 1) - 3)
    End If
    If Body = vbNullString Then
        sTab = String$(i - 1, Chr$(9))
        Serialize = Serialize & sTab & m_oTag.Serialize(ETag) & vbCrLf
    Else
        sBod1 = Body
        'Да маскирам елементите "<" и ">" от бодито
        '"<" = chr(127)
        '">" = chr(128)
        Y = 1
        Do While Y
            Y = InStr(Y, sBod1, "<")
            If Y Then
                Mid$(sBod1, Y, 1) = Chr(127)
                Y = Y + 1
            End If
        Loop
        Y = 1
        Do While Y
            Y = InStr(Y, sBod1, ">")
            If Y Then
                Mid$(sBod1, Y, 1) = Chr(128)
                Y = Y + 1
            End If
        Loop
    
    
        Serialize = Serialize & sBod1 & m_oTag.Serialize(ETag) & vbCrLf
    End If

End Function

Private Sub Class_Initialize()
    Set m_colChildren = New Collection
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXmlTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum TagStyles
    sTag        ' Start Tag
    ETag        ' End Tag
    EmptyTag    ' Empty Element Tag
End Enum

Private m_eStyle As TagStyles
Private m_colAttributes As Collection
Private m_sType As String


Friend Function Serialize(Optional eType As TagStyles = sTag) As String
    Dim oAttrib As CXmlAttribute
    
    If (eType = sTag) Or (eType = EmptyTag) Then
        Serialize = "<" & m_sType
        For Each oAttrib In m_colAttributes
            Serialize = Serialize & " " & oAttrib.Serialize
        Next
        Serialize = Serialize & IIf(eType = EmptyTag, "/>", ">")
    Else
        Serialize = "</" & m_sType & ">"
    End If
    
End Function

Public Sub AppendAttribute(ByVal Attrib As CXmlAttribute)
    On Error GoTo Err_Trap
    
    Call m_colAttributes.Add(Attrib, Attrib.KeyWord)
Exit Sub

Err_Trap:
    Select Case Err.Number
        ' Exceptions Raised:
        Case 457
            'Duplicate Attribute in Tag [add attribute with same key]
            Call Err.Raise(vbObjectError Or Err.Number, _
                           "AppendAttribute", _
                           "Attempt to Append Duplicate Attribute")
        
        ' Exceptions Forwarded:
        Case 91
            'Object Variable not Set
            Err.Raise Err.Number
            
        Case Else
        ' ToDo . . . Log All Other Errors
    End Select
End Sub
Public Property Get AttributeCount() As Long
    AttributeCount = m_colAttributes.Count
End Property

Public Property Get NewEnum() As IUnknown
    Set NewEnum = m_colAttributes.[_NewEnum]
End Property

Public Sub RemoveAttribute(ByVal Index As Variant)
    ' Forward All Exceptions
    Call m_colAttributes.Remove(Index)
End Sub

Public Property Get TagAttribute(ByVal Index As Variant) As CXmlAttribute
    ' Forward All Exceptions
    Set TagAttribute = m_colAttributes(Index)
End Property

' Start is the position of the Tag Start Element
' Upon Return it should be positioned on the end tag on the Tag Parsed
Friend Function Parse(Data() As Byte, Start As Long, lp As Long, NewFirst As Long) As Boolean
Dim oAttrib As CXmlAttribute
Dim ub As Long
    
Dim flFindBegin    As Boolean
Dim flRead_B_Tag   As Boolean
Dim flRead_E_Tag   As Boolean
Dim flFindNextTag  As Boolean 'Търсене на следващ таг
Dim s_B_Tag        As String
Dim s_E_Tag        As String
Dim i              As Long

    On Error GoTo Err_Trap
StartNewTag:
    Parse = True
    Start = Start + 1
    ub = UBound(Data)
    Do
        If ub < Start Then
            If lp Then
                Start = Start - lp
                lp = Start
            Else
                Start = 0
            End If
            If Not GetNewData(lp, Data, NewFirst) Then GoTo Err_Trap
            ub = UBound(Data)
            lp = 0
        End If
        
        Select Case Data(Start)
            Case ascTagTerm
                If Data(Start - 1) = ascTagBegin Then
                    m_eStyle = ETag
                ElseIf Data(Start + 1) = ascTagEnd Then
                    m_eStyle = EmptyTag
                End If
        
            ' Delimiting Characters
            Case ascSpace, ascTab, ascLf, ascCr
                ' Ignored
                
            ' Valid Name Starting Characetrs
            Case ascUpperFirst To ascUpperLast, _
                 ascLowerFirst To ascLowerLast, _
                 ascNumFirst To ascNumLast, _
                 ascUnderScore, ascColon, 92
                 
                If m_sType <> vbNullString Then
                    Set oAttrib = New CXmlAttribute
                    Call oAttrib.Parse(Data(), Start, lp)
                    ub = UBound(Data)
                    Call AppendAttribute(oAttrib)
                    Start = Start - 1
                Else
                    If Not ParseName(Data(), Start, lp, m_sType) Then GoTo Err_Trap
                    ub = UBound(Data)
                    Start = Start - 1
                End If
                
            Case ascTagEnd
                'Ако таг е поле
'''                If m_sType = sTAG_SECURITY Then
'''                    If m_eStyle = sTag Then
'''                        flFastRead = False
'''                    ElseIf m_eStyle = ETag Then
'''                        flFastRead = flTMPFastRead
'''                    End If
'''                End If
               
'''''                If flFastRead Then
'''''
'''''                    If m_sType = sTAG_POLE Or _
'''''                       m_sType = sTAG_LAYOUTS Or _
'''''                       m_sType = sTAG_BAZI Then
'''''                        i = 0
'''''                        'Намерен е таг POLE
'''''                        'Започвам да търся край на този таг
'''''                        lp = 0
'''''                        Do
'''''                            If ub < Start Then
'''''                                If lp Then
'''''                                    Start = Start - lp
'''''                                    lp = Start
'''''                                Else
'''''                                    Start = 0
'''''                                End If
'''''                                If Not GetNewData(lp, Data, NewFirst) Then GoTo Err_Trap
'''''                                ub = UBound(Data)
'''''                                lp = 0
'''''                            End If
'''''                            Select Case Data(Start)
'''''                            Case ascTagBegin
'''''                                flFindBegin = True  'Започва нов таг "<"
'''''                                flRead_B_Tag = True 'Чете започващ таг
'''''                                s_B_Tag = vbNullString
'''''                                If flFindNextTag Then
'''''                                    flFindNextTag = False
'''''                                    'Намерен е за
'''''                                    lp = Start
'''''                                    m_sType = vbNullString
'''''                                    GoTo StartNewTag
'''''                                End If
'''''                            Case ascTagTerm
'''''                                flRead_B_Tag = False
'''''                                flRead_E_Tag = True 'Чете завършващ таг
'''''                                If flFindBegin Then s_E_Tag = vbNullString
'''''                            Case ascTagEnd
'''''                                'Намерен е край на таг
'''''                                'Дали е края на нашия таг???
'''''                                If flRead_B_Tag Then
'''''                                    If m_sType = s_B_Tag Then
'''''                                        'Намерен е още 1 начален таг "POSITION"
'''''                                        i = i + 1
'''''                                    End If
'''''                                End If
'''''                                If flRead_E_Tag Then
'''''                                    If m_sType = s_E_Tag Then
'''''                                        If i = 0 Then flFindNextTag = True
'''''                                        i = i - 1: If i < 0 Then i = 0
'''''                                    End If
'''''                                End If
'''''                                flRead_B_Tag = False
'''''                                flRead_E_Tag = False
'''''                                s_B_Tag = vbNullString
'''''                                s_E_Tag = vbNullString
'''''                            Case Else
'''''                                If flRead_B_Tag Then s_B_Tag = s_B_Tag & Chr$(Data(Start))
'''''                                If flRead_E_Tag Then s_E_Tag = s_E_Tag & Chr$(Data(Start))
'''''
'''''                            End Select
'''''                            Start = Start + 1
'''''                        Loop
'''''
'''''                    End If
'''''                End If
'''''
                Exit Do
            
            Case Else
                ' Error Condition
                
        End Select
        Start = Start + 1
    Loop    ' Until we reach the End of the Tag
Exit Function

Err_Trap:
    Parse = False
    Select Case Err.Number
        ' Exceptions Raised:
        Case 9
            'Unexpected End of Data [array index out of bounds]
            Err.Raise vbObjectError Or Err.Number, "Parse", "Unexpected end of data", vbNullString, 0
            
        ' Exceptions Forwarded:
        Case vbObjectError Or 457, vbObjectError Or 9
            'CXmlAttribute.Parse Exceptions
            'ParseName Exceptions
            'AppendAttribute Exceptions
            Err.Raise Err.Number
        
        Case Else
        ' ToDo . . . Log All Other Errors
    
    End Select
End Function

Public Property Get TagStyle() As TagStyles
    TagStyle = m_eStyle
End Property

Public Property Let TagType(ByVal RHS As String)
    ' ToDo . . . XML Validate RHS Data
    m_sType = RHS
End Property

Public Property Get TagType() As String
    TagType = m_sType
End Property

Private Sub Class_Initialize()
    Set m_colAttributes = New Collection
End Sub

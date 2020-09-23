VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'################################################################
' XML Entity
' Created 27 June 2000 by James Vincent Carnicelli
'################################################################

Option Explicit


'##################### Private Declarations #####################

'Property mirrors
Private propName As String
Private propAttributes As XmlAttributes
Private propText As String
Private propChild As Collection


'####################### Public Properties ######################

Public Property Get Name() As String
    Name = propName
End Property
Public Property Let Name(newName As String)
    propName = newName
    propText = ""
End Property

Public Property Get Attributes() As XmlAttributes
    Set Attributes = propAttributes
End Property

Public Property Get Text() As String
    Text = propText
End Property
Public Property Let Text(newText As String)
    propText = newText
    propName = ""
    Set propAttributes = New XmlAttributes
    Set propChild = New Collection
End Property

Public Property Get Child(Index As Long) As XmlElement
Attribute Child.VB_UserMemId = 0
    Set Child = propChild(Index)
End Property

Public Property Get ChildCount() As Long
    ChildCount = propChild.Count
End Property


'######################## Public Methods ########################

Public Function CreateChild(Optional BeforeIndex As Long = 0) As XmlElement
    Dim NewChild As XmlElement
    Set NewChild = New XmlElement
    If BeforeIndex < 1 Then
        propChild.Add NewChild
    Else
        propChild.Add NewChild, , BeforeIndex
    End If
    Set CreateChild = NewChild
End Function

Public Sub RemoveChild(Index As Long)
    propChild.Remove Index
End Sub

Public Function IsTag() As Boolean
    IsTag = (propName <> "")
End Function

Public Function IsText() As Boolean
    IsText = (propName = "")
End Function

Public Function ToXml() As String
    Dim Xml As String, i As Long, Key As String, InnerXml As String
    If IsTag Then
        Xml = "<" & propName
        For i = 1 To propAttributes.Count
            Key = propAttributes.Key(i)
            Xml = Xml & " " & Key & "=""" & Encode(propAttributes(Key)) & """"
        Next
        For i = 1 To propChild.Count
            InnerXml = InnerXml & propChild(i).ToXml
        Next
        If InnerXml = "" Then
            Xml = Xml & "/>"
        Else
            Xml = Xml & ">" & InnerXml & "</" & propName & ">"
        End If
    Else
        Xml = Encode(propText)
    End If
    ToXml = Xml
End Function

'Each row contains:
'- Tag name
'- Text to put at start of element
'- Text to replace guts of element with or Null
'- Text to put at end of element
Public Function ToPlainText(Optional TagReplacements) As String
    Dim PlainText As String, i As Long, Key As String, InnerPlainText As String
    Dim TagReplacement
    If IsMissing(TagReplacements) Then
        If IsTag Then
            For i = 1 To propChild.Count
                PlainText = PlainText & propChild(i).ToPlainText
            Next
        Else
            PlainText = Encode(propText)
        End If
        ToPlainText = PlainText
    Else
        If IsTag Then
            For i = 0 To UBound(TagReplacements, 2)
                If TagReplacements(0, i) = propName Then
                    TagReplacement = Array( _
                      TagReplacements(0, i), _
                      TagReplacements(1, i), _
                      TagReplacements(2, i), _
                      TagReplacements(3, i) _
                    )
                    Exit For
                End If
            Next
            If IsEmpty(TagReplacement) Then
                For i = 1 To propChild.Count
                    PlainText = PlainText & propChild(i).ToPlainText(TagReplacements)
                Next
            Else
                PlainText = "" & TagReplacement(1)
                If IsNull(TagReplacement(2)) Then
                    For i = 1 To propChild.Count
                        PlainText = PlainText & propChild(i).ToPlainText(TagReplacements)
                    Next
                Else
                    PlainText = PlainText & TagReplacement(2)
                End If
                PlainText = PlainText & TagReplacement(3)
            End If
        Else
            PlainText = Encode(propText)
        End If
        ToPlainText = PlainText
    End If
End Function

Public Function Encode(ByVal PlainText As String) As String
    Encode = Replace(PlainText, "&", "&amp;")
    Encode = Replace(Encode, "<", "&lt;")
    Encode = Replace(Encode, ">", "&gt;")
    Encode = Replace(Encode, """", "&quot;")
End Function

Public Function Decode(ByVal XmlText As String) As String
    Decode = Replace(XmlText, "&quot;", """")
    Decode = Replace(Decode, "&gt;", ">")
    Decode = Replace(Decode, "&lt;", "<")
    Decode = Replace(Decode, "&amp;", "&")
End Function

Public Sub Clear()
    propName = ""
    Set propAttributes = New XmlAttributes
    Set propChild = New Collection
    propText = ""
End Sub

Public Sub CloneFrom(Element As XmlElement)
    Dim i As Long, Key As String
    Clear
    If Element.IsTag Then
        propName = Element.Name
        For i = 1 To Element.Attributes.Count
            Key = Element.Attributes.Key(i)
            propAttributes(Key) = Element.Attributes(Key)
        Next
        For i = 1 To Element.ChildCount
            CreateChild
            Child(i).CloneFrom Element.Child(i)
        Next
    Else
        propText = Element.Text
    End If
End Sub


'######################## Private Methods #######################


'#################### Private Event Handlers ####################

Private Sub Class_Initialize()
    Set propAttributes = New XmlAttributes
    Set propChild = New Collection
End Sub

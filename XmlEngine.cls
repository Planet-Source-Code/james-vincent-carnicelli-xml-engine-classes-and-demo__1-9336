VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'################################################################
' XML Parsing and Generating Engine
' Created 27 June 2000 by James Vincent Carnicelli
'################################################################

Option Explicit


'##################### Private Declarations #####################

Private FoundFirstTag As Boolean
Private ExtraElement As XmlElement
Private TreeBuildStack As Collection

'Property mirrors
Private propRootElement As XmlElement
Private propBuffer As String
Private propBuildTreeDuringParse As Boolean


'###################### Public Declarations #####################

Public Event FoundStartTag(TagName As String, Attributes As XmlAttributes)
Public Event FoundEndTag(TagName As String)
Public Event FoundContents(Text As String)


'####################### Public Properties ######################

Public Property Get RootElement() As XmlElement
    Set RootElement = propRootElement
End Property

Public Property Get Buffer() As String
    Buffer = propBuffer
End Property

Public Property Get BuildTreeDuringParse() As Boolean
    BuildTreeDuringParse = propBuildTreeDuringParse
End Property
Public Property Let BuildTreeDuringParse(newBuildTreeDuringParse As Boolean)
    propBuildTreeDuringParse = newBuildTreeDuringParse
End Property


'######################## Public Methods ########################

Public Sub InitializeBeforeParsing()
    CleanupAfterParsing
    propRootElement.Clear
End Sub

Public Sub CleanupAfterParsing()
    propBuffer = ""
    Set TreeBuildStack = New Collection
    FoundFirstTag = False
End Sub

Public Sub AppendAndParse(XmlFragment As String)
    Dim Name As String, Attributes As XmlAttributes
    Dim IsStart As Boolean, IsEnd As Boolean, Contents As String
    Dim CurrentElement As XmlElement, NewElement As XmlElement
    propBuffer = propBuffer & XmlFragment
    While FindTag(Name, Attributes, Contents, IsStart, IsEnd)
        
        '---- Automatically build tree ----
        If propBuildTreeDuringParse Then
            If TreeBuildStack.Count = 0 Then
                Set CurrentElement = RootElement
            Else
                Set CurrentElement = TreeBuildStack(TreeBuildStack.Count)
            End If
            
            If FoundFirstTag Then
                Set NewElement = CurrentElement.CreateChild
                NewElement.Text = Contents
                If IsStart Then
                    Set NewElement = CurrentElement.CreateChild
                    NewElement.Name = Name
                    NewElement.Attributes.CloneFrom Attributes
                End If
            Else
                Set NewElement = CurrentElement
                NewElement.Name = Name
                NewElement.Attributes.CloneFrom Attributes
                FoundFirstTag = True
            End If
            
            If IsStart Then
                TreeBuildStack.Add NewElement
            End If
            If IsEnd Then
                TreeBuildStack.Remove TreeBuildStack.Count
            End If
        
        '---- Let owner object handle parse events ----
        Else
            If FoundFirstTag Then
                RaiseEvent FoundContents(Contents)
            Else
                FoundFirstTag = True
            End If
            If IsStart Then RaiseEvent FoundStartTag(Name, Attributes)
            If IsEnd Then RaiseEvent FoundEndTag(Name)
        
        End If
    Wend
End Sub


'######################## Private Methods #######################

Private Function FindTag(Name As String, Attributes As XmlAttributes, Contents As String, IsStart As Boolean, IsEnd As Boolean) As Boolean
    Dim StartPos As Long, EndPos As Long, i As Long, i2 As Long
    Dim Key As String, Value As String
    StartPos = InStr(1, propBuffer, "<")  'No opening bracket
    If StartPos = 0 Then Exit Function
    i = PassWhitespace(StartPos + 1)
    If Mid(propBuffer, i, 1) = "/" Then
        i = i + 1
        IsStart = False
        IsEnd = True
    Else
        IsStart = True
    End If
    i = PassWhitespace(i)
    i = FindId(Name, i)
    If i = 0 Then Exit Function  'No tag name yet
    Set Attributes = New XmlAttributes
    Do
        i = PassWhitespace(i)
        i2 = FindId(Key, i)
        If i2 = 0 Then Exit Do
        i2 = PassWhitespace(i2)
        If Mid(propBuffer, i2, 1) <> "=" Then Exit Function
        i2 = PassWhitespace(i2 + 1)
        i2 = FindAttributeValue(Value, i2)
        If i2 = 0 Then Exit Do
        Attributes(Key) = Value
        i = i2
    Loop
    i = PassWhitespace(i)
    If IsStart Then
        If Mid(propBuffer, i, 1) = "/" Then
            i = i + 1
            IsEnd = True
        Else
            IsEnd = False
        End If
    End If
    EndPos = PassWhitespace(i)
    If Mid(propBuffer, EndPos, 1) <> ">" Then Exit Function
    Contents = Left(propBuffer, StartPos - 1)
    propBuffer = Mid(propBuffer, EndPos + 1)
    FindTag = True
End Function

Private Function PassWhitespace(StartAt As Long) As Long
    Dim i As Long, Char As String
    PassWhitespace = StartAt
    Do
        If PassWhitespace > Len(propBuffer) Then Exit Function
        Char = Mid(propBuffer, PassWhitespace, 1)
        If InStr(1, " " & vbTab & vbCrLf, Char) = 0 Then Exit Do
        PassWhitespace = PassWhitespace + 1
    Loop
End Function

Private Function FindId(Id As String, ByVal StartAt As Long) As Long
    Dim i As Long, Char As String
    i = StartAt
    Do
        If i > Len(propBuffer) Then Exit Do
        Char = UCase(Mid(propBuffer, i, 1))
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890_", Char) = 0 Then Exit Do
        i = i + 1
    Loop
    If i = StartAt Then Exit Function
    Id = Mid(propBuffer, StartAt, i - StartAt)
    FindId = i
End Function

Private Function FindAttributeValue(Value As String, StartAt As Long) As Long
    If Mid(propBuffer, StartAt, 1) <> """" Then Exit Function
    FindAttributeValue = InStr(StartAt + 1, propBuffer, """")
    If FindAttributeValue = 0 Then Exit Function
    FindAttributeValue = FindAttributeValue + 1
    Value = ExtraElement.Decode(Mid(propBuffer, StartAt + 1, FindAttributeValue - StartAt - 2))
End Function


'#################### Private Event Handlers ####################

Private Sub Class_Initialize()
    Set propRootElement = New XmlElement
    Set ExtraElement = New XmlElement
    Set TreeBuildStack = New Collection
End Sub

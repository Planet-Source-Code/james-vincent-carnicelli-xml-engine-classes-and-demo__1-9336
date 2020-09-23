VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'################################################################
' XML Entity Attributes
' Created 27 June 2000 by James Vincent Carnicelli
'################################################################

Option Explicit


'##################### Private Declarations #####################

'Property mirrors
Private propKeys As Collection
Private propValues As Collection


'####################### Public Properties ######################

Public Property Get Keys() As Collection
    Dim i As Integer, Key As String
    Set Keys = New Collection
    For i = 1 To propKeys.Count
        Key = propKeys(i)
        Keys.Add Key
    Next
End Property

Public Property Get Key(Index As Long) As String
    Key = propKeys(Index)
End Property

Public Property Get Item(ByVal Key As String) As String
Attribute Item.VB_UserMemId = 0
    Dim Value As String
    If Not Exists(Key, Value) Then Value = ""
    Item = Value
End Property
Public Property Let Item(ByVal Key As String, ByVal Value As String)
    If Exists(Key) Then Remove Key
    propKeys.Add Key, "i" & Key
    propValues.Add Value, "i" & Key
End Property


'######################## Public Methods ########################

Public Function Exists(ByVal Key As String, Optional Value) As Boolean
    On Error Resume Next
    Value = propValues.Item("i" & Key)
    Exists = (Err.Number = 0)
End Function

Public Function Count() As Long
    Count = propValues.Count
End Function

Public Sub Remove(ByVal Key As String)
    On Error Resume Next
    propKeys.Remove "i" & Key
    propValues.Remove "i" & Key
End Sub

Public Sub Clear()
    Set propKeys = New Collection
    Set propValues = New Collection
End Sub

Public Sub CloneFrom(Attributes As XmlAttributes)
    Dim i As Long, Key As String
    Clear
    For i = 1 To Attributes.Count
        Key = Attributes.Key(i)
        Item(Key) = Attributes(Key)
    Next
End Sub


'#################### Private Event Handlers ####################

Private Sub Class_Initialize()
    Clear
End Sub

VERSION 5.00
Begin VB.Form Demo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox DemoList 
      Height          =   1230
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   7155
   End
   Begin VB.TextBox Output 
      Height          =   3915
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1560
      Width           =   7155
   End
   Begin VB.Label Label 
      Caption         =   "Demo:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1515
   End
End
Attribute VB_Name = "Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents XE As XmlEngine
Attribute XE.VB_VarHelpID = -1

Private Indent As String

'Here are the actual demonstrations
Private Sub DemoList_Click()
    Dim ERoot As XmlElement, E1 As XmlElement, E2 As XmlElement
    Dim EClone As XmlElement, TagReplacement() As Variant
    Output.Text = ""
    Select Case DemoList.List(DemoList.ListIndex)
        
        Case "Actively parse an XML stream"
            Set XE = New XmlEngine
            XE.InitializeBeforeParsing
            XE.AppendAndParse "<XM"
            'Imaging this is a gap in time
            XE.AppendAndParse "L>This demonstrates how a <BOLD SIZE=""Huge"">"
            'Imaging this is a gap in time
            XE.AppendAndParse "broken stream</BOLD> of XML can "
            'Imaging this is a gap in time
            XE.AppendAndParse "be processed by your program as "
            'Imaging this is a gap in time
            XE.AppendAndParse "it is received, as from Winsock</XML>"
            XE.CleanupAfterParsing
            Output.SelText = vbCrLf & "Try stepping through this particular demo with SHIFT-F8"
        
        Case "Passively parse an XML stream"
            Set XE = New XmlEngine
            XE.InitializeBeforeParsing
            XE.BuildTreeDuringParse = True
            XE.AppendAndParse "<XM"
            'Imaging this is a gap in time
            XE.AppendAndParse "L>This demonstrates how a <BOLD SIZE=""Huge"">"
            'Imaging this is a gap in time
            XE.AppendAndParse "broken stream</BOLD> of XML can "
            'Imaging this is a gap in time
            XE.AppendAndParse "be processed by the class "
            'Imaging this is a gap in time
            XE.AppendAndParse "into an easy-to-use data structure</XML>"
            XE.CleanupAfterParsing
            Output.SelText = XE.RootElement.ToXml
        
        Case "Construct tree and export to text"
            Set XE = New XmlEngine
            Set ERoot = XE.RootElement
            ERoot.Name = "XmlDemo"
            ERoot.Attributes("Version") = "1.0"
            Set E1 = ERoot.CreateChild
              E1.Text = "This demonstrates how to "
            Set E1 = ERoot.CreateChild
              E1.Name = "Bold"
              Set E2 = E1.CreateChild
                E2.Text = "create"
            Set E1 = ERoot.CreateChild
              E1.Text = " a data structure ""on the fly"""
            Set E1 = ERoot.CreateChild(1)
              E1.Text = "Hi there!  "
            Output.SelText = ERoot.ToXml
        
        Case "Clone an XML tree"
            Set ERoot = New XmlElement
            ERoot.Name = "NOTICE"
            Set E1 = ERoot.CreateChild
              E1.Text = "This tree will be "
            Set E1 = ERoot.CreateChild
              E1.Name = "WOW"
              Set E2 = E1.CreateChild
                E2.Text = "cloned and the clone modified"
            Set E1 = ERoot.CreateChild
              E1.Text = "!"
            Output.SelText = ERoot.ToXml & vbCrLf & vbCrLf
            
            Set EClone = New XmlElement
            EClone.CloneFrom ERoot
            EClone(1).Text = "This tree has been "
            Output.SelText = EClone.ToXml
        
        Case "Export tree to XML and plain text"
            Set XE = New XmlEngine
            XE.InitializeBeforeParsing
            XE.BuildTreeDuringParse = True
            XE.AppendAndParse "<XML>This <B>tree <I>will</I> be exported</B> to XML and plain text</XML>"
            XE.CleanupAfterParsing
            Output.SelText = XE.RootElement.ToXml & vbCrLf & vbCrLf
            Output.SelText = XE.RootElement.ToPlainText & vbCrLf & vbCrLf
            
            'Now try it with tag replacement rules
            ReDim TagReplacement(3, 1)
            
            TagReplacement(0, 0) = "B"
            TagReplacement(1, 0) = "^"
            TagReplacement(2, 0) = Null
            TagReplacement(3, 0) = "^"
            
            TagReplacement(0, 1) = "I"
            TagReplacement(1, 1) = ""
            TagReplacement(2, 1) = "(some italic junk was here)"
            TagReplacement(3, 1) = ""
            
            Output.SelText = XE.RootElement.ToPlainText(TagReplacement)
        
    End Select
End Sub

Private Sub Form_Load()
    DemoList.AddItem "Actively parse an XML stream"
    DemoList.AddItem "Passively parse an XML stream"
    DemoList.AddItem "Construct tree and export to text"
    DemoList.AddItem "Clone an XML tree"
    DemoList.AddItem "Export tree to XML and plain text"
    
    Output.Text = "Select a demo.  Review the code in DemoList_Click()"
End Sub

Private Sub XE_FoundContents(Text As String)
    Output.SelText = Indent & "- """ & Text & """" & vbCrLf
End Sub

Private Sub XE_FoundEndTag(TagName As String)
    Indent = Left(Indent, Len(Indent) - 8)
    Output.SelText = Indent & "- </" & TagName & ">" & vbCrLf
End Sub

Private Sub XE_FoundStartTag(TagName As String, Attributes As XmlAttributes)
    Dim Key
    Output.SelText = Indent & "- <" & TagName
    Indent = Indent & "        "
    For Each Key In Attributes.Keys
        Output.SelText = vbCrLf & Indent & "  - " & Key & " = """ & Attributes(Key) & """" & vbCrLf
    Next
    If Attributes.Count > 0 Then
        Output.SelText = Indent & ">" & vbCrLf
    Else
        Output.SelText = ">" & vbCrLf
    End If
End Sub

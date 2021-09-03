VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COMPONENTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub Clear_Components()
    Dim LO As ListObject
    Set LO = COMPONENTS.ListObjects(COMPONENTS.CodeName)
    LO.AutoFilter.ShowAllData
    While LO.ListRows.Count > 0
        LO.ListRows(1).Delete
    Wend
End Sub

Sub Parse_Components(ByVal Node As MSXML2.IXMLDOMNode)

    Dim Obj As VBA.Collection
    Set Obj = COMP_SET(Node)
    
    Dim LO As Excel.ListObject, _
        LR As Excel.ListRow
    Set LO = COMPONENTS.ListObjects(COMPONENTS.CodeName)
    
    Dim O As Scripting.Dictionary, v As Variant, i As Integer
    For Each O In Obj
        For Each v In O("RULES")
            Set LR = LO.ListRows.Add
            i = 0
            i = i + 1: LR.Range(i) = O("CALENDAR")
            i = i + 1: LR.Range(i) = O("NAME")
            i = i + 1: LR.Range(i) = O("DESCRIPTION")
            i = i + 1: LR.Range(i) = O("EFFECTIVE_START_DATE")
            i = i + 1: LR.Range(i) = O("EFFECTIVE_END_DATE")
            i = i + 1: LR.Range(i) = v("NAME")
            i = i + 1: LR.Range(i) = v("TYPE")
            i = i + 1: LR.Range(i) = v("EFFECTIVE_START_DATE")
            i = i + 1: LR.Range(i) = v("EFFECTIVE_END_DATE")
        Next v
    Next O
    
End Sub

Private Function COMP_SET(ByVal Node As MSXML2.IXMLDOMNode) As VBA.Collection
    
    Set COMP_SET = New VBA.Collection
    
    Dim N As MSXML2.IXMLDOMNode
    For Each N In Node.ChildNodes
        COMP_SET.Add COMP(N)
    Next N
    
End Function

Private Function COMP(ByVal Node As MSXML2.IXMLDOMNode) As Scripting.Dictionary
    
    Set COMP = New Scripting.Dictionary
    
    Dim Attr As MSXML2.IXMLDOMAttribute
    For Each Attr In Node.Attributes
        COMP.Add Attr.Name, Attr.Text
    Next Attr
    
    Dim N As MSXML2.IXMLDOMNode
    Dim RULES As New VBA.Collection, RULE As Scripting.Dictionary
    For Each N In Node.ChildNodes
        Select Case N.nodeName
            Case "RULE_REF"
                Set RULE = New Scripting.Dictionary
                For Each Attr In N.Attributes
                    RULE.Add Attr.Name, Attr.Text
                Next Attr
                RULES.Add RULE
            Case Else
                Debug.Print N.nodeName
                Stop
        End Select
    Next N
    COMP.Add "RULES", RULES
    
End Function

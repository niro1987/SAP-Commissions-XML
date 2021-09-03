VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VARIABLES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub Clear_Variables()
    Dim LO As ListObject
    Set LO = VARIABLES.ListObjects(VARIABLES.CodeName)
    LO.AutoFilter.ShowAllData
    While LO.ListRows.Count > 0
        LO.ListRows(1).Delete
    Wend
End Sub

Sub Parse_Variables(ByVal Node As MSXML2.IXMLDOMNode)

    Dim Obj As VBA.Collection
    Set Obj = VARIABLE_SET(Node)
    
    Dim LO As Excel.ListObject, _
        LR As Excel.ListRow
    Set LO = VARIABLES.ListObjects(VARIABLES.CodeName)
    
    Dim O As Scripting.Dictionary, v As Variant, i As Integer
    For Each O In Obj
        Set LR = LO.ListRows.Add
        i = 0
        i = i + 1: LR.Range(i) = O("CALENDAR")
        i = i + 1: LR.Range(i) = O("NAME")
        i = i + 1: LR.Range(i) = O("DESCRIPTION")
        i = i + 1: LR.Range(i) = O("EFFECTIVE_START_DATE")
        i = i + 1: LR.Range(i) = O("EFFECTIVE_END_DATE")
        i = i + 1: LR.Range(i) = O("RETURN_TYPE")
        i = i + 1: LR.Range(i) = O("VARIABLE_TYPE")
        i = i + 1: LR.Range(i) = O("DEFAULT_ELEMENT_NAME")
    Next O
    
End Sub

Private Function VARIABLE_SET(ByVal Node As MSXML2.IXMLDOMNode) As VBA.Collection
    
    Set VARIABLE_SET = New VBA.Collection
    
    Dim N As MSXML2.IXMLDOMNode
    For Each N In Node.ChildNodes
        VARIABLE_SET.Add VARIABLE(N)
    Next N
    
End Function

Private Function VARIABLE(ByVal Node As MSXML2.IXMLDOMNode) As Scripting.Dictionary
    
    Dim N As MSXML2.IXMLDOMNode, Attr As MSXML2.IXMLDOMAttribute, i As Integer
    
    Set VARIABLE = New Scripting.Dictionary
    For Each Attr In Node.Attributes
        VARIABLE.Add Attr.Name, Attr.Text
    Next Attr
    
End Function

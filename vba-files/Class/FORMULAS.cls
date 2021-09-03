VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FORMULAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub Clear_Formulas()
    Dim LO As ListObject
    Set LO = FORMULAS.ListObjects(FORMULAS.CodeName)
    LO.AutoFilter.ShowAllData
    While LO.ListRows.Count > 0
        LO.ListRows(1).Delete
    Wend
End Sub

Sub Parse_Formulas(ByVal Node As MSXML2.IXMLDOMNode)

    Dim Obj As VBA.Collection
    Set Obj = FORMULA_SET(Node)
    
    Dim LO As Excel.ListObject, _
        LR As Excel.ListRow
    Set LO = FORMULAS.ListObjects(FORMULAS.CodeName)
    
    Dim O As Scripting.Dictionary, v As Variant, i As Integer
    For Each O In Obj
        Set LR = LO.ListRows.Add
        i = 0
        i = i + 1: LR.Range(i) = O("CALENDAR")
        i = i + 1: LR.Range(i) = O("NAME")
        i = i + 1: LR.Range(i) = O("DESCRIPTION")
        i = i + 1: LR.Range(i) = O("EFFECTIVE_START_DATE")
        i = i + 1: LR.Range(i) = O("EFFECTIVE_END_DATE")
        i = i + 1: LR.Range(i) = O("EXPRESSION")
        i = i + 1: LR.Range(i) = O("RETURN_TYPE")
    Next O
    
End Sub

Private Function FORMULA_SET(ByVal Node As MSXML2.IXMLDOMNode) As VBA.Collection
    
    Set FORMULA_SET = New VBA.Collection
    
    Dim N As MSXML2.IXMLDOMNode
    For Each N In Node.ChildNodes
        FORMULA_SET.Add FORMULA(N)
    Next N
    
End Function

Private Function FORMULA(ByVal Node As MSXML2.IXMLDOMNode) As Scripting.Dictionary
    
    Dim N As MSXML2.IXMLDOMNode, Attr As MSXML2.IXMLDOMAttribute, i As Integer
    
    Set FORMULA = New Scripting.Dictionary
    For Each Attr In Node.Attributes
        FORMULA.Add Attr.Name, Attr.Text
    Next Attr
    
    Set N = Node.FirstChild.FirstChild
    FORMULA.Add "EXPRESSION", mdFunctions.Parse_Function(N)
    
End Function

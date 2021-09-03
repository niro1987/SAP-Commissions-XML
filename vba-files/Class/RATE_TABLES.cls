VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RATE_TABLES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub Clear_RateTables()
    Dim LO As ListObject
    Set LO = RATE_TABLES.ListObjects(RATE_TABLES.CodeName)
    LO.AutoFilter.ShowAllData
    While LO.ListRows.Count > 0
        LO.ListRows(1).Delete
    Wend
End Sub

Sub Parse_RateTables(ByVal Node As MSXML2.IXMLDOMNode)

    Dim Obj As VBA.Collection
    Set Obj = RATE_TABLE_SET(Node)
    
    Dim LO As Excel.ListObject, _
        LR As Excel.ListRow
    Set LO = RATE_TABLES.ListObjects(RATE_TABLES.CodeName)
    
    Dim O As Scripting.Dictionary, v As Variant, i As Integer
    For Each O In Obj
        For Each v In O("SELECTORS")
            Set LR = LO.ListRows.Add
            i = 0
            i = i + 1: LR.Range(i) = O("CALENDAR")
            i = i + 1: LR.Range(i) = O("NAME")
            i = i + 1: LR.Range(i) = O("DESCRIPTION")
            i = i + 1: LR.Range(i) = O("EFFECTIVE_START_DATE")
            i = i + 1: LR.Range(i) = O("EFFECTIVE_END_DATE")
            i = i + 1: LR.Range(i) = O("RATE_EXP_UNIT_TYPE")
            i = i + 1: LR.Range(i) = O("RETURN_TYPE")
            i = i + 1: LR.Range(i) = O("SELECTOR_UNIT_TYPE")
            i = i + 1: LR.Range(i) = v("START_VALUE_INCLUSIVE")
            i = i + 1: LR.Range(i) = v("START_VALUE")
            i = i + 1: LR.Range(i) = v("START_VALUE_UNIT_TYPE")
            i = i + 1: LR.Range(i) = v("END_VALUE_INCLUSIVE")
            i = i + 1: LR.Range(i) = v("END_VALUE")
            i = i + 1: LR.Range(i) = v("END_VALUE_UNIT_TYPE")
            i = i + 1: LR.Range(i) = v("VALUE")
            i = i + 1: LR.Range(i) = v("VALUE_UNIT_TYPE")
        Next v
    Next O
    
End Sub

Private Function RATE_TABLE_SET(ByVal Node As MSXML2.IXMLDOMNode) As VBA.Collection
    
    Set RATE_TABLE_SET = New VBA.Collection
    
    Dim N As MSXML2.IXMLDOMNode
    For Each N In Node.ChildNodes
        RATE_TABLE_SET.Add RATE_TABLE(N)
    Next N
    
End Function

Private Function RATE_TABLE(ByVal Node As MSXML2.IXMLDOMNode) As Scripting.Dictionary
    
    Dim N As MSXML2.IXMLDOMNode, Attr As MSXML2.IXMLDOMAttribute, i As Integer
    
    Set RATE_TABLE = New Scripting.Dictionary
    For Each Attr In Node.Attributes
        RATE_TABLE.Add Attr.Name, Attr.Text
    Next Attr
    
    Dim S As MSXML2.IXMLDOMNode, Selectors As VBA.Collection, Selector As Scripting.Dictionary
    Set Selectors = New VBA.Collection
    For Each N In Node.ChildNodes
        Set Selector = New Scripting.Dictionary
        For Each S In N.ChildNodes
            If S.nodeName = "START_VALUE" Or S.nodeName = "END_VALUE" Then
                Selector.Add S.nodeName, S.Attributes(0).Text
                Selector.Add S.nodeName & "_INCLUSIVE", S.Attributes(1).Text
                Selector.Add S.nodeName & "_UNIT_TYPE", S.Attributes(2).Text
            Else
                Selector.Add S.nodeName, S.Attributes(0).Text
                Selector.Add S.nodeName & "_UNIT_TYPE", S.Attributes(1).Text
            End If
        Next S
        Selectors.Add Selector
    Next N
    RATE_TABLE.Add "SELECTORS", Selectors
    
End Function



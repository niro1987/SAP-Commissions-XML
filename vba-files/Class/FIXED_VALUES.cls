VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FIXED_VALUES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub Clear_FixedValues()
    Dim LO As ListObject
    Set LO = FIXED_VALUES.ListObjects(FIXED_VALUES.CodeName)
    LO.AutoFilter.ShowAllData
    While LO.ListRows.Count > 0
        LO.ListRows(1).Delete
    Wend
End Sub

Sub Parse_FixedValues(ByVal Node As MSXML2.IXMLDOMNode)

    Dim Obj As VBA.Collection
    Set Obj = FIXED_VALUE_SET(Node)
    
    Dim LO As Excel.ListObject, _
        LR As Excel.ListRow
    Set LO = FIXED_VALUES.ListObjects(FIXED_VALUES.CodeName)
    
    Dim O As Scripting.Dictionary, v As Variant, i As Integer
    For Each O In Obj
        Set LR = LO.ListRows.Add
        i = 0
        i = i + 1: LR.Range(i) = O("CALENDAR")
        i = i + 1: LR.Range(i) = O("NAME")
        i = i + 1: LR.Range(i) = O("DESCRIPTION")
        i = i + 1: LR.Range(i) = O("EFFECTIVE_START_DATE")
        i = i + 1: LR.Range(i) = O("EFFECTIVE_END_DATE")
        i = i + 1: LR.Range(i) = O("DECIMAL_VALUE")
        i = i + 1: LR.Range(i) = O("UNIT_TYPE")
    Next O
    
End Sub

Private Function FIXED_VALUE_SET(ByVal Node As MSXML2.IXMLDOMNode) As VBA.Collection
    
    Set FIXED_VALUE_SET = New VBA.Collection
    
    Dim N As MSXML2.IXMLDOMNode
    For Each N In Node.ChildNodes
        FIXED_VALUE_SET.Add FIXED_VALUE(N)
    Next N
    
End Function

Private Function FIXED_VALUE(ByVal Node As MSXML2.IXMLDOMNode) As Scripting.Dictionary
    
    Dim N As MSXML2.IXMLDOMNode, Attr As MSXML2.IXMLDOMAttribute, i As Integer
    
    Set FIXED_VALUE = New Scripting.Dictionary
    For Each Attr In Node.Attributes
        FIXED_VALUE.Add Attr.Name, Attr.Text
    Next Attr
    
End Function

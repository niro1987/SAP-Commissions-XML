VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LOOKUP_TABLES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub Clear_LookupTables()
    Dim LO As ListObject
    Set LO = LOOKUP_TABLES.ListObjects(LOOKUP_TABLES.CodeName)
    LO.AutoFilter.ShowAllData
    While LO.ListRows.Count > 0
        LO.ListRows(1).Delete
    Wend
End Sub

Sub Parse_LookupTables(ByVal Node As MSXML2.IXMLDOMNode)

    Dim Obj As VBA.Collection
    Set Obj = LOOKUP_TABLE_SET(Node)
    
    Dim LO As Excel.ListObject, _
        LR As Excel.ListRow
    Set LO = LOOKUP_TABLES.ListObjects(LOOKUP_TABLES.CodeName)
    
    Dim O As Scripting.Dictionary, v As Variant, i As Integer
    For Each O In Obj
        For Each v In O("DIMENSIONS")
            Set LR = LO.ListRows.Add
            i = 0
            i = i + 1: LR.Range(i) = O("CALENDAR")
            i = i + 1: LR.Range(i) = O("NAME")
            i = i + 1: LR.Range(i) = O("DESCRIPTION")
            i = i + 1: LR.Range(i) = O("EFFECTIVE_START_DATE")
            i = i + 1: LR.Range(i) = O("EFFECTIVE_END_DATE")
            i = i + 1: LR.Range(i) = O("UNIT_TYPE")
            i = i + 1: LR.Range(i) = O("CONVERT_NULL_VALUE")
            i = i + 1: LR.Range(i) = v("DIM_NAME")
            i = i + 1: LR.Range(i) = v("EFFECTIVE_START_DATE")
            i = i + 1: LR.Range(i) = v("EFFECTIVE_END_DATE")
            i = i + 1: LR.Range(i) = v("CASE_SENSITIVE")
            i = i + 1: LR.Range(i) = v("TYPE")
            i = i + 1: LR.Range(i) = v("VALUE")
            i = i + 1: LR.Range(i) = v("LOW_VALUE")
            i = i + 1: LR.Range(i) = v("HIGH_VALUE")
        Next v
    Next O
    
End Sub

Private Function LOOKUP_TABLE_SET(ByVal Node As MSXML2.IXMLDOMNode) As VBA.Collection
    
    Set LOOKUP_TABLE_SET = New VBA.Collection
    
    Dim N As MSXML2.IXMLDOMNode
    For Each N In Node.ChildNodes
        LOOKUP_TABLE_SET.Add LOOKUP_TABLE(N)
    Next N
    
End Function

Private Function LOOKUP_TABLE(ByVal Node As MSXML2.IXMLDOMNode) As Scripting.Dictionary
    
    Dim N As MSXML2.IXMLDOMNode, Attr As MSXML2.IXMLDOMAttribute, i As Integer
    
    Set LOOKUP_TABLE = New Scripting.Dictionary
    For Each Attr In Node.Attributes
        LOOKUP_TABLE.Add Attr.Name, Attr.Text
    Next Attr
    
    Set N = Node.ChildNodes(0)
    For Each Attr In N.Attributes
        LOOKUP_TABLE.Add Attr.Name, Attr.Text
    Next Attr
    
    Dim D As MSXML2.IXMLDOMNode, Dimensions As New VBA.Collection, Dimension As Scripting.Dictionary
    For Each N In Node.ChildNodes(0).ChildNodes
        If N.nodeName = "DIM_NAME" Then
            For Each D In N.ChildNodes
                Set Dimension = New Scripting.Dictionary
                For Each Attr In N.Attributes
                    Dimension.Add Attr.Name, Attr.Text
                Next Attr
                For Each Attr In D.Attributes
                    Dimension.Add Attr.Name, Attr.Text
                Next Attr
                Dimensions.Add Dimension
            Next D
        End If
    Next N
    LOOKUP_TABLE.Add "DIMENSIONS", Dimensions
    
End Function

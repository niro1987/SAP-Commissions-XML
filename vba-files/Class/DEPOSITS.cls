VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DEPOSITS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub Clear_Deposits()
    Dim LO As ListObject
    Set LO = DEPOSITS.ListObjects(DEPOSITS.CodeName)
    LO.AutoFilter.ShowAllData
    While LO.ListRows.Count > 0
        LO.ListRows(1).Delete
    Wend
End Sub

Sub Parse_Deposits(ByVal Node As MSXML2.IXMLDOMNode)
    
    Dim LO As Excel.ListObject, _
        LR As Excel.ListRow
    Set LO = DEPOSITS.ListObjects(DEPOSITS.CodeName)
    
    Dim O As Scripting.Dictionary
    Set O = RULE(Node)
    
    Dim v As Variant, i As Integer, y As Integer
    For Each v In O("OUTPUT")
        Set LR = LO.ListRows.Add
        i = 0
        i = i + 1: LR.Range(i) = O("CALENDAR")
        i = i + 1: LR.Range(i) = O("NAME")
        i = i + 1: LR.Range(i) = O("DESCRIPTION")
        i = i + 1: LR.Range(i) = O("EFFECTIVE_START_DATE")
        i = i + 1: LR.Range(i) = O("EFFECTIVE_END_DATE")
        i = i + 1: LR.Range(i) = O("CONDITION")
        i = i + 1: LR.Range(i) = O("TYPE")
        i = i + 1: LR.Range(i) = v("NAME")
        i = i + 1: LR.Range(i) = v("DISPLAY_NAME_FOR_REPORTS")
        i = i + 1: LR.Range(i) = v("IS_REPORTABLE")
        i = i + 1: LR.Range(i) = v("PERIOD_TYPE")
        i = i + 1: LR.Range(i) = v("TYPE")
        i = i + 1: LR.Range(i) = v("UNIT_TYPE")
        i = i + 1: LR.Range(i) = v("VALUE")
        i = i + 1: LR.Range(i) = v("HOLD_REF_NAME")
        i = i + 1: LR.Range(i) = v("HOLD_CONDITION")
        i = i + 1: LR.Range(i) = v("EARNING_GROUP")
        i = i + 1: LR.Range(i) = v("EARNING_CODE")
        
        y = 0
        For i = i + 1 To i + 16
            y = y + 1
            LR.Range(i) = v("GA" & y)
        Next i
        y = 0
        For i = i To i + 5
            y = y + 1
            LR.Range(i) = v("GN" & y)
        Next i
        y = 0
        For i = i To i + 5
            y = y + 1
            LR.Range(i) = v("GD" & y)
        Next i
        y = 0
        For i = i To i + 5
            y = y + 1
            LR.Range(i) = v("GB" & y)
        Next i
    Next v
    
End Sub

Private Function RULE(ByVal Node As MSXML2.IXMLDOMNode) As Scripting.Dictionary
    
    Set RULE = New Scripting.Dictionary
    
    Dim Attr As MSXML2.IXMLDOMAttribute
    For Each Attr In Node.Attributes
        RULE.Add Attr.Name, Attr.Text
    Next Attr
    
    Dim N As MSXML2.IXMLDOMNode
    Dim ACTION_EXPRESSION_SET As New VBA.Collection
    For Each N In Node.ChildNodes
        Select Case N.nodeName
            Case "ACTION_EXPRESSION_SET"
                Set ACTION_EXPRESSION_SET = ACTION_SET(N)
            Case "CONDITION_EXPRESSION"
                RULE.Add "CONDITION", mdFunctions.Parse_Function(N.FirstChild)
            Case Else
                Debug.Print N.nodeName & " is currently not supported."
        End Select
    Next N
    RULE.Add "OUTPUT", ACTION_EXPRESSION_SET
    
End Function

Private Function ACTION_SET(ByVal Node As MSXML2.IXMLDOMNode) As VBA.Collection
    
    Set ACTION_SET = New VBA.Collection
    
    Dim N As MSXML2.IXMLDOMNode
    For Each N In Node.ChildNodes
        ACTION_SET.Add ACTION(N.FirstChild)
    Next N
    
End Function

Private Function ACTION(ByVal Node As MSXML2.IXMLDOMNode) As Scripting.Dictionary
    
    Dim N As MSXML2.IXMLDOMNode, Attr As MSXML2.IXMLDOMAttribute, i As Integer
    Set ACTION = New Scripting.Dictionary
    
    Set N = Node.ChildNodes(0)      'OUTPUT_REFERENCE
    For Each Attr In N.Attributes
        ACTION.Add Attr.Name, Attr.Text
    Next Attr
    
    Set N = Node.ChildNodes(1)      'EARNING_CODE
    ACTION.Add "EARNING_CODE", mdFunctions.Parse_Function(N)
    
    Set N = Node.ChildNodes(2)      'EARNING_GROUP
    ACTION.Add "EARNING_GROUP", mdFunctions.Parse_Function(N)
    
    Set N = Node.ChildNodes(3)      'VALUE
    ACTION.Add "VALUE", mdFunctions.Parse_Function(N)
    
    Set N = Node.ChildNodes(4)      'HOLD
    For Each Attr In N.Attributes
        ACTION.Add "HOLD_REF_" & Attr.Name, Attr.Text
    Next Attr
    
    ' TODO: Parse HOLD_CONDITION Node
    ACTION.Add "HOLD_CONDITION", N.Text
    
    For i = 1 To 16
        Set N = Node.ChildNodes(i + 4)    'GA's 1 trough 16
        ACTION.Add "GA" & i, mdFunctions.Parse_Function(N)
    Next i
    
    For i = 1 To 6
        Set N = Node.ChildNodes(i + 20)  'GN's 1 trough 6
        ACTION.Add "GN" & i, mdFunctions.Parse_Function(N)
    Next i
    
    For i = 1 To 6
        Set N = Node.ChildNodes(i + 26)    'GD's 1 trough 6
        ACTION.Add "GD" & i, mdFunctions.Parse_Function(N)
    Next i
    
    For i = 1 To 6
        Set N = Node.ChildNodes(i + 32)    'GB's 1 trough 6
        ACTION.Add "GB" & i, mdFunctions.Parse_Function(N)
    Next i

End Function

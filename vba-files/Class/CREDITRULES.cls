VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CREDITRULES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub Clear_CreditRules()
    Dim LO As ListObject
    Set LO = CREDITRULES.ListObjects(CREDITRULES.CodeName)
    LO.AutoFilter.ShowAllData
    While LO.ListRows.Count > 0
        LO.ListRows(1).Delete
    Wend
End Sub

Sub Parse_CreditRules(ByVal Node As MSXML2.IXMLDOMNode)
    
    Dim LO As Excel.ListObject, _
        LR As Excel.ListRow
    Set LO = CREDITRULES.ListObjects(CREDITRULES.CodeName)
    
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
        i = i + 1: LR.Range(i) = O("EVENT_TYPE")
        i = i + 1: LR.Range(i) = O("TYPE")
        i = i + 1: LR.Range(i) = v("NAME")
        i = i + 1: LR.Range(i) = v("DISPLAY_NAME_FOR_REPORTS")
        i = i + 1: LR.Range(i) = v("IS_REPORTABLE")
        i = i + 1: LR.Range(i) = v("PERIOD_TYPE")
        i = i + 1: LR.Range(i) = v("TYPE")
        i = i + 1: LR.Range(i) = v("UNIT_TYPE")
        i = i + 1: LR.Range(i) = v("VALUE")
        i = i + 1: LR.Range(i) = v("HOLD_REF_NAME")
        i = i + 1: LR.Range(i) = v("HOLD_REF_PERIOD_TYPE")
        i = i + 1: LR.Range(i) = v("CONDITION")
        i = i + 1: LR.Range(i) = v("CREDIT_TYPE")
        i = i + 1: LR.Range(i) = v("DUPLICATE")
        i = i + 1: LR.Range(i) = v("ROLL_UP")
        i = i + 1: LR.Range(i) = v("ROLL_DATE")
        
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
            Case "EVENT_TYPE_EXPRESSION"
                RULE.Add "EVENT_TYPE", mdFunctions.Parse_Function(N.FirstChild)
            Case Else
                Debug.Print N.nodeName
                Stop
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
    
    Set N = Node.ChildNodes(1)      'VALUE
    ACTION.Add "VALUE", mdFunctions.Parse_Function(N)
    
    Set N = Node.ChildNodes(2)      'HOLD
    For Each Attr In N.Attributes
        ACTION.Add "HOLD_REF_" & Attr.Name, Attr.Text
    Next Attr
    
    Set N = Node.ChildNodes(3)      'HOLD_CONDITION
    ACTION.Add "CONDITION", mdFunctions.Parse_Function(N)
    
    Set N = Node.ChildNodes(4)      'CREDIT_TYPE
    ACTION.Add "CREDIT_TYPE", N.Text
    
    Set N = Node.ChildNodes(5)      'DUPLICATES
    ACTION.Add "DUPLICATE", mdFunctions.Parse_Function(N)
    
    Set N = Node.ChildNodes(6)      'ROLL_UP
    ACTION.Add "ROLL_UP", mdFunctions.Parse_Function(N)
    
    Set N = Node.ChildNodes(7)      'ROLL_DATE
    ACTION.Add "ROLL_DATE", mdFunctions.Parse_Function(N)
    
    For i = 1 To 16
        Set N = Node.ChildNodes(i + 7)    'GA's 1 trough 16
        ACTION.Add "GA" & i, mdFunctions.Parse_Function(N)
    Next i
    
    For i = 1 To 6
        Set N = Node.ChildNodes(i + 23)    'GN's 1 trough 6
        ACTION.Add "GN" & i, mdFunctions.Parse_Function(N)
    Next i
    
    For i = 1 To 6
        Set N = Node.ChildNodes(i + 29)    'GD's 1 trough 6
        ACTION.Add "GD" & i, mdFunctions.Parse_Function(N)
    Next i
    
    For i = 1 To 6
        Set N = Node.ChildNodes(i + 35)    'GB's 1 trough 6
        ACTION.Add "GB" & i, mdFunctions.Parse_Function(N)
    Next i
    
    
End Function

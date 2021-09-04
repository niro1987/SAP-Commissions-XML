VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PLANS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub Clear_Plans()
    Dim LO As ListObject
    Set LO = PLANS.ListObjects(PLANS.CodeName)
    LO.AutoFilter.ShowAllData
    While LO.ListRows.Count > 0
        LO.ListRows(1).Delete
    Wend
End Sub

Sub Parse_Plans(ByVal Node As MSXML2.IXMLDOMNode)

    Dim Obj As VBA.Collection
    Set Obj = PLAN_SET(Node)
    
    Dim LO As Excel.ListObject, _
        LR As Excel.ListRow
    Set LO = PLANS.ListObjects(PLANS.CodeName)
    
    Dim O As Scripting.Dictionary, v As Variant, i As Integer
    For Each O In Obj
        For Each v In O("COMPONENTS")
            Set LR = LO.ListRows.Add
            i = 0
            i = i + 1: LR.Range(i) = O("CALENDAR")
            i = i + 1: LR.Range(i) = O("NAME")
            i = i + 1: LR.Range(i) = O("DESCRIPTION")
            i = i + 1: LR.Range(i) = O("EFFECTIVE_START_DATE")
            i = i + 1: LR.Range(i) = O("EFFECTIVE_END_DATE")
            i = i + 1: LR.Range(i) = v("NAME")
            i = i + 1: LR.Range(i) = v("EFFECTIVE_START_DATE")
            i = i + 1: LR.Range(i) = v("EFFECTIVE_END_DATE")
        Next v
    Next O
    
End Sub

Private Function PLAN_SET(ByVal Node As MSXML2.IXMLDOMNode) As VBA.Collection
    
    Set PLAN_SET = New VBA.Collection
    
    Dim N As MSXML2.IXMLDOMNode
    For Each N In Node.ChildNodes
        PLAN_SET.Add PLAN(N)
    Next N
    
End Function

Private Function PLAN(ByVal Node As MSXML2.IXMLDOMNode) As Scripting.Dictionary
    
    Set PLAN = New Scripting.Dictionary
    
    Dim Attr As MSXML2.IXMLDOMAttribute
    For Each Attr In Node.Attributes
        PLAN.Add Attr.Name, Attr.Text
    Next Attr
    
    Dim N As MSXML2.IXMLDOMNode
    Dim COMPONENTS As New VBA.Collection, COMPONENT As Scripting.Dictionary
    For Each N In Node.ChildNodes
        Select Case N.nodeName
            Case "COMPONENT_REF"
                Set COMPONENT = New Scripting.Dictionary
                For Each Attr In N.Attributes
                    COMPONENT.Add Attr.Name, Attr.Text
                Next Attr
                COMPONENTS.Add COMPONENT
            Case "VARIABLE_ASSIGNMENT"
                Debug.Assert True ' TODO
            Case Else
                Debug.Print N.nodeName & " is currently not supported."
        End Select
    Next N
    PLAN.Add "COMPONENTS", COMPONENTS
    
End Function

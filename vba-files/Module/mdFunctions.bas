Attribute VB_Name = "mdFunctions"
Option Explicit

Function Parse_Function(ByVal Node As MSXML2.IXMLDOMNode) As String
    
    Dim Prefix As String, Suffix As String, Operator As String
    
    Select Case Node.nodeName
        Case _
            "UNIT_TYPE", _
            "CREDIT_TYPE", _
            "BOOLEAN", _
            "DATA_FIELD", _
            "PERIOD_TYPE"
                Parse_Function = Node.Text
        Case "MDLTVAR_REF"
            Parse_Function = Node.Attributes.getNamedItem("NAME").Text
        Case "RULE_ELEMENT_REF"
            Parse_Function = Node.Attributes.getNamedItem("NAME").Text
        Case "MEASUREMENT_REF"
            Parse_Function = Node.Attributes.getNamedItem("NAME").Text
            If Node.Attributes.getNamedItem("PERIOD_TYPE").Text <> Node.Attributes.getNamedItem("OUTPUT_REFERENCE_PERIOD_TYPE").Text Then
                Parse_Function = Parse_Function & "(" & Node.Attributes.getNamedItem("PERIOD_TYPE").Text & ")"
            End If
            If Node.Attributes.getNamedItem("PERIOD_OFFSET").Text <> "0" Then
                Parse_Function = Parse_Function & "-" & Node.Attributes.getNamedItem("PERIOD_OFFSET").Text
            End If
        Case "INCENTIVE_REF"
            Parse_Function = Node.Attributes.getNamedItem("NAME").Text
        Case "MDLT_REF"
            Parse_Function = Node.Attributes.getNamedItem("NAME").Text
        Case "FUNCTION"
            Parse_Function = F_Parse(Node)
        Case "OPERATOR"
            Parse_Function = O_Parse(Node)
        Case "STRING_LITERAL"
            If Node.Text = "NULL" Then
                Parse_Function = "-"
            Else
                Parse_Function = """" & Node.Text & """"
            End If
        Case "VALUE"
            Parse_Function = Node.Attributes(0).Text
        Case Else
            Debug.Print Node.nodeName & " is currently not supported."
    End Select
    
End Function

Private Function F_Parse(ByVal Node As MSXML2.IXMLDOMNode) As String
    
    Dim FuncName As String, FuncParts() As String
    FuncName = Node.Attributes.getNamedItem("ID").Text
    ReDim FuncParts(1 To Node.ChildNodes.Length)
    Dim i As Integer
    For i = 1 To Node.ChildNodes.Length
        FuncParts(i) = mdFunctions.Parse_Function(Node.ChildNodes(i - 1))
    Next i
    F_Parse = FuncName & "(" & Join(FuncParts, ", ") & ")"
    
End Function

Private Function O_Parse(ByVal Node As MSXML2.IXMLDOMNode) As String
    
    Dim Operator As String, Wrapped As Boolean
    
    Wrapped = Node.Attributes.Length = 2
    
    Select Case Node.Attributes.getNamedItem("ID").Text
        Case "ISEQUALTO_OPERATOR":          Operator = " = "
        Case "NOTEQUALTO_OPERATOR":         Operator = " <> "
        Case "AND_OPERATOR":                Operator = " AND "
        Case "OR_OPERATOR":                 Operator = " OR "
        Case "MULTIPLY_OPERATOR":           Operator = " * "
        Case "DIVISION_OPERATOR":           Operator = " / "
        Case "SUBTRACT_OPERATOR":           Operator = " - "
        Case "ADD_OPERATOR":                Operator = " + "
        Case "GREATERTHAN_OPERATOR":        Operator = " > "
        Case "LESSTHAN_OPERATOR":           Operator = " < "
        Case "GREATERTHANEQUALTO_OPERATOR": Operator = " >= "
        Case "LESSTHANEQUALTO_OPERATOR":    Operator = " <= "
        Case Else
            Debug.Print Node.Attributes.getNamedItem("ID").Text & " is currently not supported."
    End Select
    
    O_Parse = _
        mdFunctions.Parse_Function(Node.ChildNodes(0)) & _
        Operator & _
        mdFunctions.Parse_Function(Node.ChildNodes(1))
    
    If Wrapped Then
        O_Parse = "(" & O_Parse & ")"
    End If
    
End Function


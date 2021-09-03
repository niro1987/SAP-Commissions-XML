Attribute VB_Name = "mdInit"
' Creator:      Niels Perfors
' Github:       https://github.com/niro1987/SAP-Commissions-XML
' License:      SAP-Commissions-XML is licensed under the GNU General Public License v3.0

Option Explicit

Sub Select_Plan_File_Path()
    
    With Application.FileDialog(msoFileDialogFilePicker)
        
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Plan.xml", "*.xml", 1
        .InitialFileName = ThisWorkbook.Names("Plan_File_Path").RefersToRange.Text
        
        If .Show Then
            ThisWorkbook.Names("Plan_File_Path").RefersToRange = .SelectedItems(1)
        End If
        
    End With
    
End Sub

Sub Parse()
Attribute Parse.VB_ProcData.VB_Invoke_Func = "P\n14"
    
    Dim PlanFilePath As String
    PlanFilePath = ThisWorkbook.Names("Plan_File_Path").RefersToRange.Text
    
    Dim FSO As New Scripting.FileSystemObject
    Dim PlanFile As Scripting.TextStream
    Set PlanFile = FSO.OpenTextFile(PlanFilePath)
    
    Dim PlanXML As New MSXML2.DOMDocument60
    PlanXML.LoadXML PlanFile.ReadAll
    Set PlanFile = Nothing
    Set FSO = Nothing
    
    Dim PlanData As MSXML2.IXMLDOMNode
    Set PlanData = PlanXML.ChildNodes(1)
    Set PlanXML = Nothing
    
    'Wipe the slate clean
    PLANS.Clear_Plans
    COMPONENTS.Clear_Components
    CREDITRULES.Clear_CreditRules
    MEASUREMENTS.Clear_Measurements
    INCENTIVES.Clear_Incentives
    DEPOSITS.Clear_Deposits
    LOOKUP_TABLES.Clear_LookupTables
    RATE_TABLES.Clear_RateTables
    FIXED_VALUES.Clear_FixedValues
    VARIABLES.Clear_Variables
    FORMULAS.Clear_Formulas
    
    Dim N As MSXML2.IXMLDOMNode
    For Each N In PlanData.ChildNodes
        Parse_Node N
    Next N
    
End Sub

Sub Parse_Node(ByVal Node As MSXML2.IXMLDOMNode)
    
    Dim N As MSXML2.IXMLDOMNode
    
    Select Case Node.nodeName
        Case "PLAN_SET"
            PLANS.Parse_Plans Node
            
        Case "PLANCOMPONENT_SET"
            COMPONENTS.Parse_Components Node
            
        Case "RULE_SET"
            For Each N In Node.ChildNodes
                Select Case N.Attributes.getNamedItem("TYPE").Text
                    Case "DIRECT_TRANSACTION_CREDIT"
                        CREDITRULES.Parse_CreditRules N
                    Case "PRIMARY_MEASUREMENT", "SECONDARY_MEASUREMENT"
                        MEASUREMENTS.Parse_Measurements N
                    Case "BULK_COMMISSION"
                        INCENTIVES.Parse_Incentives N
                    Case "DEPOSIT"
                        DEPOSITS.Parse_Deposits N
                    Case Else
                        Debug.Print N.Attributes.getNamedItem("TYPE").Text
                End Select
            Next N
            
        Case "MD_LOOKUP_TABLE_SET"
            LOOKUP_TABLES.Parse_LookupTables Node
            
        Case "RATETABLE_SET"
            RATE_TABLES.Parse_RateTables Node
            
        Case "FIXED_VALUE_SET"
            FIXED_VALUES.Parse_FixedValues Node
            
        Case "VARIABLE_SET"
            VARIABLES.Parse_Variables Node
            
        Case "FORMULA_SET"
            FORMULAS.Parse_Formulas Node
        
        Case Else
            Debug.Print Node.nodeName & " is currently not supported."
            
    End Select
    
End Sub


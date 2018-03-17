Attribute VB_Name = "JE_Index"
Option Explicit
Option Private Module
Public I As Integer
Public JE_num As String
Public T_row As Integer
Public D_row As Integer
Public T As Integer
Public T_column As Integer
Public max_Row As Integer
Public A_range As Variant
Public Docum As String 'The value of document row
Public A_Name As Variant 'the value of the account object
Public Name As String  'The name of the form object
Public Caption As String 'the caption of the form object
Public H As Integer 'the hight of the form
Public L As Integer 'the left margng of the form
Public CmbBx As Object 'combobox object
Public oneControl As Object
Public C_list As Variant
Public F_V As Boolean
Public Txt_Lbl As Object 'TEXT box object
Public Cbx As Object 'check box for documents
Public Lb1 As Object 'this will create a label
Public M_range As Variant 'minimum doc
Public S_range As Variant  'sub category
Public D_range As Variant   'Documents
Public A_item As Variant 'Account Different
Public T_wbk As Workbook 'Name of the template workbook

Sub JE_Index()

    With Application
        .EnableEvents = False
        .DisplayAlerts = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
             
             JE_CheckList.Show
        
    With Application
        .EnableEvents = True
        .DisplayAlerts = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    
End Sub

Sub Account()

    Set T_wbk = ThisWorkbook
    
    T_wbk.Sheets("Template").Visible = True
    
    'Create a new sheet for calculation
    T_wbk.Sheets.Add.Name = "Calculation"
    
    'Now we will select the template sheet
    
    T_wbk.Sheets("Template").Select
    
    'Find the last non-blank cell in column A(1)
    T_row = T_wbk.Sheets("Template").Cells(Rows.Count, 1).End(xlUp).Row
    
    'Copy the value and paste in the target sheet without using Clipboard
    T_wbk.Sheets("Template").Range(Cells(2, 1), Cells(T_row, 1)).SpecialCells(xlCellTypeVisible).Copy T_wbk.Sheets("Calculation").Range("A1")
    
    'Deleting any duplicate value
    T_wbk.Sheets("Calculation").Columns("A:A").RemoveDuplicates Columns:=1, Header:=xlNo
    
    'Find the last non-blank cell in column A(1)
    T_row = T_wbk.Sheets("Calculation").Cells(Rows.Count, 1).End(xlUp).Row
    
    'Selecting the calculated sheet
    T_wbk.Sheets("Calculation").Select

    'Assigning the account to A_range variable
    A_range = T_wbk.Sheets("Calculation").Range(Cells(1, 1), Cells(T_row, 1))
    
    'Now we will delete the calculation sheet
    T_wbk.Sheets("Calculation").Delete
    
    'Now we will create a combo box name with label Account
    Caption = "Account"
    Name = "AL_1"
    H = 40
    L = 30
    
    Call JE_Module.create_Label
    
    Caption = "Account"
    Name = "AC_1"
    H = 40
    L = 96
    
    Call JE_Module.create_ComboBox
    
End Sub

Sub Minimum()

    'Now we will create a Calculated sheet
    T_wbk.Sheets.Add.Name = "Calculation"

    T_wbk.Sheets("Template").Select
    
    'Find the last non-blank cell in column A(1)
    T_row = T_wbk.Sheets("Template").Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    T_column = T_wbk.Sheets("Template").Cells(1, Columns.Count).End(xlToLeft).Column
    
    'Filter column A from criteria
    T_wbk.Sheets("Template").Range("$A$1:$D$" & T_row).AutoFilter Field:=1, Criteria1:=A_Name
    
    'Now we will copy only the Mimimum required cells from column B ans paste it to column A
    T_wbk.Sheets("Template").Range("$B$2:$B$" & T_row).SpecialCells(xlCellTypeVisible).Copy T_wbk.Sheets("Calculation").Range("A1")
    
    'Deleting any duplicate value
    T_wbk.Sheets("Calculation").Columns("A:A").RemoveDuplicates Columns:=1, Header:=xlNo
    
    'Now we will copy only the SubCategory cells from column C and paste it to B
    T_wbk.Sheets("Template").Range("$C$2:$C$" & T_row).SpecialCells(xlCellTypeVisible).Copy T_wbk.Sheets("Calculation").Range("B1")
    
    'Deleting any duplicate value
    T_wbk.Sheets("Calculation").Columns("B:B").RemoveDuplicates Columns:=1, Header:=xlNo

    'Now we will copy only the Documents Required cells from column D and paste it to C
    T_wbk.Sheets("Template").Range("$D$2:$D$" & T_row).SpecialCells(xlCellTypeVisible).Copy T_wbk.Sheets("Calculation").Range("C1")
    
    'Deleting any duplicate value
    T_wbk.Sheets("Calculation").Columns("C:C").RemoveDuplicates Columns:=1, Header:=xlNo
    
    'Removing the filter from the Template sheet
    T_wbk.Sheets("Template").Range("$A$1:$D$" & T_row).AutoFilter
    
    'Selecting the calculated sheet
    T_wbk.Sheets("Calculation").Select
    
    'Find the last non-blank cell in column A(1)
    T_row = T_wbk.Sheets("Calculation").Cells(Rows.Count, 1).End(xlUp).Row
    
    'Assigning the munimum document required into variant M_range
    M_range = T_wbk.Sheets("Calculation").Range(Cells(1, 1), Cells(T_row, 1))
    
    'Find the last non-blank cell in column B(1)
    T_row = T_wbk.Sheets("Calculation").Cells(Rows.Count, 2).End(xlUp).Row
    
    'Assigning the sub category required into variant S_range
    S_range = T_wbk.Sheets("Calculation").Range(Cells(1, 2), Cells(T_row, 2))
    
    'Find the last non-blank cell in column C(1)
    T_row = T_wbk.Sheets("Calculation").Cells(Rows.Count, 3).End(xlUp).Row
    
    'Assigning the document required into variant D_range
    D_range = T_wbk.Sheets("Calculation").Range(Cells(1, 3), Cells(T_row, 3))
    
    'Now we will delete the calculation sheet
    T_wbk.Sheets("Calculation").Delete
    
    'Now we need to check which array has max rows
    max_Row = Excel.WorksheetFunction.Max(UBound(M_range), UBound(D_range))
    
    If A_Name = "Accrued Revenue" Then
            Caption = "Sub - Category"
            Name = "SUb_CAT1"
            Call Dyna_Form.create_Label_heading
            Call Dyna_Form.sub_Cat_combo
        Else
            Caption = "Documents"
            Name = "Doc_R"
            Call Dyna_Form.create_Label_heading
            Call Dyna_Form.Doc_From_chkBox
    End If
    
    T_wbk.Sheets("Template").Visible = False
    
    Call Dyna_Form.create_Form_Label
    
    Call Dyna_Form.create_Form_txtBox
    
    
End Sub

Sub sub_Cat()

If A_Name = "Accrued Revenue" Then
        Select Case CmbBx.Value
            Case "T&M "
                Docum = "ETES report, SOW, LOE, Client confirmation"
            Case "Fixed Price (POC)"
                Docum = "Financial Plan and YTD Cost Dump of the Relevant Period with WBS focus; Approved Contribution margin Percentage, Contracts, EAC Templates (POC Base), RDF / RRCL"
            Case "Materials ODC"
                Docum = "Cost Dump, Approved Markup revenue percentage, Contract"
            Case "Fixed Price (Baseline / installment)"
                Docum = "Contract/Excerpts, Pricing extracts / schedules, Prior month invoice, Confirmation to accrue (not being billed in current period), RDF / RRCL"
            Case "License Revenue"
                Docum = "Confirmation of License Installation /Delivery Note"
        End Select
    End If

End Sub

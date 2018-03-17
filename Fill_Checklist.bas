Attribute VB_Name = "Fill_Checklist"
Option Explicit
Option Private Module
Sub fillChecklist()

    'Now we will select the CheckList sheet
    
    T_wbk.Sheets("CheckList").Select
    
    'Find the last non-blank cell in column B(1)
    T_row = T_wbk.Sheets("CheckList").Cells(Rows.Count, 3).End(xlUp).Row + 1
    
     'Find the last non-blank cell in column D(1)
    D_row = T_wbk.Sheets("CheckList").Cells(Rows.Count, 5).End(xlUp).Row + 1
    
    'Now we need to check which column has max rows
    max_Row = Excel.WorksheetFunction.Max(T_row, D_row)
    
    'entering the JE_Number
    T_wbk.Sheets("CheckList").Range("A" & max_Row).Value = JE_num
    
    'entering the Account_Number
    T_wbk.Sheets("CheckList").Range("B" & max_Row).Value = A_Name
    
    ' entering the minimum document to address the following
    I = 1
    T = 0
    For Each oneControl In JE_Details.Controls
        If TypeName(oneControl) = "TextBox" Then
            If oneControl.Name = "Min_Box" & I Then
                T_wbk.Sheets("CheckList").Range("C" & max_Row + T).Value = oneControl
                T = T + 1
            End If
                I = I + 1
        End If
    Next oneControl
    
    ' entering the Sub category
    
    For Each oneControl In JE_Details.Controls
        If TypeName(oneControl) = "ComboBox" Then
            If oneControl.Name = "Sub_Com_box" Then
                T_wbk.Sheets("CheckList").Range("D" & max_Row).Value = oneControl
            End If
        End If
    Next oneControl
    
      ' entering the Documents
    
    If A_Name = "Accrued Revenue" Then
            T_wbk.Sheets("CheckList").Range("E" & max_Row).Value = Docum
        Else
            I = 1
            T = 0
            ' entering the Documents
            For Each oneControl In JE_Details.Controls
                If TypeName(oneControl) = "CheckBox" Then
                    If oneControl.Name = "Doc_C" & I And oneControl = True Then
                        T_wbk.Sheets("CheckList").Range("E" & max_Row + T).Value = oneControl.Caption
                        T = T + 1
                    End If
                        I = I + 1
                End If
            Next oneControl
    End If
    
End Sub

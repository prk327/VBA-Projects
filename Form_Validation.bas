Attribute VB_Name = "Form_validation"
Option Explicit
Option Private Module

Sub JE_CheckList_Validation()

    If JE_num <> "" Then
        
        Else
            MsgBox "Please Enter JE_Number"
            JE_CheckList.Show
    End If
    
    F_V = False
    
    For Each oneControl In JE_CheckList.Controls
        If TypeName(oneControl) = "ComboBox" Then
            For Each C_list In A_range
                If oneControl = C_list And oneControl.Name = "AC_1" Then
                    F_V = True
                End If
            Next
            If F_V = False Then
                MsgBox "Please select One Account from the list"
                JE_CheckList.Show
            End If
        End If
    Next oneControl
    
    
End Sub

Sub JE_Details_Validation()

    'Checking the text box for values
    I = 1
    For Each oneControl In JE_Details.Controls
        If TypeName(oneControl) = "TextBox" Then
            If oneControl = "" And oneControl.Name = "Min_Box" & I Then
                MsgBox "Please enter all minimum document required!!"
                JE_Details.Show
            End If
            I = I + 1
        End If
    Next oneControl

    'Checking the check box for values
    
    F_V = False
    I = 1
    If A_Name = "Accrued Revenue" Then
            'Checking the combo box for values of sub category
            For Each oneControl In JE_Details.Controls
                If TypeName(oneControl) = "ComboBox" Then
                    For Each C_list In S_range
                        If oneControl = C_list And oneControl.Name = "Sub_Com_box" Then
                            F_V = True
                        End If
                    Next
                    If F_V = False Then
                        MsgBox "Please select One Sub_category from the list"
                        JE_Details.Show
                    End If
                End If
            Next oneControl
        Else
            'Checking the check box for values of documents
            For Each oneControl In JE_Details.Controls
                If TypeName(oneControl) = "CheckBox" Then
                    If oneControl = True And oneControl.Name = "Doc_C" & I Then
                        F_V = True
                    End If
                    I = I + 1
                End If
            Next oneControl
            If F_V = False Then
                MsgBox "Please check Minimum one document"
                JE_Details.Show
            End If
    End If
    
End Sub

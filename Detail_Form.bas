Attribute VB_Name = "Detail_Form"
Option Explicit
Option Private Module
Sub create_Label()
    
    'We can use Add method to add the new controls on run time
    Set Lb1 = JE_Details.Controls.Add("Forms.Label.1")
    With Lb1
        .Top = T
        .Left = L
        .Width = 156
        .Height = H
        .Caption = Caption
        .Name = Name
        .Font.Name = "Times New Roman"
        .Font.Bold = True
        .Font.Size = 10
        .AutoSize = True
    End With
End Sub

Sub create_ComboBox()
    'Add Dynamic Combo Box and assign it to object 'CmbBx'
    Set CmbBx = JE_Details.Controls.Add("Forms.comboBox.1")
    'Combo Box Position
    With CmbBx
        .Top = T
        .Left = L
        .Height = H
        .Width = 156
        .Name = Name
        .ColumnCount = UBound(A_range)
        .List = A_range
        .AutoSize = True
    End With
End Sub

Sub create_CheckBox()
    
    'Add Dynamic Checkbox and assign it to object 'Cbx'
    Set Cbx = JE_Details.Controls.Add("Forms.CheckBox.1")
    With Cbx
        .Top = T 'Checkbox Position
        .Left = L
        .Height = H
        .Width = 156
        .Caption = Caption 'Assign Checkbox Name
        .Name = Name
        .Font.Name = "Times New Roman"
        .Font.Bold = True
        .Font.Size = 10
'        .AutoSize = True
    End With
End Sub

Sub create_TextBox()

    Set Txt_Lbl = JE_Details.Controls.Add("Forms.TextBox.1")
    With Txt_Lbl
        .Top = T 'Checkbox Position
        .Left = L
        .Height = H
        .Width = 156
        .Name = Name
    End With
End Sub



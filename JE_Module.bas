Attribute VB_Name = "JE_Module"
Option Explicit
Option Private Module
Sub create_Label()
    
    'We can use Add method to add the new controls on run time
    Set Lb1 = JE_CheckList.Controls.Add("Forms.Label.1")
    With Lb1
        .Top = H
        .Left = L
        .Width = 150
        .Height = 18
        .Caption = Caption
        .Name = Name
        .Font.Name = "Times New Roman"
        .Font.Bold = True
        .Font.Size = 10
    End With
End Sub

Sub create_ComboBox()
    
    'Add Dynamic Combo Box and assign it to object 'CmbBx'
    Set CmbBx = JE_CheckList.Controls.Add("Forms.comboBox.1")
    'Combo Box Position
    With CmbBx
        .Top = H
        .Left = L
        .Width = 150
        .Height = 18
        .Name = Name
        .ColumnCount = UBound(A_range)
        .List = A_range
    End With
    
    
End Sub


Attribute VB_Name = "Dyna_Form"
Option Explicit
Option Private Module
Sub create_Form_Label()
    I = 1
    T = 25
    L = 30
    H = 40
    'will create label for minimum documentation
    For Each A_item In M_range
        If A_item <> "" Then
            Caption = A_item
            Name = "Min_L" & I
                Call Detail_Form.create_Label
            T = T + H
'            L = 66
            I = I + 1
        End If
    Next
End Sub

Sub create_Form_txtBox()
    I = 1
    T = 25
    L = 186
    H = 40
    'will create label for minimum documentation
    For Each A_item In M_range
        If A_item <> "" Then
            Name = "Min_Box" & I
                Call Detail_Form.create_TextBox
            T = T + H
            I = I + 1
        End If
    Next
End Sub

Sub sub_Cat_combo()
    'Add Dynamic Combo Box and assign it to object 'CmbBx'
    Set CmbBx = JE_Details.Controls.Add("Forms.comboBox.1")
    'Combo Box Position
    With CmbBx
        .Top = 25
        .Left = 372
        .Height = 40
        .Width = 156
        .Name = "Sub_Com_box"
        .ColumnCount = UBound(S_range)
        .List = S_range
    End With
End Sub

Sub Doc_From_chkBox()

    I = 1
    T = 25
    L = 372
    H = 40
    'will create label for minimum documentation
    For Each A_item In D_range
        If A_item <> "" Then
            Caption = A_item
            Name = "Doc_C" & I
                Call Detail_Form.create_CheckBox
            T = T + H
            I = I + 1
        End If
    Next
End Sub

Sub create_Label_heading()
    T = 10
    L = 372
    H = 12
    'will create label for heading documentation
    Caption = Caption
    Name = Name
    Call Detail_Form.create_Label
End Sub

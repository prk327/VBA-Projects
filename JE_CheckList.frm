VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} JE_CheckList 
   Caption         =   "UserForm1"
   ClientHeight    =   2328
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   3768
   OleObjectBlob   =   "JE_CheckList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "JE_CheckList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Me.Hide
    JE_num = JE_CheckList.JE_Num_Tbox.Value
    A_Name = CmbBx.Value
    Call Form_validation.JE_CheckList_Validation
    
    If JE_CheckList.JE_Num_Tbox.Value <> "" And CmbBx.Value <> "" Then
            Unload Me ' unloading the checklist form
            JE_Details.Show
            End
    End If
    
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

    Me.Caption = "JE CheckList"
'    Me.BackColor = RGB(10, 25, 100)
     Call JE_Index.Account
     
    Me.Height = 137
    Me.Width = 300
    
    Set CMD = JE_CheckList.CommandButton1
    
    With CMD
        .Top = 75
        .Left = 120
        .Width = 66
    End With

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

'    msg = "Now Unloading " & Me.Caption
'    MsgBox prompt:=msg, Title:="QueryClose Event"

End Sub

Private Sub UserForm_Terminate()

'    msg = "Now Unloading " & Me.Caption
'    MsgBox prompt:=msg, Title:="Terminate Event"

End Sub

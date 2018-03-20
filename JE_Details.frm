VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} JE_Details 
   Caption         =   "UserForm1"
   ClientHeight    =   2460
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10860
   OleObjectBlob   =   "JE_Details.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "JE_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_1_Click()
    Unload Me
    End
End Sub

Private Sub Submit_Click()
    Me.Hide
    
    'need to call the sub category
    Call JE_Index.sub_Cat
    
    'need to call the from validation
    Call Form_validation.JE_Details_Validation
    
    'need to call the fill checklist to fill the values in the sheet
    Call Fill_Checklist.fillChecklist
    
    'saving the file in the current location

    Unload Me
    
    End
    
End Sub

Private Sub UserForm_Initialize()

    Me.Caption = "JE_Details"
    
    Call JE_Index.Minimum
    
    Me.Height = (max_Row * 40) + 90
    Me.Width = 558
    
    With JE_Details.Submit
        .Height = 24
        .Width = 66
        .Top = (max_Row * 40) + 35
        .Left = 140
    End With
    
    With JE_Details.Close_1
        .Height = 24
        .Width = 66
        .Top = (max_Row * 40) + 35
        .Left = 418
    End With
    
    

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

   If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If

End Sub

Private Sub UserForm_Terminate()

'    msg = "Now Unloading " & Me.Caption
'    MsgBox prompt:=msg, Title:="Terminate Event"

End Sub


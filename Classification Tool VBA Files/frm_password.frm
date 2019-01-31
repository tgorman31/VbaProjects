VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_password 
   Caption         =   "Password Required"
   ClientHeight    =   1092
   ClientLeft      =   39
   ClientTop       =   390
   ClientWidth     =   3705
   OleObjectBlob   =   "frm_password.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_passsubmit_Click()
    If Me.tb_password = "CW" Then
            ThisWorkbook.Visible = True
            Worksheets("Lookups").Activate
            
            Unload Me
            Unload frm_Mast
        
        Else:  Unload Me
    End If
    
End Sub

Private Sub UserForm_Initialize()
    With frm_password
        .StartUpPosition = 0
        .Top = 250
        .Left = 200
        
    End With
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

        If CloseMode = 0 Then Cancel = True

End Sub

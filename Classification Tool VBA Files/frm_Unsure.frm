VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Unsure 
   Caption         =   "You have not answered one/some of the questions"
   ClientHeight    =   2717
   ClientLeft      =   39
   ClientTop       =   377
   ClientWidth     =   6227
   OleObjectBlob   =   "frm_Unsure.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Unsure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_GoBack_Click()

    myALog = 1
    Unload Me
End Sub

Private Sub cmd_Comment_Click()

    myALog = 2
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()

    With frm_Unsure
         .StartUpPosition = 0
         .Top = 250
         .Left = 250
    End With
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

        If CloseMode = 0 Then Cancel = True

End Sub

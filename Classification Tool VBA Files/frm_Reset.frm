VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Reset 
   Caption         =   "Are you sure you wish to reset?"
   ClientHeight    =   2808
   ClientLeft      =   39
   ClientTop       =   377
   ClientWidth     =   6097
   OleObjectBlob   =   "frm_Reset.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Reset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmd_warnNo_Click()
    Unload Me

End Sub

Private Sub cmd_warnyes_Click()
    Unload Me
    Range("DataAns38").Value = "No"
    Application.Run ("Module1.ADOFromExcelToAccess")
    Mail_noSubmit
    closLog = 0
    Unload frm_Mast
    frm_Mast.Show
End Sub

Private Sub UserForm_Initialize()

    With frm_Reset
         .StartUpPosition = 0
         .Top = 250
         .Left = 250
    End With
End Sub

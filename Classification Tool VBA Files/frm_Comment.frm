VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Comment 
   Caption         =   "Please provide comment for the below Question"
   ClientHeight    =   3978
   ClientLeft      =   39
   ClientTop       =   390
   ClientWidth     =   11596
   OleObjectBlob   =   "frm_Comment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Comment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Comment_Click()
Dim str As String
    
    If isOne = 0 Then
        If Me.Controls("opt_Low") = False And Me.Controls("opt_Medium") = False And Me.Controls("opt_High") = False Then
            MsgBox ("Please select one of the options")
            Exit Sub
        Else
            If Me.Controls("opt_Low") = True Then
                rskRating = lowRate
                str = "Yes - Low Impact - "
            ElseIf Me.Controls("opt_Medium") = True Then
                rskRating = medRate
                str = "Yes - Medium Impact - "
            ElseIf Me.Controls("opt_High") = True Then
                rskRating = hiRate
                str = "Yes - High Impact - "
            End If
        End If
    Else
        str = "Unsure - "
    End If
    
    Range(CCode2).Value = str & Me.tb_Comment
    myALog = 0
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
    
    With frm_Comment
        .StartUpPosition = 0
        .Top = 250
        .Left = 250
    End With
    
    Me.tb_Quest.Value = Range(QCode2).Value
    
    If isOne = 0 Then
        Me.fr_rate.Visible = True
    End If
        
    myALog = 1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

        If CloseMode = 0 Then Cancel = True

End Sub

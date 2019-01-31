VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim objExcel As Excel.Application

'Dim win As Window

Dim FileName As String

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    hideSheets
    Worksheets("Data").Protect ("CW")
End Sub

Public Sub W_open()

    objExcel.Workbooks.Open FileName:=FileName

End Sub
Private Sub Workbook_open()

    'Set win = ActiveWindow

    'ThisWorkbook.ChangeFileAccess Mode:=xlReadOnly
    
    frm_Mast.Show
    'win.Visible = False
    
End Sub
Private Sub fixER()

    'ThisWorkbook.ChangeFileAccess Mode:=xlReadWrite
    'ActiveWindow.Visible = True
    
End Sub



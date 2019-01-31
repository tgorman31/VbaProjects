Attribute VB_Name = "PDF_Print"
Sub PDF()
'Save page as PDF
Dim ws As Worksheet
Dim strPath As String
Dim myFile As Variant
Dim strFile As String
On Error GoTo errHandler

    Set ws = Worksheets("Report Sheet")
    
    ws.Visible = xlSheetVisible
    
    'enter name and select folder for file
    ' start in current workbook folder
    'Replace(nmCode, " ", "_")
    strFile = nmCode _
                & "-" _
                & "Classification Report-" _
                & Format(Now(), "ddmmmyyyy") _
                & ".pdf"
    strFile = ThisWorkbook.path & "\" & strFile
    
    myFile = Application.GetSaveAsFilename _
        (InitialFileName:=strFile, _
            FileFilter:="PDF Files (*.pdf), *.pdf", _
            Title:="Select Folder and FileName to save")
    
    If myFile <> "False" Then
        ws.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            FileName:=myFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    
        MsgBox "PDF file has been created."
        ws.Visible = xlSheetVeryHidden
    End If
        Application.Run ("RPT_Update.tbClear")
exitHandler:
        Exit Sub
errHandler:
        MsgBox "Could not create PDF file"
        Resume exitHandler
        

End Sub





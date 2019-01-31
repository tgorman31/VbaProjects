Attribute VB_Name = "RPT_Update"
Sub UpdateReportSheet()

Dim frm As Object
Dim fld, val, aCode, tb As String
Dim ws, wt As Worksheet
Dim tbl As Integer

    Set ws = Worksheets("Report Sheet")
    Set frm = frm_Mast


    For i = 1 To 11
        Select Case i
            Case 1
                fld = "RPT_SchoolName"
                val = "tb_SchoolName"
                tb = "TextBox 32"
            Case 2
                fld = "RPT_ProjectName"
                val = "tb_ProjectName"
                tb = "TextBox 33"
            Case 3
                fld = "RPT_ProjectNum"
                val = "tb_ProjetNo"
                tb = "TextBox 35"
            Case 4
                fld = "RPT_DelManager"
                val = "tb_DM"
                tb = "TextBox 37"
            Case 5
                fld = "RPT_ProjDur"
                val = "cmb_EstProjectDur"
                tb = "TextBox 40"
            Case 6
                fld = "RPT_ProjectVal"
                val = "cmb_EstProjectVal"
                tb = "TextBox 38"
            Case 7
                fld = "RPT_ProjectCat"
                val = "cmb_ProjCat"
                tb = "TextBox 34"
            Case 8
                fld = "RPT_DelMethod"
                val = "cmb_DelMethod"
                tb = "TextBox 36"
            Case 9
                fld = "RPT_Class"
                val = "txt_Overall"
                tb = "TextBox 42"
            Case 10
                fld = "RPT_ClassUser"
                val = "tb_YourName"
                tb = "TextBox 92"
            Case 11
                fld = "RPT_Date"
                val = "tb_date"
                tb = "TextBox 93"
        End Select
        
        ws.Range(fld).Value = frm.Controls(val)
                
        'ws.Shapes(tb).TextFrame.Characters.Text = frm.Controls(val)
    ActiveWorkbook.Sheets("Calc").Activate

    Next
    tbl = 1
    For i = 1 To 24
        aCode = Range("ANS" & i & "_").Value
        If aCode = "Yes" Then
                ws.Range("RPT_Info" & tbl).Value = Range("STAT" & i)
                tbl = tbl + 1
        End If
        
'        If tbl = 78 Then
'            tbl = tbl + 1
'        End If
    Next
    ActiveWorkbook.Sheets("Calc").Activate
    x = 2
    tbl = 21
    For Each cell In Range("CommentAll")
        If Not cell = "" Then
                ActiveWorkbook.Sheets("Report Sheet").Range("B" & tbl).Value = ActiveWorkbook.Sheets("Calc").Range("I" & x).Value
                tbl = tbl + 1
                ActiveWorkbook.Sheets("Report Sheet").Range("B" & tbl).Value = ActiveWorkbook.Sheets("Calc").Range("G" & x).Value
                tbl = tbl + 1
                x = x + 1
        Else
            x = x + 1
        End If
    Next
        
'        If tbl = 78 Then
'            tbl = tbl + 1
'        End If
'    Next
    
    Range("AnswerAll").Clear
    
End Sub

Sub tbClear()

Dim tbl As Integer
Dim aCode As String
Dim ws As Worksheet

    Set ws = Worksheets("Report Sheet")

    tbl = 1
    For i = 1 To 24
        aCode = Range("ANS" & i & "_").Value
        If aCode = "No" Then
                ws.Range("RPT_Info" & tbl).Value = ""
                tbl = tbl + 1
        End If
        
'        If tbl = 78 Then
'            tbl = tbl + 1
'        End If
        
        If tbl = 16 Then
            Exit For
        End If
    Next
    
    Range("stateRng").ClearContents
    Range("AnswerAll").ClearContents
    Range("CommentAll").ClearContents
    Range("com2Rng").ClearContents
    
End Sub

Sub newRpt()
Dim ws As Worksheet
Dim WordApp As New Word.Application
Dim objDoc As Object
Dim objRange
Dim col, rw, x As Integer
Dim val, comRng  As String
Dim objBB As BuildingBlock
Dim objTemplate As Template
Dim strPath As String
Dim myFile As Variant
Dim strFile As String
Dim wb As Workbook

    strFile = nmCode _
                & "-" _
                & "Classification Report-" _
                & Format(Now(), "ddmmmyyyy") _
                & ".pdf"
    strFile = ThisWorkbook.path & "\" & strFile
    Set wb = ActiveWorkbook
    
    ActiveWorkbook.Sheets("Report Sheet").Visible = xlSheetVisible
    Set ws = wb.Sheets("Report Sheet")
    ws.Activate
    
    Set objDoc = WordApp.Documents.Add(Template:="\\moe.govt.nz\Shares\Property\Capital Works CS\ADMINISTRATION\CW Database Tools\Reports\Classification Reports\CW_Classification_Report.dotx", _
    NewTemplate:=False, DocumentType:=0, Visible:=True)
    
    WordApp.Visible = True
    
    Set objRange = objDoc.Range()
    
    Set objTable = objDoc.Tables(1)
    
    For i = 1 To 5
        Select Case i
        Case 1
            rw = 5
            col = 4
        Case 2
            rw = 1
            col = 2
        Case 3
            rw = 16
            col = 1
        Case 4
            rw = 2
            col = 2
        Case 5
            rw = 5
            col = 4
        End Select
        
        Set objTable = objDoc.Tables(i)
        
        For j = 1 To rw
            For k = 1 To col
                Select Case True
                Case i = 1 And j = 2 And k = 2
                    val = "RPT_SchoolName"
                Case i = 1 And j = 2 And k = 4
                    val = "RPT_ProjectNum"
                Case i = 1 And j = 3 And k = 2
                    val = "RPT_ProjectName"
                Case i = 1 And j = 3 And k = 4
                    val = "RPT_DelManager"
                Case i = 1 And j = 4 And k = 2
                    val = "RPT_ProjectCat"
                Case i = 1 And j = 4 And k = 4
                    val = "RPT_ProjectVal"
                Case i = 1 And j = 5 And k = 2
                    val = "RPT_DelMethod"
                Case i = 1 And j = 5 And k = 4
                    val = "RPT_ProjDur"
                Case i = 2 And j = 1 And k = 2
                    val = "RPT_Class"
                Case i = 3 And j = 2 And k = 1
                    Do Until Range("RPT_Info" & (j - 1)) = ""
                        objTable.cell(j, k).Range.Text = ws.Range("RPT_Info" & (j - 1)).Value
                        j = j + 1
                    Loop
                Case i = 4 And j = 1 And k = 2
                    val = "RPT_ShrtCode"
                Case i = 4 And j = 2 And k = 2
                    val = "RPT_ClassUser"
                Case i = 5 And j = 2 And k = 2
                    val = "RPT_Date"
                Case Else
                    val = ""
                End Select
                
                If Not val = "" Then
                    objTable.cell(j, k).Range.Text = ws.Range(val).Value
                    val = ""
                End If
            Next
        Next
    Next
    
    WordApp.Selection.EndKey Unit:=wdStory, Extend:=wdMove
    
    Set objTemplate = WordApp.Templates("\\moe.govt.nz\Shares\Property\Capital Works CS\ADMINISTRATION\CW Database Tools\Reports\Classification Reports\CW_Classification_Report.dotx")
    Set objBB = objTemplate.BuildingBlockEntries.Item("CommentTable")
    x = 21
    comRng = Range("b" & x).Value
    objBB.Insert WordApp.Selection.Range
    Set objTable = objDoc.Tables(6)
    
    For i = 1 To 42
        objTable.cell(i, 1).Range.Text = ws.Range("b" & x).Value
        x = x + 1
        i = i + 1
        objTable.cell(i, 1).Range.Text = ws.Range("b" & x).Value
        x = x + 1
        comRng = ws.Range("b" & x).Value
        If comRng = "" Then
            Exit For
        End If
        objBB.Insert WordApp.Selection.Range
    Next
        
    myFile = Application.GetSaveAsFilename _
    (InitialFileName:=strFile, _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Select Folder and FileName to save")
        
    If myFile <> "False" Then
        objDoc.ExportAsFixedFormat _
            OutputFileName:=myFile, _
            ExportFormat:=wdExportFormatPDF, _
            OpenAfterExport:=False, _
            OptimizeFor:=wdExportOptimizeForPrint, _
            Range:=wdExportAllDocument, _
            From:=1, To:=1, _
            Item:=wdExportDocumentContent, _
            IncludeDocProps:=False, _
            KeepIRM:=False, _
            CreateBookmarks:=wdExportCreateHeadingBookmarks, _
            DocStructureTags:=True, _
            BitmapMissingFonts:=False, _
            UseISO19005_1:=False
    
        MsgBox "PDF file has been created."
        ws.Visible = xlSheetVeryHidden
    End If
       
    Application.DisplayAlerts = False
    WordApp.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
    WordApp.Quit
    Set WordApp = Nothing
    Application.DisplayAlerts = True
    
'    objDoc.ExportAsFixedFormat OutputFileName:= _
'    strFile _
'    , ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
'    wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
'    Item:=wdExportDocumentContent, IncludeDocProps:=False, KeepIRM:=False, _
'    CreateBookmarks:=wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
'    BitmapMissingFonts:=False, UseISO19005_1:=False
'
'    ActiveWorkbook.Sheets("Report Sheet").Visible = xlSheetVeryHidden
        
    
End Sub

Sub createPDF()

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

Attribute VB_Name = "Module1"
Public noRate, unRate, lowRate, medRate, hiRate As Integer
'Public variables that are re-used throughout

Public myQLog, myALog, i, j, k, closLog, rskRating, isOne, x As Integer

Public QCode, CCode, QCode2, CCode2, frCode, nmCode, usrName As String

Public ws As Worksheet

Public wb As Workbook

Public rng, rngd, acs As Range

Public Sub TestButton()

    Set ws = Worksheets("Home")
    
    If ws.Range("A1").Value = "Test" Then
        
        Worksheets("Calc").Select
        hideSheets
        frm_Mast.Show
    
    Else
        End If

  
End Sub


Sub ADOFromExcelToAccess()

' exports data from the active worksheet to a table in an Access database
' this procedure must be edited before use
Dim cn As ADODB.Connection, rs As ADODB.Recordset, r As Long
Dim fld, rnge, pth As String
Dim ws As Worksheet

    fld = "DataHead"
    rnge = "DataAns"
    Set ws = Worksheets("Data")
    pth = Range("pthDef")
    
    If pth = "" Then
'        pth = "I:\Shares\Capital Works CS\PORTFOLIO MANAGEMENT\Management Control\Data Files\CWPortfolioManagementDatabase.accdb;"
        pth = "\\moe.govt.nz\Shares\Property\Capital Works CS\PORTFOLIO MANAGEMENT\Management Control\Data Files\CWPortfolioManagementDatabase.accdb;"
'        pth = "H:\CWPortfolioManagementDatabase.accdb;"
    End If
    
    ' connect to the Access database
    Set cn = New ADODB.Connection
        cn.Open "provider = microsoft.ace.oledb.12.0;" & _
        "Data Source=" & pth
        
    ' open a recordset
    Set rs = New ADODB.Recordset
    rs.Open "CWClassificationTool", cn, adOpenKeyset, adLockOptimistic, adCmdTable
    ' all records in a table
    r = 2 ' the start row in the worksheet
        
    'Do While Len(Range("A" & r).Formula) > 0
    ' repeat until first empty cell in column A
        With rs
            .AddNew ' create a new record
            ' add values to each field in the record
            For i = 1 To 38
                rs.Fields(Range(fld & i).Value) = Range(rnge & i).Value
            Next
            ' add more fields if necessary…
            .Update ' stores the new record
        End With
        'r = r + 1 ' next row
    'Loop
    Range(rnge).Clear
    rs.Close
    Set rs = Nothing
    
    cn.Close
    Set cn = Nothing
    
End Sub

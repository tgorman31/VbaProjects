VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Mast 
   Caption         =   "Capital Works Classifiaction Tool"
   ClientHeight    =   10842
   ClientLeft      =   39
   ClientTop       =   377
   ClientWidth     =   15899
   OleObjectBlob   =   "frm_Mast.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm_Mast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Option Explicit\





'Question / Capitons are edited within the excel file
'Ranges withing the excel file are defined by name to allow for easier linking
'Created by Thomas Gorman 2019

Private Sub Image2_Click()
    'access password form to get to excel application
    frm_password.Show
    
End Sub





Sub UserForm_Initialize()
        'set up form on open

    With frm_Mast
     .StartUpPosition = 0
     .Top = 100
     .Left = 100
    End With
    MultiPage1.Value = 0
    myALog = 0
    
     If Not acs Is Nothing And closLog = 0 Then
        MsgBox ("You have been highlighted by the System as a SuperUser" & vbCrLf & _
        "Please use the below Codes in the School Name field and use the Next button to run" & vbCrLf & vbCrLf & _
        "weight adj - allows SuperUser to adjust weightings of each question" & vbCrLf & _
        "access - allows SuperUser to access the background file to edit & monitor the tool" & vbCrLf & vbCrLf & _
        "After closing the file you will be prompted to save, Please do so.")
    End If
    
    closLog = 1
    
        
    'set current date
    Me.tb_date.Value = Format(VBA.Date, "dd/mm/yyyy")
    'Places the current users name within the userform, this is important as only certain users have access to change the details in the file
    Me.tb_YourName = Application.UserName
    usrName = Application.UserName
    Set acs = Range("AccessList").Find(Me.tb_YourName)
    
    
    noRate = Range("noRate").Value
    unRate = Range("unRate").Value
    lowRate = Range("lowRate").Value
    medRate = Range("medRate").Value
    hiRate = Range("hiRate").Value
        
    Application.Calculation = xlCalculationAutomatic
    addLookups
    
    Application.ScreenUpdating = False
    
    showSheets

End Sub

Sub addLookups()

    Set ws = Worksheets("Lookups")
    
    'add Details page lookups
    With Me
        .cmb_EstProjectVal.List = ws.Range("LU_CostRng").Value
        .cmb_EstProjectDur.List = ws.Range("LU_DurRng").Value
        .cmb_ProjCat.List = ws.Range("LU_ProjCat").Value
        .cmb_DelMethod.List = ws.Range("LU_DelMethod").Value
    End With
    
    'add questions from calc sheet
    For i = 1 To 3
        Select Case i
        Case 1
            QCode2 = "SIQUES"
            frCode = "fr_SI_Q"
        Case 2
            QCode2 = "DEPQUES"
            frCode = "fr_Depen_Q"
        Case 3
            QCode2 = "COMQUES"
            frCode = "fr_Complex_Q"
        End Select
        
        'gets total number of each question type form the Lookups sheet
        j = Range(QCode2 & "NO").Value
        
        For k = 1 To j
            Me.Controls(frCode & k).Caption = Range(QCode2 & k)
            If Range(QCode2 & k).Value = "" Then
                Me.Controls(frCode & k).Visible = False
            End If
        Next k
        

    Next i
    
    
End Sub

Private Sub cmd_resetDetails_Click()
        'remove all values from details tab
    With Me
        .tb_SchoolName.Value = ""
        .tb_ProjectName.Value = ""
        .tb_ProjetNo.Value = ""
        .tb_DM.Value = ""

    End With

End Sub
Private Sub cmd_ProjDNext_Click()

    ''looking for keywords in School project to run macro differently
        
    'looks for value "test" within the school field to allow certain users to not have to put in all the entries in the
    'first page for testing purposes
    If Me.tb_SchoolName.Value = "test" And Not acs Is Nothing Then
        detailsproceed
        Exit Sub
    End If
    
    'allows certain users to acces the microsoft excel file to change/monitor any of the information
    If Me.tb_SchoolName.Value = "access" And Not acs Is Nothing Then
        closLog = 0
        Unload Me
        Application.DisplayAlerts = False
        'ThisWorkbook.ChangeFileAccess Mode:=xlReadWrite
        'ActiveWindow.Visible = True
        showSheets
        Application.DisplayAlerts = True
        Worksheets("Analysis").Activate
        MsgBox ("Analysis sheet pulls classification log from access database where the info is stored" & vbCrLf & _
        "Access database is located in the file location displayed on the Home sheet" & vbCrLf & _
        "Project Type & Details weightings can be adjusted on the Details Calc Sheet")
                
        Exit Sub
    End If
    
    'allows certain users to skip to the final page, for testing purposes
    If Me.tb_SchoolName.Value = "skip" And Not acs Is Nothing Then
        With Me
            .txt_ProSum = Me.tb_SchoolName.Value & " - " & Me.tb_ProjectName.Value
            .txt_Overall = Range("FinalClass").Value
            .tb_ProvisClass = Range("ProvisClass").Value
            .lbl_ProvisClass.Visible = True
            .tb_ProvisClass.Visible = True
        End With
        Me.MultiPage1.Pages(0).Visible = False
        Me.MultiPage1.Pages(4).Visible = True
        Me.MultiPage1.Value = 4
        Exit Sub
    End If
    
    'allows certain users to adjust the weight of the questions within the userform
    If Me.tb_SchoolName.Value = "weight adj" Then
        If Not acs Is Nothing Then
            weightAdjust
            With Me
                .cmd_sw.Visible = True
                .tb_SchoolName.Value = ""
                .cmd_ProjDNext.Visible = False
                .cmd_resetDetails.Visible = False
                .cmd_ComplexBack.Visible = False
                .cmd_ComplexNext.Visible = False
                .cmd_SIBack.Visible = False
                .cmd_SINext.Visible = False
                .cmd_DepenBack.Visible = False
                .cmd_DepenNext.Visible = False
            End With
            
            MsgBox ("Weightings are categorized by following guidelines " & vbCrLf & "20 - Minor" & vbCrLf & _
            "40 - Minor/Moderate" & vbCrLf & "60 - Moderate" & vbCrLf & "80 - Major")
            MsgBox ("After adjusting the weighting for each question please click the save weighting button")
            Exit Sub
        Else
            MsgBox ("You are not authorised to adjust question weighting, please contact Portfolio team")
            Exit Sub
        End If
    End If
    
    ' check all details page fields have been filled
    If Me.tb_SchoolName.Value = "" Or Me.tb_ProjectName.Value = "" Or Me.tb_ProjetNo.Value = "" Or Me.tb_DM.Value = "" _
                Or Me.cmb_EstProjectVal.Value = "" Or Me.cmb_ProjCat.Value = "" Or Me.cmb_EstProjectDur.Value = "" _
                Or Me.cmb_DelMethod.Value = "" Then
                
            MsgBox "Please complete ALL fields before proceeding."
       
    ElseIf Not IsNumeric(Me.tb_ProjetNo.Value) Then
        MsgBox ("Please only put numbers in the Project Number Field")
        Me.tb_ProjetNo.Text = ""
        
    Else
        detailsproceed
        
    End If
        
End Sub
Private Sub detailsproceed()
        'if all details fields complete then continue
Dim lupsTC As Range
Dim l As Worksheet
Dim mystr As String
Dim mydate As Variant
        
    Set l = Worksheets("Lookups")
    Set ws = Worksheets("Data")
    Set rngd = ws.Cells(Rows.count, "A").End(xlUp)
    
    ws.Unprotect ("CW")
    'write form values to Data sheet
    With rngd
        .Offset(1, 0) = Me.tb_date.Value
        .Offset(1, 1) = Me.tb_YourName.Value
        .Offset(1, 2) = Me.tb_SchoolName.Value
        .Offset(1, 3) = Me.tb_ProjectName.Value
        .Offset(1, 4) = Me.tb_ProjetNo.Value
        .Offset(1, 5) = Me.tb_DM.Value
        .Offset(1, 6) = Me.cmb_EstProjectVal.Value
        .Offset(1, 7) = Me.cmb_EstProjectDur.Value
        .Offset(1, 8) = Me.cmb_ProjCat.Value
        .Offset(1, 9) = Me.cmb_DelMethod.Value
    End With
    
    'writes detail values to DetailsCalc sheet
    Range("ValueCalc").Value = Me.cmb_EstProjectVal.Value
    Range("DurCalc").Value = Me.cmb_EstProjectDur.Value
    Range("TypeCalc").Value = Me.cmb_ProjCat.Value
    Range("DelMetCalc").Value = Me.cmb_DelMethod.Value
    
    nmCode = Me.tb_ProjetNo.Value & "-" & Me.tb_SchoolName.Value & "-" & Me.tb_ProjectName.Value
    
    'set up next page
    Me.MultiPage1.Pages(0).Visible = False
    Me.MultiPage1.Pages(1).Visible = True
    Me.MultiPage1.Value = 1
                                           
End Sub


Private Sub cmd_SIBack_Click()
'to go back to the previous page

    Me.MultiPage1.Pages(1).Visible = False
    Me.MultiPage1.Pages(0).Visible = True
    Me.MultiPage1.Value = 0
    Range("DataAns").Clear

End Sub
Private Sub cmd_DepenBack_Click()
'to go back to the previous page

    Me.MultiPage1.Pages(2).Visible = False
    Me.MultiPage1.Pages(1).Visible = True
    Me.MultiPage1.Value = 1

End Sub
Private Sub cmd_ComplexBack_Click()
'to go back to the previous page

    Me.MultiPage1.Pages(3).Visible = False
    Me.MultiPage1.Pages(2).Visible = True
    Me.MultiPage1.Value = 2

End Sub
Private Sub opt_SI_Q1_Y_Click()
    Me.fr_SI_Q3.Visible = False
    Me.opt_SI_Q3_N = True
    Me.fr_Complex_Q8.Visible = False
    Me.opt_Complex_Q8_N = True
End Sub

Private Sub opt_SI_Q1_N_Click()
    Me.fr_SI_Q3.Visible = True
    Me.opt_SI_Q3_N = False
    Me.fr_Complex_Q8.Visible = True
    Me.opt_Complex_Q8_N = False
End Sub
Private Sub cmd_SINext_Click()
    'submit Strategic Importance questions
    myQLog = 1
    
    Call QuestLog
    
    'Checks if user has not answered a question and wants to go back to answer it
    If myALog = 1 Then
        Exit Sub
    End If
    
    Me.MultiPage1.Pages(1).Visible = False
    Me.MultiPage1.Pages(2).Visible = True
    Me.MultiPage1.Value = 2

End Sub

Private Sub cmd_DepenNext_Click()
    'submit Dependencies questions
    myQLog = 2
    
    Call QuestLog
    
    'Checks if user has not answered a question and wants to go back to answer it
    If myALog = 1 Then
        Exit Sub
    End If
        
    Me.MultiPage1.Pages(2).Visible = False
    Me.MultiPage1.Pages(3).Visible = True
    Me.MultiPage1.Value = 3
End Sub

Private Sub cmd_ComplexNext_Click()
    'submit Complexities questions
    myQLog = 3
    
    Call QuestLog
    
    'Checks if user has not answered a question and wants to go back to answer it
    If myALog = 1 Then
        Exit Sub
    End If
    
    'Calls sub to calculate classification
    Call profileClass
        
    Me.MultiPage1.Pages(3).Visible = False
    Me.MultiPage1.Pages(4).Visible = True
    Me.MultiPage1.Value = 4
    
    'allows certain users to veiw what the original weighthing from the details page was before the
    'rest of the questions
    If Not acs Is Nothing Then
        Me.tb_ProvisClass.Visible = True
        Me.lbl_ProvisClass.Visible = True
    End If
    
End Sub
Private Sub QuestLog()
'QuestLog logs the answers to the questions, check if any is unanswered and allow the user to go back and
'answer or provide a comment
Dim opt, cel As String

    Set ws = Worksheets("Calc")
    
    Select Case myQLog
    
    Case 1
        opt = "opt_SI_Q"
        frCode = "fr_SI_Q"
        cel = "STRQ"
        QCode = "SIQUES"
        CCode = "STRC"
        Set rng = ws.Range("STRRNG")
    Case 2
        opt = "opt_Depen_Q"
        frCode = "fr_Depen_Q"
        cel = "DEPQ"
        QCode = "DEPQUES"
        CCode = "DEPC"
        Set rng = ws.Range("DEPRNG")
    Case 3
        opt = "opt_Complex_Q"
        frCode = "fr_Complex_Q"
        cel = "COMQ"
        QCode = "COMQUES"
        CCode = "COMC"
        Set rng = ws.Range("COMRNG")
    End Select
    
    j = Range(QCode & "NO").Value
    
    'Searching for unanswered questions
        
        For i = 1 To j
            If Me.Controls(opt & i & "_Y") = False And Me.Controls(opt & i & "_N") = False Then
                If Me.Controls(frCode & i).Visible = True Then
                    frm_Unsure.Show
                    If myALog = 1 Then
                        Exit Sub
                    ElseIf myALog = 2 Then
                        Exit For
                    End If
                End If
            End If
        Next i
        
    
    'Logging selections for Questions
    
        For i = 1 To j
            
            If Me.Controls(opt & i & "_Y") = True Then
                Range(cel & i).Value = "4"
            ElseIf Me.Controls(opt & i & "_N") = True Then
                Range(cel & i).Value = Range("noRate").Value
            ElseIf Me.Controls(opt & i & "_Y") = False And Me.Controls(opt & i & "_N") = False And Me.Controls(frCode & i).Visible = True Then
                Range(cel & i).Value = Range("unRate").Value
            ElseIf Me.Controls(opt & i & "_Y") = False And Me.Controls(opt & i & "_N") = False And Me.Controls(frCode & i).Visible = False Then
                Range(cel & i).Value = Range("noRate").Value
            End If
            
        Next i
        
        i = 0
    
    'Searching if the user wants to add a comment if they are unsure or if they answered yes then requesting and saving comments for these questions
    
        For Each cell In rng
            i = i + 1
            If cell.Value = "1" Or cell.Value = "4" Then
                If cell.Value = "1" Then 'if the cell value is 2 then sets isOne to 1 - this stops the riskrating appearing in the comment window
                    isOne = 1
                ElseIf cell.Value = "4" Then 'if the cell value is 4 then sets isOne to 0 - this allows the riskrating to appear in the comment window
                    isOne = 0
                End If
                QCode2 = QCode & i
                CCode2 = CCode & i
                frm_Comment.Show
                If Not cell.Value = "0" Then 'adjusts the value of the answer in the calc sheet based on what the user selected in the risk rating
                    cell.Value = rskRating
                End If
            End If
        Next
            
        myALog = 0
    
End Sub
            
Sub weightAdjust()

Dim x, y, z As Integer
Dim tb As String
    
    'ThisWorkbook.ChangeFileAccess Mode:=xlReadWrite

    'loops through each page and sets weighting from excel file
    For x = 1 To 3
        Me.MultiPage1.Pages(x).Visible = True
        Me.MultiPage1.Value = x
        
        Select Case x
        Case 1
            tb = "STRW"
            QCode2 = "SIQUES"
            Me.MultiPage1.Pages(1).Visible = True
            Me.MultiPage1.Value = 1
        Case 2
            tb = "DEPW"
            QCode2 = "DEPQUES"
            Me.MultiPage1.Pages(2).Visible = True
            Me.MultiPage1.Value = 2
        Case 3
            tb = "COMW"
            QCode2 = "COMQUES"
            Me.MultiPage1.Pages(3).Visible = True
            Me.MultiPage1.Value = 3
        End Select
        
        z = Range(QCode2 & "NO").Value
        
        For y = 1 To z
            Me.Controls("tb_" & tb & y).Value = Range(tb & y)
            Me.Controls("tb_" & tb & y).Visible = True
        Next

    Next
    
    'ThisWorkbook.ChangeFileAccess Mode:=xlReadOnly

    Me.MultiPage1.Value = 0

End Sub

Private Sub cmd_sw_Click()

Dim x, y, z As Integer
Dim tb As String


    'loops through each page and sets adjusted weighting to excel file
    For x = 1 To 3
        Me.MultiPage1.Pages(x).Visible = True
        Me.MultiPage1.Value = x
        
        Select Case x
        Case 1
            tb = "STRW"
            QCode2 = "SIQUES"
        Case 2
            tb = "DEPW"
            QCode2 = "DEPQUES"
        Case 3
            tb = "COMW"
            QCode2 = "COMQUES"
        End Select
        
        z = Range(QCode2 & "NO").Value
        
        For y = 1 To z
            Range(tb & y).Value = Me.Controls("tb_" & tb & y)
            Me.Controls("tb_" & tb & y).Visible = False
        Next
        
        Me.MultiPage1.Pages(x).Visible = False
        
    Next
    
    With Me
        .cmd_sw.Visible = False
        .cmd_ProjDNext.Visible = True
        .cmd_resetDetails.Visible = True
        .cmd_ComplexBack.Visible = True
        .cmd_ComplexNext.Visible = True
        .cmd_SIBack.Visible = True
        .cmd_SINext.Visible = True
        .cmd_DepenBack.Visible = True
        .cmd_DepenNext.Visible = True
    End With

    MsgBox ("Weighting of questions has been saved")
    ActiveWorkbook.Save
    Me.MultiPage1.Value = 0
    
End Sub
            
Sub profileClass()

Dim x, z As Integer

    x = 13
    
    'writes the answers to the questions to the data sheet in the workbook
    For z = 1 To Range("TOTALNO").Value
        rngd.Offset(1, x) = Range("ANS" & z & "_").Value
        x = x + 1
    Next
    
    'adds the calculated classification to the summary page
    With Me
        .txt_ProSum = Me.tb_SchoolName.Value & " - " & Me.tb_ProjectName.Value
        .txt_Overall = Range("FinalClass").Value ' error here if first page details are not entered
        .tb_ProvisClass = Range("ProvisClass").Value
    End With

    rngd.Offset(1, 10) = Range("ProvisClass").Value
    rngd.Offset(1, 11) = Range("FinalClass").Value
    
    'Range("CommentAll").Clear
    
End Sub

Private Sub cmb_ResetClass_Click()
    frm_Reset.Show
End Sub
            
Private Sub cmd_Close_Click()
    
    Range("DataAns38").Value = "Yes"
    
    Mail_Submit
  
    Application.DisplayAlerts = False
    
    ThisWorkbook.Saved = True
    
    Application.Run ("RPT_Update.UpdateReportSheet")
    
    Application.Run ("RPT_Update.newRpt")

    Application.Run ("Module1.ADOFromExcelToAccess")
    
'    Range("CommentAll").Clear
'    Range("stateRng").Clear
'    Range("com2Rng").Clear
    
    
    Set rng = Nothing
    Set rngd = Nothing
    Set acs = Nothing
    Set ws = Nothing
    Set wb = Nothing
'    QCode = Nothing
'    CCode = Nothing
'    QCode2 = Nothing
'    CCode2 = Nothing
'    frCode = Nothing
'    nmCode = Nothing
'    usrName = Nothing
'    myQLog = Nothing
'    myALog = Nothing
'    i = Nothing
'    k = Nothing
'    closLog = Nothing
'    rskRating = Nothing
'    isTwo = Nothing
     
    UserFormTerminate
    
    hideSheets
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Sub UserFormTerminate()
    
    Worksheets("Data").Protect ("CW")
    Application.DisplayAlerts = False
    Unload Me
    'ActiveWindow.Visible = True
    Application.Quit
    Application.DisplayAlerts = True
    
End Sub


Private Sub UserForm_Terminate()
    If closLog = 1 Then
        UserFormTerminate
    Else: Unload Me
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

        If CloseMode = 0 Then Cancel = True

End Sub

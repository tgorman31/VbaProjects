Attribute VB_Name = "Module2"
Sub hideSheets()
Attribute hideSheets.VB_ProcData.VB_Invoke_Func = " \n14"
'Hides all sheets except data

    ActiveWorkbook.Sheets("Home").Visible = xlSheetVeryHidden
    ActiveWorkbook.Sheets("Lookups").Visible = xlSheetVeryHidden
    ActiveWorkbook.Sheets("Calc").Visible = xlSheetVeryHidden
    ActiveWorkbook.Sheets("Matrix").Visible = xlSheetVeryHidden
    ActiveWorkbook.Sheets("DetailsCalc").Visible = xlSheetVeryHidden
    ActiveWorkbook.Sheets("Report Sheet").Visible = xlSheetVeryHidden
    ActiveWorkbook.Sheets("Analysis").Visible = xlSheetVeryHidden
    
End Sub

Sub showSheets()
'shows sheets

    ActiveWorkbook.Sheets("Home").Visible = xlSheetVisible
    ActiveWorkbook.Sheets("Lookups").Visible = xlSheetVisible
    ActiveWorkbook.Sheets("Calc").Visible = xlSheetVisible
    ActiveWorkbook.Sheets("Matrix").Visible = xlSheetVisible
    ActiveWorkbook.Sheets("DetailsCalc").Visible = xlSheetVisible
    ActiveWorkbook.Sheets("Report Sheet").Visible = xlSheetVisible
    ActiveWorkbook.Sheets("Analysis").Visible = xlSheetVisible
    
End Sub

Sub Mail_noSubmit()
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Working in Office 2000-2016
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    strbody = "Kia Ora" & vbNewLine & vbNewLine & _
              "This is an automated notification that" & vbNewLine & vbNewLine & _
              usrName & vbNewLine & vbNewLine & _
              "Has completed the Classification tool but has not submitted their answers" & vbNewLine & vbNewLine & _
              "Please do not respond to this message" & vbNewLine & vbNewLine & _
              "Thanks"

    On Error Resume Next
    With OutMail
'        .To = "thomas.gorman@education.govt.nz"
        .To = "Cwportfolio.Team@education.govt.nz"
        .CC = ""
        .BCC = ""
        .Subject = "Classification Tool Reset Notification"
        .Body = strbody
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
        .DeleteAfterSubmit = True
        .Send   'or use .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Sub Mail_Submit()
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Working in Office 2000-2016
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    strbody = "Kia Ora" & vbNewLine & vbNewLine & _
              "This is an automated notification that" & vbNewLine & vbNewLine & _
              usrName & vbNewLine & vbNewLine & _
              "Has completed the Classification tool and submitted their answers" & vbNewLine & vbNewLine & _
              "Please do not respond to this message." & vbNewLine & vbNewLine & _
              "Thanks"

    On Error Resume Next
    With OutMail
'        .To = "thomas.gorman@education.govt.nz"
        .To = "Cwportfolio.Team@education.govt.nz"
        .CC = ""
        .BCC = ""
        .Subject = "Classification Tool Completion Notification"
        .Body = strbody
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
        .DeleteAfterSubmit = True
        .Send   'or use .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Attribute VB_Name = "InventoryAutomation"
Option Explicit

' Returns the value in a named column of a table row
Private Function ColVal(tbl As ListObject, r As ListRow, colName As String) As Variant
    ColVal = r.Range(1, tbl.ListColumns(colName).Index).Value
End Function

' Builds a list of labs that have at least one discrepancy
Private Function LabsWithIssues(rec As ListObject) As Collection
    Dim labs As New Collection, r As ListRow, id As String
    On Error Resume Next ' adding a duplicate key errors; skipping it keeps the list unique
    For Each r In rec.ListRows
        If ColVal(rec, r, "Status") <> "Match" Then
            id = CStr(ColVal(rec, r, "LabID"))
            labs.Add id, id
        End If
    Next r
    On Error GoTo 0
    Set LabsWithIssues = labs
End Function

' Macro 1: one discrepancy report workbook per lab
Sub CreateLabReports()
    Dim rec As ListObject, labs As Collection, lab As Variant
    Dim r As ListRow, wb As Workbook, ws As Worksheet
    Dim folder As String, outRow As Long

    Set rec = ThisWorkbook.Worksheets("Reconciliation").ListObjects("Reconciliation")
    Set labs = LabsWithIssues(rec)
    folder = ThisWorkbook.Path & Application.PathSeparator & "reports" & Application.PathSeparator
    If Dir(folder, vbDirectory) = "" Then MkDir folder

    Application.ScreenUpdating = False
    For Each lab In labs
        Set wb = Workbooks.Add
        Set ws = wb.Worksheets(1)
        ws.Range("A1").Value = "Chemical inventory discrepancies: lab " & lab
        ws.Range("A2").Value = "Generated " & Format(Date, "yyyy-mm-dd")
        ws.Range("A4:D4").Value = Array("Container", "Chemical", "Issue", "Detail")
        ws.Range("A1,A4:D4").Font.Bold = True
        outRow = 5
        For Each r In rec.ListRows
            If CStr(ColVal(rec, r, "LabID")) = lab And ColVal(rec, r, "Status") <> "Match" Then
                ws.Cells(outRow, 1).Value = ColVal(rec, r, "ContainerID")
                ws.Cells(outRow, 2).Value = ColVal(rec, r, "Chemical")
                ws.Cells(outRow, 3).Value = ColVal(rec, r, "Status")
                ws.Cells(outRow, 4).Value = ColVal(rec, r, "Detail")
                outRow = outRow + 1
            End If
        Next r
        ws.Columns("A:D").AutoFit
        Application.DisplayAlerts = False
        wb.SaveAs folder & "Lab_" & lab & "_discrepancies.xlsx", FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        wb.Close SaveChanges:=False
    Next lab
    Application.ScreenUpdating = True
    MsgBox labs.Count & " lab reports saved to " & folder
End Sub

' Macro 2: one reminder email draft per lab, written to a sheet
Sub DraftReminderEmails()
    Dim rec As ListObject, ppl As ListObject, ws As Worksheet
    Dim labs As Collection, lab As Variant, r As ListRow, p As ListRow
    Dim issues As Long, toAddr As String, body As String, outRow As Long

    Set rec = ThisWorkbook.Worksheets("Reconciliation").ListObjects("Reconciliation")
    Set ppl = ThisWorkbook.Worksheets("People").ListObjects("People")
    Set labs = LabsWithIssues(rec)

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("EmailDrafts")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "EmailDrafts"
    End If
    ws.Cells.Clear
    ws.Range("A1:E1").Value = Array("LabID", "To", "Subject", "Body", "Issues")
    ws.Range("A1:E1").Font.Bold = True

    outRow = 2
    For Each lab In labs
        issues = 0
        For Each r In rec.ListRows
            If CStr(ColVal(rec, r, "LabID")) = lab And ColVal(rec, r, "Status") <> "Match" Then issues = issues + 1
        Next r

        toAddr = ""
        For Each p In ppl.ListRows
            If CStr(ColVal(ppl, p, "lab_id")) = lab And ColVal(ppl, p, "role") = "Lab manager" Then
                toAddr = ColVal(ppl, p, "email")
                Exit For
            End If
        Next p
        If toAddr = "" Then toAddr = "NO LAB MANAGER ON FILE"

        body = "Hello," & vbLf & vbLf & _
               "During our recent inventory visit we found " & issues & " items in lab " & lab & _
               " that need updating in the chemical inventory system. The attached report lists each one." & vbLf & vbLf & _
               "Please make the updates by " & Format(Date + 14, "mmmm d") & _
               ". Reply to this email if anything is unclear or you would like help." & vbLf & vbLf & _
               "Thank you," & vbLf & "EH&S Data Programs"

        ws.Cells(outRow, 1).Value = lab
        ws.Cells(outRow, 2).Value = toAddr
        ws.Cells(outRow, 3).Value = "Action needed: chemical inventory updates for lab " & lab
        ws.Cells(outRow, 4).Value = body
        ws.Cells(outRow, 5).Value = issues
        outRow = outRow + 1
    Next lab
    ws.Columns("A:C").AutoFit
    ws.Columns("D").ColumnWidth = 80
    ws.Columns("D").WrapText = True
    MsgBox labs.Count & " email drafts written to the EmailDrafts sheet"
End Sub

' Macro 3: the whole monthly cycle in one click
Sub RunMonthlyCycle()
    ThisWorkbook.RefreshAll
    CreateLabReports
    DraftReminderEmails
End Sub

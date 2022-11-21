Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
        Range("F:F").NumberFormat = "0%"
        Range("G:G").NumberFormat = "0.00"
        Range("H:AM").NumberFormat = "0"
        Range("I:J").NumberFormat = "0.0"
        Range("E:E").NumberFormat = "dd/mm/yyyy"
End Sub
---------------------------------------------------------------------------------------------
Private Sub Workbook_Open(Optional HideMe As String)
'***********************************
'** Author: Marco Cot DAS:A669714 **
'***********************************
'
' ACCOUNT: Global (weekly reporting)
' Automatically sets Monday date and prepares sheet for data entry

'Yes/No menu
Dim Choose As Integer
Dim Clear As Integer

'Data input
Dim EnteredYear As Long
Dim Weeknumber As Integer

'Constant "Jan1" declared "01/01/"
Dim Jan1Year As String '("Jan1+Year")
Dim Jan1Date As Date

'References for sheet deletion
Dim ReportedWeek As Integer
Dim Currentsheet As String
Dim PrevSheetNo As Integer
Dim PreviousSheet As String
Dim RowsCounter As Long


    On Error GoTo ErrorControl
        Application.ScreenUpdating = False
        Application.AskToUpdateLinks = False
        Application.DisplayAlerts = False
Choose = MsgBox("Do you need to report a new week?", vbQuestion + vbYesNo, "UPDATE")
    If Choose = vbYes Then
        ActiveSheet.Select
        ActiveSheet.Copy After:=Sheets(Sheets.Count)
        Range("DATECLEAN").ClearContents
        
ControlYear:
        EnteredYear = Application.InputBox("Please enter reported YEAR", Default:=Year(Date))
        If EnteredYear < 2022 Then
            MsgBox "Wrong year set", vbCritical + vbOKOnly, "ERROR"
            GoTo ControlYear
        End If
        ReportedWeek = WorksheetFunction.IsoWeekNum(Date) - 1
        
ControlWeek:
        Weeknumber = Application.InputBox("Please enter reported week number [1-52]", Default:=ReportedWeek)
        If Weeknumber < 1 Or Weeknumber > 52 Then
            MsgBox "Wrong week number set", vbCritical + vbOKOnly, "ERROR"
            GoTo ControlWeek
        End If
            Jan1Year = Jan1 & EnteredYear
            Jan1Date = DateValue(Jan1Year)

            If Weeknumber < 10 Then
                ActiveSheet.Name = "RW0" & Weeknumber
                Else: ActiveSheet.Name = "RW" & Weeknumber
            End If

        Range("F:F").NumberFormat = "0%"
        Range("G:G").NumberFormat = "0.00"
        Range("H:AM").NumberFormat = "0"
        Range("I:J").NumberFormat = "0.0"
        Range("WEEKMONDAY").NumberFormat = "dd/mm/yyyy"

        Range("WEEKMONDAY").FormulaR1C1 = (Weeknumber - 1) * 7 + (Jan1Date + 7) - Weekday(Jan1Date + 5, 1)

        RowsCounter = Application.WorksheetFunction.CountA(Range("A:A"))
        Range("WEEKMONDAY").Copy
                  
        Range("E2:E" & RowsCounter).Select
        ActiveSheet.Paste
    
        Range("WEEKMONDAY").Select
        Application.CutCopyMode = False
        Application.ScreenUpdating = True
        Application.AskToUpdateLinks = True
        Application.DisplayAlerts = True
        
        Clear = MsgBox("Do want to delete the previous sheet?", vbQuestion + vbYesNo, "DELETE")
        If Clear = vbYes Then
            Currentsheet = Right(ActiveSheet.Name, 2)
            PrevSheetNo = CInt(Currentsheet) - 1
            If PrevSheetNo < 10 Then
                PreviousSheet = "RW0" & PrevSheetNo
                Sheets(PreviousSheet).Delete
            Else:   PreviousSheet = "RW" & PrevSheetNo
                    Sheets(PreviousSheet).Delete
                        End If
    Else:   Application.ScreenUpdating = True
                Application.AskToUpdateLinks = True
                Application.DisplayAlerts = True
                        
            End If
        End If
       Exit Sub
ErrorControl:
        If Err.Number = 9 Then
        MsgBox "Cannot delete previous sheet" & vbCrLf & "Please delete manually", vbInformation, "ERROR"
        Else
        MsgBox "An error occurred." & vbCrLf & "Code: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbInformation, "ERROR"
        End If
End Sub

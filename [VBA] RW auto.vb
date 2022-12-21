'********************************************************************************************************************************************
'Author: Marco Cot DAS:A669714
'********************************************************************************************************************************************

Private Sub Workbook_Open()

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

ActiveSheet.Unprotect "NoEdit"
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

        Range("1048576:1048576").Copy
        Range("2:101").PasteSpecial xlFormats

        Range("WEEKMONDAY").FormulaR1C1 = (Weeknumber - 1) * 7 + (Jan1Date + 7) - Weekday(Jan1Date + 5, 1)

        RowsCounter = Application.WorksheetFunction.CountA(Range("A:A"))
        Range("WEEKMONDAY").Copy
                 
        Range("E2:E" & RowsCounter).Select
        ActiveSheet.Paste
    
        Range("WEEKMONDAY").Select
        Application.CutCopyMode = False
        Application.ScreenUpdating = True
        Application.AskToUpdateLinks = True
                
        Clear = MsgBox("Do want to delete the previous sheet?", vbQuestion + vbYesNo, "DELETE")
        If Clear = vbYes Then
            Currentsheet = Right(ActiveSheet.Name, 2)
            PrevSheetNo = CInt(Currentsheet) - 1
            If PrevSheetNo < 10 Then
                PreviousSheet = "RW0" & PrevSheetNo
                Sheets(PreviousSheet).Delete
                Application.DisplayAlerts = True
            Else:   PreviousSheet = "RW" & PrevSheetNo
                    Sheets(PreviousSheet).Delete
                    Application.DisplayAlerts = True
                        End If
    Else:   Application.ScreenUpdating = True
            
                Application.AskToUpdateLinks = True
                Application.DisplayAlerts = True
                        
            End If
        Else:     ActiveSheet.Protect Password:="NoEdit", AllowFormattingColumns:=True
        End If
       Exit Sub
ErrorControl:
        If Err.Number = 9 Then
            MsgBox "Cannot delete previous sheet" & vbCrLf & "Please delete manually", vbInformation, "ERROR"
            ActiveSheet.Protect Password:="NoEdit", AllowFormattingColumns:=True
        ElseIf Err.Number = 1004 Then
            MsgBox "Duplicate tab name" & vbCrLf & vbCrLf & "Please correct manually" & vbCrLf & "and delete previous tab.", vbInformation, "ERROR"
            Range("WEEKMONDAY").FormulaR1C1 = (Weeknumber - 1) * 7 + (Jan1Date + 7) - Weekday(Jan1Date + 5, 1)
            ActiveSheet.Protect Password:="NoEdit", AllowFormattingColumns:=True
        Else
        MsgBox "An error occurred." & vbCrLf & "Code: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbInformation, "ERROR"
        ActiveSheet.Protect Password:="NoEdit", AllowFormattingColumns:=True
        End If
End Sub

'********************************************************************************************************************************************
'Author: Marco Cot DAS:A669714
'********************************************************************************************************************************************

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    If Application.WorksheetFunction.CountA(Range("A:A")) <> Application.WorksheetFunction.CountA(Range("AN:AN")) Then
        
        Dim Recuento As Integer
        Dim RecuentoDelete As Integer
        
        If Application.WorksheetFunction.CountA(Range("A:A")) > Application.WorksheetFunction.CountA(Range("AN:AN")) Then
            ActiveSheet.Unprotect "NoEdit"
            Application.ScreenUpdating = False
            Recuento = Application.WorksheetFunction.CountA(Range("A:A"))
            Range("AN2:AO2").Copy
            Range("AN3:AN" & Recuento).Select
            ActiveSheet.Paste
            Range("AQ2:BA2").Copy
            Range("AQ3:AQ" & Recuento).Select
            ActiveSheet.Paste
            Range("E2").Copy
            Range("E3:E" & Recuento).Select
            ActiveSheet.Paste
            Range("B" & Recuento).Select
            Application.ScreenUpdating = True
            Columns("A:BA").AutoFit
            ActiveSheet.Protect Password:="NoEdit", AllowFormattingColumns:=True
        
        Else: If Application.WorksheetFunction.CountA(Range("A:A")) < Application.WorksheetFunction.CountA(Range("AN:AN")) Then ActiveSheet.Unprotect "NoEdit"
            Application.ScreenUpdating = False
            Recuento = Application.WorksheetFunction.CountA(Range("A:A"))
            RecuentoDelete = Application.WorksheetFunction.CountA(Range("A:A")) + 1
            Range("AN" & RecuentoDelete & ":BA100").ClearContents
            Range("E" & RecuentoDelete & ":E100").ClearContents
            Range("B" & Recuento).Select
            Application.ScreenUpdating = True
            Columns("A:BA").AutoFit
            ActiveSheet.Protect Password:="NoEdit", AllowFormattingColumns:=True
            
        End If
    End If

End Sub

'********************************************************************************************************************************************
'Author: Marco Cot DAS:A669714
'********************************************************************************************************************************************

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim improperData As Double
    Dim improperDataL2 As Double
    Dim improperSLA As Double
    Dim improperCSAT As Double
    Dim wrongCSAT As Double
    Dim BlnEventState As Boolean
    Dim SelectedButton As Integer
    
        

    ActiveSheet.Unprotect "NoEdit"
    Range("1048576:1048576").Copy
    Range("2:101").PasteSpecial xlFormats
    Range("Y1").Select
    Columns("A:BA").AutoFit
    ActiveSheet.Protect Password:="NoEdit", AllowFormattingColumns:=True
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    improperData = Application.WorksheetFunction.Max(Range("AQ:AQ"))
    improperDataL2 = Application.WorksheetFunction.Max(Range("AT:AT"))
    improperSLA = Application.WorksheetFunction.Max(Range("F:F"))
    wrongCSAT = Application.WorksheetFunction.Min(Range("G:G"))
    improperCSAT = Application.WorksheetFunction.Max(Range("G:G"))
    BlnEventState = Application.EnableEvents
    Application.EnableEvents = True

      
    If improperData > 1 Or improperDataL2 > 1 Then
        Cancel = True
           MsgBox "Utilisation cannot be above 100%" & vbCrLf & "Please recheck volume and AHT", vbCritical + vbOKOnly, "DATA REVIEW"
           Range("L:S").Select
            Cancel = True
            Application.EnableEvents = BlnEventState
        
        ElseIf improperSLA > 1 Then
                     Cancel = True
                    MsgBox "SLA cannot exceed 100%", vbCritical + vbOKOnly, "DATA REVIEW"
                Range("F2").Select
                 Cancel = True
                 Application.EnableEvents = BlnEventState
        ElseIf improperCSAT > 10 Then
                     Cancel = True
                    MsgBox "CSAT cannot exceed 10", vbCritical + vbOKOnly, "DATA REVIEW"
                Range("G2").Select
                 Cancel = True
                 Application.EnableEvents = BlnEventState
        ElseIf wrongCSAT > 0 And wrongCSAT < 1 Then
                     Cancel = True
                    MsgBox "CSAT must be a value from 1 to 10" & vbCrLf & "If CSAT is actually below 1, state 1.00 instead", vbCritical + vbOKOnly, "DATA REVIEW"
                Range("G2").Select
                 Cancel = True
                 Application.EnableEvents = BlnEventState
        Else
        
                Application.ScreenUpdating = True
                SelectedButton = MsgBox("YES = Create final files" & vbCrLf & "NO = just close and save locally" & vbCrLf & "Cancel = If red data are visible, please review", vbYesNoCancel, "Last check")

           If SelectedButton = vbYes Then
                Range("A2").Select
                ActiveWorkbook.SaveAs Filename:="C:\temp\" & Range("A2").Value & " SD Weekly Data INTERNAL " & Year(Range("E2").Value) & "-" & ActiveSheet.Name
                ActiveWorkbook.SaveAs Filename:="C:\temp\" & Range("A2").Value & " SD Weekly Data TO SEND " & Year(Range("E2").Value) & "-" & ActiveSheet.Name, FileFormat:=51
                MsgBox "Your files has been saved in C:\temp" & vbCrLf & vbCrLf & "INTERNAL to be kept and reused" & vbCrLf & "COLLECT to be sent", vbInformation, "Saving confirmation"
                Shell "explorer.exe" & " " & "C:\temp", vbNormalFocus
  
            ElseIf SelectedButton = vbCancel Then
                BlnEventState = Application.EnableEvents
                Application.EnableEvents = True
                Cancel = True
                Application.EnableEvents = BlnEventState
        End If
    End If
 
End Sub

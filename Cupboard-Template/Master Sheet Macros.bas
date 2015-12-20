Attribute VB_Name = "Module1"

' NewWeekWorkSheet creates a new worksheet for the week, with appropriate column widths and labels
Private Sub NewWeekWorkSheet()
    ' Sheet index
    Dim i As Integer
    
    ' Sheets.Count outputs the number of current sheets (this number is the end of the Worksheets array)
    Worksheets.Add After:=Sheets(Sheets.Count)
    
    i = (Sheets.Count)
    
    ' Sets Lables and style
    Worksheets(i).Range("A1") = "Date"
    Worksheets(i).Range("B1") = "ID"
    Worksheets(i).Range("C1") = "Items"
    Worksheets(i).Range("D1") = "Box"
    Worksheets(i).Range("E1") = "Time In"
    Worksheets(i).Range("A1:E1").Font.Bold = True
    Worksheets(i).Range("A1:E1").HorizontalAlignment = xlCenter
    Worksheets(i).Columns("A").ColumnWidth = 10
    Worksheets(i).Columns("B").ColumnWidth = 8.5
    Worksheets(i).Columns("C").ColumnWidth = 18
    Worksheets(i).Columns("D").ColumnWidth = 4
    Worksheets(i).Columns("E").ColumnWidth = 11
    
    ' Sets the name of the Worksheet
    Worksheets(i).Name = Format(Date, "mm-dd-yy") & "(" & i & ")"

End Sub

' WeeklyReportsP1 closes out the last sheet on the workbook. To do this, it calls TotalVists, which totals the number
' of students that used the cupboard this reporting week. It calls UserForm3, which then tallies the data for the sheets
' selected by the user and posts that data to the Totals page.
Private Sub WeeklyReportsP1()
    ' B_ColCounter represents the number of rows with data in them.
    Dim Items, Unique, Visits, B_ColCounter As Integer
    
    ' The following finds the last row with information in it.
    B_ColCounter = Worksheets(Sheets.Count).Cells(Rows.Count, 2).End(xlUp).Row
    
    ' Calls TotalVisits and passes number of rows with data. TotalVisits returns the number of students that visited on this sheet.
    Visits = TotalVisits(B_ColCounter)
    
    ' Moves focus to the first empty row.
    B_ColCounter = B_ColCounter + 1
    
    ' The following line was needed with the original code, before the data entry UserForm was made. It should be able to be phased
    ' out at the beginning of the 2016-2017 school year, as the reason for its existence should be out of habit. It might be possible
    ' to add a check to see if it's needed. It simply clears any excess dates that were entered without corrisponding student data.
    Worksheets(Sheets.Count).Range("A" & B_ColCounter & ":A1000") = ""
    
    ' Sets number of Items distributed to number of current rows, minus 1 for the label row, and minus one for the row added above.
    Items = B_ColCounter - 2
    
    ' Copies all unique values from Column B to Column G
    Worksheets(Sheets.Count).Range("B2:B1000").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets(Sheets.Count).Range("G1"), Unique:=True
    
    ' Sets the unique amount of students to the number of used rows in column G
    Unique = Worksheets(Sheets.Count).Cells(Rows.Count, 7).End(xlUp).Row
    
    ' The following creates a totals block at the bottom of the sheet. The block starts in column C with one row of space between
    ' the last set of data.
    Worksheets(Sheets.Count).Cells(B_ColCounter + 1, 3) = "Total Visits:"
    Worksheets(Sheets.Count).Cells(B_ColCounter + 1, 4) = Visits
    Worksheets(Sheets.Count).Cells(B_ColCounter + 2, 3) = "Total Items:"
    Worksheets(Sheets.Count).Cells(B_ColCounter + 2, 4) = Items
    Worksheets(Sheets.Count).Cells(B_ColCounter + 3, 3) = "Unique Served:"
    Worksheets(Sheets.Count).Cells(B_ColCounter + 3, 4) = Unique
    
    ' Clears the column that the unique data was counted on.
    Worksheets(Sheets.Count).Range("G1:G" & (Unique + 1)) = ""
    
    ' Calls UserForm3, which allows the user to pick what sheets to collect data from for the totals page.
    ActivateUserForm3

End Sub

' ActivateUserForm calls UserForm1, which is the primary form for day-today data entry.
Sub ActivateUserForm()
' Display form for inputing a person
    UserForm1.Show
End Sub

' ForReentryONLY calls a userform (Reentry) that is similar to UserForm1, but also includes boxes for Time and Date.
' This should only be used in instances of data loss.
Sub ForReentryONLY()
' Displays Reentry window
    Reentry.Show
End Sub

' NewWorkbookONLY transforms the template excel workbook into a fully functional workbook ready for use at the beginning
' of a new school year. General flow: Format the totals page -> call ButtonMaker, which creates buttons for the reporting
' programs on the totals page -> call UserForm3, which gets the start and end dates of the school year -> saves the 
' workbook with the correct label and file format.
Sub NewWorkbookONLY()
' Generates new workbook to be saved at file address near bottom of sub
    Dim textNumItem, textNumVisits, textUpdate, textWeek, textSemester, textStudents, textCount, textStart, textEnd, textReport As String
    Dim monthIndex, cellIndex, sheetLabel As Integer
    Dim rng1, rng2, rng3, rng4, rng5, rng6, rng7 As Range
    
    If Worksheets(1).Range("A1").Value <> "" Then
        End
    End If
    
    ' Assign Labels
    textNumItem = "Number of Items in "
    textNumVisits = "Number of Visits in "
    textUpdate = "Update this field "
    textWeek = "each week"
    textWeight = "Weight (lb) of Items:"
    textStudents = "Total Students Served "
    textCount = "Unique Count "
    textReport = "Reports "
    textStart = "Start Page:"
    textEnd = "End Page:"
    
    ' Inclusive ranges for borders
    Set rng1 = Range("A3:B15")
    Set rng2 = Range("A17:B17")
    Set rng3 = Range("A20:B31")
    Set rng4 = Range("D3:E15")
    Set rng5 = Range("D17:E17")
    Set rng6 = Range("D23:E31")
    Set rng7 = Range("A33:E34")
    Set rng8 = Range("F23:G28")
    
    ' Labels cell A1
    Worksheets(1).Range("A1").Value = "Campus Cupboard Totals"
    
    ' Set labels for monthly totals
    monthIndex = 8
    For cellIndex = 3 To 7
        Worksheets(1).Range("A" & cellIndex).Value = textNumVisits & MonthName(monthIndex)
        Worksheets(1).Range("D" & cellIndex).Value = textNumItem & MonthName(monthIndex)
        monthIndex = monthIndex + 1
    Next cellIndex
    monthIndex = 1
    For cellIndex = 8 To 15
        Worksheets(1).Range("A" & cellIndex).Value = textNumVisits & MonthName(monthIndex)
        Worksheets(1).Range("D" & cellIndex).Value = textNumItem & MonthName(monthIndex)
        monthIndex = monthIndex + 1
    Next cellIndex
    
    Worksheets(1).Range("A17").Value = "Total " & textNumVisits & "This Year:"
    Worksheets(1).Range("A18").Value = textUpdate & textWeek
    Worksheets(1).Range("A20").Value = textStudents & "Fall " & Year(Date)
    Worksheets(1).Range("A21").Value = textStudents & "Spring " & Year(Date) + 1
    Worksheets(1).Range("A22").Value = textStudents & "Summer " & Year(Date) + 1
    Worksheets(1).Range("A23").Value = "Number of Items Donated to GIH"
    Worksheets(1).Range("A26").Value = "Number of Items Donated to Campus Cupboard"
    Worksheets(1).Range("A29").Value = "Monetary Donations"
    Worksheets(1).Range("A33").Value = "Unique Students Between"
    Worksheets(1).Range("A34").Value = "Unique Students Served to Date"
    
    Worksheets(1).Range("B17").Value = "=SUM(B3:B15)"
    Worksheets(1).Range("B20").Value = "=SUM(B3:B7)"
    Worksheets(1).Range("B21").Value = "=SUM(B8:B12)"
    Worksheets(1).Range("B22").Value = "=SUM(B13:B15)"
    Worksheets(1).Range("B23").Value = "Fall:"
    Worksheets(1).Range("B24").Value = "Spring:"
    Worksheets(1).Range("B25").Value = "Summer:"
    Worksheets(1).Range("B26").Value = "Fall:"
    Worksheets(1).Range("B27").Value = "Spring:"
    Worksheets(1).Range("B28").Value = "Summer:"
    Worksheets(1).Range("B29").Value = "Fall:"
    Worksheets(1).Range("B30").Value = "Spring:"
    Worksheets(1).Range("B31").Value = "Summer:"
    Worksheets(1).Range("B33").Value = "Date:"
    Worksheets(1).Range("B34").Value = "Date:"
    Worksheets(1).Range("D17").Value = "Total Items Distributed this Year"
    
    For cellIndex = 23 To 28
        Worksheets(1).Range("D" & cellIndex).Value = textNumItem & ":"
    Next cellIndex
    
    Worksheets(1).Range("D29").Value = "Amount:"
    Worksheets(1).Range("D30").Value = "Amount:"
    Worksheets(1).Range("D31").Value = "Amount:"
    Worksheets(1).Range("D33").Value = "Number of Students:"
    Worksheets(1).Range("D34").Value = "Number of Students:"
    
    Worksheets(1).Range("E17").Value = "=SUM(E3:E15)"
    
    Worksheets(1).Range("F17").Value = textUpdate & textWeek
    
    For cellIndex = 23 To 28
        Worksheets(1).Range("F" & cellIndex).Value = textWeight
    Next cellIndex
    
    ' If some form of auto sizing is possible maybe use that?
    Worksheets(1).Columns("A").ColumnWidth = 32.29
    Worksheets(1).Columns("B").ColumnWidth = 7.71
    Worksheets(1).Columns("C").ColumnWidth = 15.71
    Worksheets(1).Columns("D").ColumnWidth = 28.71
    Worksheets(1).Columns("F").ColumnWidth = 18.71
    
    ' Set borders color and style
    With rng1.Borders
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng2.Borders
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng3.Borders
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng4.Borders
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng5.Borders
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng6.Borders
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng7.Borders
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' Text style formating
    Worksheets(1).Range("A1").Font.Size = 18
    Worksheets(1).Range("A1").Font.Bold = True
    Worksheets(1).Range("A1").WrapText = True
    Worksheets(1).Range("A18").Font.Color = vbRed
    Worksheets(1).Range("A26").WrapText = True
    Worksheets(1).Range("A33").WrapText = True
    Worksheets(1).Range("F17").Font.Color = vbRed
    
    ' Call ButtonMaker
    ButtonMaker
    
    sheetLabel = Year(Date) - 2000
    Worksheets(1).Name = "Total " & sheetLabel & "-" & sheetLabel + 1 & "(1)"
    
    ActivateUserForm2
    
    NewWeekWorkSheet
    
    Application.DisplayAlerts = False
    ' File save location
    ' ThisWorkbook.SaveAs Filename:="\\empfs1\ShrDirs\Inet\Private\Student Life\Service-Learning\Campus Cupboard\Program\Cupboard Startup\Update Master " & sheetLabel & "-" & sheetLabel + 1 & "USE THIS ONE!!.xlsm", FileFormat:=52
    ' Test Save location
    ' USE THIS ONE -> ThisWorkbook.SaveAs Filename:="C:\Users\Allen\Documents\Normandale\Cupboard\Test Folder\Update Master " & sheetLabel & "-" & sheetLabel + 1 & "USE THIS ONE!!.xlsm", FileFormat:=52
    Application.DisplayAlerts = True

End Sub

' ButtonMaker creates the buttons for the totals page. It's currently only called by the NewWorkBookONLY sub.
Private Sub ButtonMaker()
    Dim Report1, Report2, Unique, NewWork As Button
    Dim Targeter As Range
    
    Set Targeter = Worksheets(1).Cells(3, 7)
    Set Report1 = Worksheets(1).Buttons.Add(Targeter.Left, Targeter.Top, Width:=144, Height:=24)
    With Report1
        .OnAction = "WeeklyReportsP1"
        .Caption = "Weekly Reports"
        .Name = "Weekly Reports"
    End With
                
    Set Targeter = Worksheets(1).Cells(7, 7)
    Set Report2 = Worksheets(1).Buttons.Add(Targeter.Left, Targeter.Top, Width:=144, Height:=24)
    With Report2
        .OnAction = "ActivateUserForm3"
        .Caption = "Reports Maintenance"
        .Name = "Reports Maintenance"
    End With
            
    Set Targeter = Worksheets(1).Cells(9, 7)
    Set Unique = Worksheets(1).Buttons.Add(Targeter.Left, Targeter.Top, Width:=144, Height:=24)
    With Unique
        .OnAction = "ActivateUserForm4"
        .Caption = "Calculate Unique"
        .Name = "Calculate Unique"
    End With
            
    Set Targeter = Worksheets(1).Cells(11, 7)
    Set NewWork = Worksheets(1).Buttons.Add(Targeter.Left, Targeter.Top, Width:=144, Height:=24)
    With NewWork
        .OnAction = "NewWeekWorkSheet"
        .Caption = "Create New Worksheet"
        .Name = "Create New Worksheet"
    End With

End Sub

' Calls UserForm2, which gets the start and end dates of the school year from the user.
Private Sub ActivateUserForm2()
' Displays user form for new school year workbook
    UserForm2.Show
End Sub

' TotalVisits is a counter that starts at 1, if the Student ID in the next row is different than the one in the current row
' then the counter is incremented. Returns the final count.
Private Function TotalVisits(totalRows As Integer)
    Dim currentRow As Integer
    
    If Worksheets(Sheets.Count).Cells(2, 2).Value <> "" Then
        TotalVisits = 1
        For currentRow = 3 To totalRows
            If Worksheets(Sheets.Count).Cells(currentRow, 2).Value <> Worksheets(Sheets.Count).Cells(currentRow - 1, 2).Value Then
                TotalVisits = TotalVisits + 1
            End If
        Next currentRow
    End If
End Function

' ActivateUserForm3 calls UserForm3, which is used for weekly reporting
Private Sub ActivateUserForm3()
' Displays window for weekly reporting
    UserForm3.Show
End Sub

' ActivateUserForm4 calls UserForm4, which is used for calulating unique students in the a range of sheets.
Private Sub ActivateUserForm4()
' Displays form for calculating unique IDs
    UserForm4.Show
End Sub

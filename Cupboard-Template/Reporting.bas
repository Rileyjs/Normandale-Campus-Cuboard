Attribute VB_Name = "Reporting"
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
    ReportingForm.Show

End Sub



' TotalVisits is a counter that starts at 1, if the Student ID in the next row is different than the one in the current row
' then the counter is incremented. Passed totalRows. Returns the final count.
Private Function TotalVisits(totalRows As Integer)
    Dim currentRow As Integer
    
    ' Checks that there is more than just a label in Column B
    If Worksheets(Sheets.Count).Cells(2, 2).Value <> "" Then
        ' Since there is data in B2, we assume that at least 1 student has visted.
        TotalVisits = 1
        
        ' Loop compares the ID in the current row to the one before it, if different TotalVisits is incremented.
        For currentRow = 3 To totalRows
            If Worksheets(Sheets.Count).Cells(currentRow, 2).Value <> Worksheets(Sheets.Count).Cells(currentRow - 1, 2).Value Then
                TotalVisits = TotalVisits + 1
            End If
        Next currentRow
    End If
' This looks weird, and I recommend that you check out the documentation on this, but it appears that TotalVisits acts as a fuction
' and an Integer varible. Meaning that it returns itself as a value.
End Function


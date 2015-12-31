VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportingForm 
   Caption         =   "Reports"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "ReportingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReportingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Begin editing section

' OK button action
Private Sub OLD_CommandButton1_Click()
Dim StartPage, EndPage, monthCode, PageCounter, A_ColCounter, totalItems, totalVisits As Integer

' The +2 accounts for the totals page and that the list index starts at 0, while the page index starts at 1.
StartPage = ComboBox1.listIndex + 2
EndPage = ComboBox2.listIndex + 2
monthCode = ComboBox3.Value
totalItems = 0
totalVisits = 0

' Check for valid month entry.
If 1 > monthCode Or monthCode > 12 Then
    MsgBox ("Please Enter Valid Month Code between 1-12")
    End
End If

For PageCounter = StartPage To EndPage
    ' Finds the number of rows with data
    A_ColCounter = Worksheets(PageCounter).Cells(Rows.Count, 1).End(xlUp).Row
    ' sums the values from the totals section at the bottom of the sheet.
    totalItems = totalItems + Worksheets(PageCounter).Cells(A_ColCounter + 3, 4)
    totalVisits = totalVisits + Worksheets(PageCounter).Cells(A_ColCounter + 2, 4)
Next PageCounter

' This block writes the collected data to the correct month on the totals page.
Select Case monthCode
    Case 8
        If Worksheets(1).Range("B14") = "" Then
            Worksheets(1).Range("B3") = totalVisits
            Worksheets(1).Range("E3") = totalItems
        Else
            Worksheets(1).Range("B15") = totalVisits
            Worksheets(1).Range("E15") = totalItems
        End If
    Case 9
        Worksheets(1).Range("B4") = totalVisits
        Worksheets(1).Range("E4") = totalItems
    Case 10
        Worksheets(1).Range("B5") = totalVisits
        Worksheets(1).Range("E5") = totalItems
    Case 11
        Worksheets(1).Range("B6") = totalVisits
        Worksheets(1).Range("E6") = totalItems
    Case 12
        Worksheets(1).Range("B7") = totalVisits
        Worksheets(1).Range("E7") = totalItems
    Case 1
        Worksheets(1).Range("B8") = totalVisits
        Worksheets(1).Range("E8") = totalItems
    Case 2
        Worksheets(1).Range("B9") = totalVisits
        Worksheets(1).Range("E9") = totalItems
    Case 3
        Worksheets(1).Range("B10") = totalVisits
        Worksheets(1).Range("E10") = totalItems
    Case 4
        Worksheets(1).Range("B11") = totalVisits
        Worksheets(1).Range("E11") = totalItems
    Case 5
        Worksheets(1).Range("B12") = totalVisits
        Worksheets(1).Range("E12") = totalItems
    Case 6
        Worksheets(1).Range("B13") = totalVisits
        Worksheets(1).Range("E13") = totalItems
    Case 7
        Worksheets(1).Range("B14") = totalVisits
        Worksheets(1).Range("E14") = totalItems
End Select

' Leftover from a previous version. Should consider removing.
Worksheets(1).Range("J1:K3") = ""

' Closes the UserForm
Unload Me
End Sub


' Initializes the Userform
Sub UserForm_Initialize()
Dim listIndex As Integer

' Adds 1-12 to the month select list
For listIndex = 0 To 11
    ComboBox3.AddItem listIndex + 1
Next listIndex

' Sets the default to Jan
ComboBox3.Value = 1
End Sub

' Ok button action. This program will check all sheets for rows where the date matches the date entered. When found, it will total and report all data for that
' month and end when it detects a new month.
Private Sub CommandButton1_Click()
    Dim currentPage, currentRow, startRow, currentID, currentMonth, totalRows, totalItems, totalVisits As Integer
    Dim monthCode As String
    
    ' Set target month to value in ComboBox3
    monthCode = ComboBox3.Value
    totalRows = totalItems = totalVisits = 0
    
    ' Verify that monthcode is valid
    If 1 > monthCode Or monthCode > 12 Then
        MsgBox ("Please Enter Valid Month Code between 1-12")
        End
    End If
    
    ' This Do-While structure allows the loop to exit if a new month after the target month is found
    Do
        ' Loop starts at the last page and counts back to the first page. This allows for the least amount of time in the loop, as well as accounting for the
        ' second instance of August found in the Summer Semester.
        For currentPage = Sheets.Count To 2 Step -1
            currentID = 0
            
            ' Sets the starting row to the last row with data in Column B in the sheet
            startRow = Worksheets(currentPage).Cells(Rows.Count, 2).End(xlUp).Row
            
            For currentRow = startRow To 2 Step -1
                
                ' Checks if the current row contains data from the target month
                If monthCode = Month(Worksheets(currentPage).Cells(currentRow, 1)) Then
                    ' Checks if current ID is different from the last seen. If it is, totalVisits is incremented.
                    If currentID <> Worksheets(currentPage).Cells(currentRow, 2) Then
                        totalVisits = totalVisits + 1
                        currentID = Worksheets(currentPage).Cells(currentRow, 2).Value
                    End If
                    
                    ' totalItems is incremented for each row with data pertaining to the target month.
                    totalItems = totalItems + 1
                Else
                    ' This checks that data for the target month has been found. This only executes if the current month is not equal to the target month. This
                    ' will cause the loop to terminate because all relevant data has been collected.
                    If totalItems > 0 Then
                        Exit Do
                    End If
                End If
                                
            Next currentRow
        Next currentPage
    ' If the Exit Do command is not executed and the workbook is completely transversed, the False below should prevent the Do-While from looping again.
    Loop While False
    
    ' This block writes the collected data to the correct month on the totals page.
    Select Case monthCode
        Case 8
            If Worksheets(1).Range("B14") = "" Then
                Worksheets(1).Range("B3") = totalVisits
                Worksheets(1).Range("E3") = totalItems
            Else
                Worksheets(1).Range("B15") = totalVisits
                Worksheets(1).Range("E15") = totalItems
            End If
        Case 9
            Worksheets(1).Range("B4") = totalVisits
            Worksheets(1).Range("E4") = totalItems
        Case 10
            Worksheets(1).Range("B5") = totalVisits
            Worksheets(1).Range("E5") = totalItems
        Case 11
            Worksheets(1).Range("B6") = totalVisits
            Worksheets(1).Range("E6") = totalItems
        Case 12
            Worksheets(1).Range("B7") = totalVisits
            Worksheets(1).Range("E7") = totalItems
        Case 1
            Worksheets(1).Range("B8") = totalVisits
            Worksheets(1).Range("E8") = totalItems
        Case 2
            Worksheets(1).Range("B9") = totalVisits
            Worksheets(1).Range("E9") = totalItems
        Case 3
            Worksheets(1).Range("B10") = totalVisits
            Worksheets(1).Range("E10") = totalItems
        Case 4
            Worksheets(1).Range("B11") = totalVisits
            Worksheets(1).Range("E11") = totalItems
        Case 5
            Worksheets(1).Range("B12") = totalVisits
            Worksheets(1).Range("E12") = totalItems
        Case 6
            Worksheets(1).Range("B13") = totalVisits
            Worksheets(1).Range("E13") = totalItems
        Case 7
            Worksheets(1).Range("B14") = totalVisits
            Worksheets(1).Range("E14") = totalItems
        End Select
        
Unload Me
End Sub

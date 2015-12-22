VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportingForm 
   Caption         =   "UserForm3"
   ClientHeight    =   3180
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
Private Sub CommandButton1_Click()
Dim StartPage, EndPage, MonthCode, PageCounter, A_ColCounter, TotalItems, TotalVisits As Integer

' The +2 accounts for the totals page and that the list index starts at 0, while the page index starts at 1.
StartPage = ComboBox1.listIndex + 2
EndPage = ComboBox2.listIndex + 2
MonthCode = ComboBox3.Value
TotalItems = 0
TotalVisits = 0

' Check for valid month entry.
If 1 > MonthCode Or MonthCode > 12 Then
    MsgBox ("Please Enter Valid Month Code between 1-12")
    End
End If

For PageCounter = StartPage To EndPage
    ' Finds the number of rows with data
    A_ColCounter = Worksheets(PageCounter).Cells(Rows.Count, 1).End(xlUp).Row
    ' sums the values from the totals section at the bottom of the sheet.
    TotalItems = TotalItems + Worksheets(PageCounter).Cells(A_ColCounter + 3, 4)
    TotalVisits = TotalVisits + Worksheets(PageCounter).Cells(A_ColCounter + 2, 4)
Next PageCounter

' This block writes the collected data to the correct month on the totals page.
Select Case MonthCode
    Case 8
        If Worksheets(1).Range("B14") = "" Then
            Worksheets(1).Range("B3") = TotalVisits
            Worksheets(1).Range("E3") = TotalItems
        Else
            Worksheets(1).Range("B15") = TotalVisits
            Worksheets(1).Range("E15") = TotalItems
        End If
    Case 9
        Worksheets(1).Range("B4") = TotalVisits
        Worksheets(1).Range("E4") = TotalItems
    Case 10
        Worksheets(1).Range("B5") = TotalVisits
        Worksheets(1).Range("E5") = TotalItems
    Case 11
        Worksheets(1).Range("B6") = TotalVisits
        Worksheets(1).Range("E6") = TotalItems
    Case 12
        Worksheets(1).Range("B7") = TotalVisits
        Worksheets(1).Range("E7") = TotalItems
    Case 1
        Worksheets(1).Range("B8") = TotalVisits
        Worksheets(1).Range("E8") = TotalItems
    Case 2
        Worksheets(1).Range("B9") = TotalVisits
        Worksheets(1).Range("E9") = TotalItems
    Case 3
        Worksheets(1).Range("B10") = TotalVisits
        Worksheets(1).Range("E10") = TotalItems
    Case 4
        Worksheets(1).Range("B11") = TotalVisits
        Worksheets(1).Range("E11") = TotalItems
    Case 5
        Worksheets(1).Range("B12") = TotalVisits
        Worksheets(1).Range("E12") = TotalItems
    Case 6
        Worksheets(1).Range("B13") = TotalVisits
        Worksheets(1).Range("E13") = TotalItems
    Case 7
        Worksheets(1).Range("B14") = TotalVisits
        Worksheets(1).Range("E14") = TotalItems
End Select

' Leftover from a previous version. Should consider removing.
Worksheets(1).Range("J1:K3") = ""

' Closes the UserForm
Unload Me
End Sub

' Initializes the Userform
Sub UserForm_Initialize()
Dim TargetMonth(11), listIndex, sheetIndex As Integer

' This block fills the first ComboBox with the sheet lables starting with the second sheet (the sheet after the totals page).
sheetIndex = 2
For listIndex = 0 To Sheets.Count - 2
    ComboBox1.AddItem Worksheets(sheetIndex).Name
    sheetIndex = sheetIndex + 1
Next listIndex

' Same as above.
sheetIndex = 2
For listIndex = 0 To Sheets.Count - 2
    ComboBox2.AddItem Worksheets(sheetIndex).Name
    sheetIndex = sheetIndex + 1
Next listIndex

' Adds 1-12 to the month select list
For listIndex = 0 To 11
    ComboBox3.AddItem listIndex + 1
Next listIndex

' Sets defaults to the first item in each list.
ComboBox1.Value = ComboBox1.List(0)
ComboBox2.Value = ComboBox2.List(0)
ComboBox3.Value = 1
End Sub
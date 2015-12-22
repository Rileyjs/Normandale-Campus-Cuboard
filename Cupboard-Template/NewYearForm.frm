VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewYearForm 
   Caption         =   "UserForm2"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   OleObjectBlob   =   "NewYearForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewYearForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Begin editing section

' This UserForm is used by NewWorkbookONLY to receive the start and end dates of the academic year. It then formats the dates and
' places them on the totals page.

' CommandButton1_Click is the OK button action
Private Sub CommandButton1_Click()
Dim suffix1, suffix2 As String
Dim dateTemp As Integer

' This block decides whether a number should have "th", "st", "nd", or "rd" as it's suffix
dateTemp = ComboBox3
Select Case dateTemp
   ' Since 11, 12, and 13 are all oddballs as far as suffix assignment, they're given their own case. Note: 1st, 11th, 21st, 31st.
    Case 11, 12, 13
        suffix1 = "th"
    Case Else
      ' This block attempts to get the date down to a single digit, so that the appropriate suffix can be assigned. This is probably
      ' refinable, should do more testing with less operations.
        dateTemp = ComboBox3 - 10
        Do While dateTemp > 0
        dateTemp = dateTemp - 10
        Loop
        dateTemp = dateTemp + 10

        Select Case dateTemp
            Case 1
                suffix1 = "st"
            Case 2
                suffix1 = "nd"
            Case 3
                suffix1 = "rd"
            Case Else
                suffix1 = "th"
        End Select
End Select
        
' Same as above
dateTemp = ComboBox4
Select Case dateTemp
    Case 11, 12, 13
        suffix2 = "th"
    Case Else
        dateTemp = ComboBox4 - 10
        Do While dateTemp > 0
        dateTemp = dateTemp - 10
        Loop
        dateTemp = dateTemp + 10

        Select Case dateTemp
            Case 1
                suffix2 = "st"
            Case 2
                suffix2 = "nd"
            Case 3
                suffix2 = "rd"
            Case Else
                suffix2 = "th"
        End Select
End Select

' This adds the formated dates into B1 along with the years.
Worksheets(1).Range("B1").Value = Year(Date) & "-" & Year(Date) + 1 & " (" & ComboBox1.Value & ". " & ComboBox3.Value & suffix1 & ", " & Year(Date) & "- " & ComboBox2 & ". " & ComboBox4 & suffix2 & ", " & Year(Date) + 1 & ")"

' Closes the UserForm
Unload Me
End Sub

' Initalizes the UserForm
Sub UserForm_Initialize()
Dim i, dateList(30) As Integer

' Creates a list of month abbriviations
ComboBox1.List = Array("Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul")
ComboBox2.List = Array("Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul")

' Creates a list of dates
For i = 0 To 30
    dateList(i) = i + 1
Next i

' This block assigns date arrays to the ComboBoxes
ComboBox3.List = dateList
ComboBox4.List = dateList

' Sets the default Values for the lists.
ComboBox1.Value = ComboBox1.List(0)
ComboBox2.Value = ComboBox2.List(0)
End Sub

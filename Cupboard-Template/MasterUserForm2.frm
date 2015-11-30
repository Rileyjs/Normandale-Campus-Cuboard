VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   OleObjectBlob   =   "MasterUserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Dim suffix1, suffix2 As String
Dim dateTemp As Integer

dateTemp = ComboBox3
Select Case dateTemp
    Case 11, 12, 13
        suffix1 = "th"
    Case Else
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

Worksheets(1).Range("B1").Value = Year(Date) & "-" & Year(Date) + 1 & " (" & ComboBox1.Value & ". " & ComboBox3.Value & suffix1 & ", " & Year(Date) & "- " & ComboBox2 & ". " & ComboBox4 & suffix2 & ", " & Year(Date) + 1 & ")"
Unload Me
End Sub
Sub UserForm_Initialize()
Dim i, dateList(30) As Integer

ComboBox1.List = Array("Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul")
ComboBox2.List = Array("Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul")

For i = 0 To 30
    dateList(i) = i + 1
Next i

ComboBox3.List = dateList
ComboBox4.List = dateList

ComboBox1.Value = ComboBox1.List(0)
ComboBox2.Value = ComboBox2.List(0)
End Sub

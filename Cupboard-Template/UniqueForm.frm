VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UniqueForm 
   Caption         =   "UserForm4"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "UniqueForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UniqueForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Begin editable section

' OK button action
Private Sub CommandButton1_Click()
    Dim uniqueCount As Long
    Dim dict As New Scripting.Dictionary ' Requires Microsoft Scripting Runtime.  Tools->References->Microsoft Scripting Runtime. Otherwise user-defined type error
    Dim studentID As Long
    Dim row As Integer


    'Sets the range of sheets being used
    If ComboBox1.listIndex <> 0 Then
        StartPage = ComboBox1.listIndex + 1
        EndPage = ComboBox2.listIndex + 1
    Else
        StartPage = 2
        EndPage = Sheets.Count
    End If

    uniqueCount = 0
    
    ' Loops through sheets, runs through B column til empty cell found the moves to new sheet. Checks IDs against dictionary to check for uniqueness
    For PageCounter = StartPage To EndPage
        row = 2
        ' If cell is empty default return value is 0
        studentID = Worksheets(PageCounter).Cells(row, 2).Value

        Do While studentID <> 0
            If Not dict.Exists(studentID) Then
                dict.Add studentID, 1
                uniqueCount = uniqueCount + 1
            End If
            row = row + 1
            studentID = Worksheets(PageCounter).Cells(row, 2).Value
        Loop

    Next PageCounter

    ' Unload the dictionary
     Set dict = Nothing

    ' Feed the data back onto the main sheet
    Select Case ComboBox1.listIndex
        Case 0
            Worksheets(1).Range("C33") = Date
            Worksheets(1).Range("E33") = uniqueCount
        Case Else
            Worksheets(1).Range("A34") = "Unique Students Between " & ComboBox1.Value & " & " & ComboBox2.Value
            Worksheets(1).Range("C34") = Date
            Worksheets(1).Range("E34") = uniqueCount
    End Select

    Unload Me
End Sub

' Initilizes UserForm
Sub UserForm_Initialize()
Dim listIndex, sheetIndex As Integer

ComboBox1.AddItem "All"
sheetIndex = 2
For listIndex = 0 To Sheets.Count - 2
    ComboBox1.AddItem Worksheets(sheetIndex).Name
    sheetIndex = sheetIndex + 1
Next listIndex

ComboBox2.AddItem "All"
sheetIndex = 2
For listIndex = 0 To Sheets.Count - 2
    ComboBox2.AddItem Worksheets(sheetIndex).Name
    sheetIndex = sheetIndex + 1
Next listIndex

ComboBox1.Value = ComboBox1.List(0)
ComboBox2.Value = ComboBox2.List(0)
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "MasterUserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Dim PageCounter, MainPageRows, CurrPageRows, TotalMainRows, StartPage, EndPage, TotalUnique As Integer

If ComboBox1.listIndex <> 0 Then
    StartPage = ComboBox1.listIndex + 1
    EndPage = ComboBox2.listIndex + 1
Else
    StartPage = 2
    EndPage = Sheets.Count
End If

For PageCounter = StartPage To EndPage
    If Worksheets(1).Range("M1") = "" Then MainPageRows = 1 Else: MainPageRows = Worksheets(1).Cells(Rows.Count, 13).End(xlUp).Row
    If Worksheets(PageCounter).Range("B2") = "" Then CurrPageRows = 1 Else: CurrPageRows = Worksheets(PageCounter).Cells(Rows.Count, 2).End(xlUp).Row
    TotalMainRows = MainPageRows + CurrPageRows
    Worksheets(PageCounter).Range("B2:B" & CurrPageRows).Copy Destination:=Worksheets(1).Range("M" & MainPageRows & ":M" & TotalMainRows)
Next PageCounter

Worksheets(1).Range("M1:M50000").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets(1).Range("N1"), Unique:=True
TotalUnique = Worksheets(1).Cells(Rows.Count, 14).End(xlUp).Row
TotalUnique = TotalUnique - 1

Select Case ComboBox1.listIndex
    Case 0
        Worksheets(1).Range("C33") = Date
        Worksheets(1).Range("E33") = TotalUnique
    Case Else
        Worksheets(1).Range("A34") = "Unique Students Between " & ComboBox1.Value & " & " & ComboBox2.Value
        Worksheets(1).Range("C34") = Date
        Worksheets(1).Range("E34") = TotalUnique
End Select

Worksheets(1).Range("M1:M50000") = ""
Worksheets(1).Range("N1:N50000") = ""

Unload Me
End Sub

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


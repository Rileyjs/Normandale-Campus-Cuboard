VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Data Entry"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6975
   OleObjectBlob   =   "MasterUserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Dim currentRows As Integer

If TextBox1.Value <> "" Then
    If 10000 > TextBox1.Value Then
        MsgBox ("The ID number must be greater than 4 digits. Please review your entry.")
    ElseIf TextBox1.Value > 99999999 Then
        MsgBox ("The ID number must be less than 9 digits. Please review your entry.")
    Else
        currentRows = Worksheets(Sheets.Count).Cells(Rows.Count, 2).End(xlUp).Row
        currentRows = currentRows + 1
        If ComboBox1.Value <> "" Then
            Call LineFill(1, currentRows, ComboBox4.Value)
        End If
        
        If ComboBox2.Value <> "" Then
            Call LineFill(2, currentRows, ComboBox5.Value)
        End If
    
        If ComboBox3.Value <> "" Then
            Call LineFill(3, currentRows, ComboBox6.Value)
        End If
        
        TextBox1.Value = ""
        ComboBox1.Value = ""
        ComboBox2.Value = ""
        ComboBox3.Value = ""
        ComboBox4.Value = "1"
        ComboBox5.Value = "1"
        ComboBox6.Value = "1"
        ComboBox7.Value = "1"
        
        Application.DisplayAlerts = False
        ThisWorkbook.Save
        Application.DisplayAlerts = True
    
        TextBox1.SetFocus
    End If
End If

End Sub
Private Sub CommandButton2_Click()
Unload Me
End Sub

Sub UserForm_Initialize()
Dim i, QuantityList(14) As Integer

ComboBox1.List = Array("Canned Fruit", "Canned Soup", "Canned Vegetables", "Cereal", "Cookies", "Chef Boyardee", "Crackers", "Dried Fruit", "Drinks", "Fruit Cups", "Fruit Snacks", "Granola Bars", "Mac & Cheese", "Meat Sticks", "Milk", "Miscellaneous", "Nuts", "Oatmeal", "Popcorn", "Poptarts", "Pudding/Jello", "Ramen Noodles", "Tuna/Chicken/Ham")
ComboBox2.List = Array("Canned Fruit", "Canned Soup", "Canned Vegetables", "Cereal", "Cookies", "Chef Boyardee", "Crackers", "Dried Fruit", "Drinks", "Fruit Cups", "Fruit Snacks", "Granola Bars", "Mac & Cheese", "Meat Sticks", "Milk", "Miscellaneous", "Nuts", "Oatmeal", "Popcorn", "Poptarts", "Pudding/Jello", "Ramen Noodles", "Tuna/Chicken/Ham")
ComboBox3.List = Array("Canned Fruit", "Canned Soup", "Canned Vegetables", "Cereal", "Cookies", "Chef Boyardee", "Crackers", "Dried Fruit", "Drinks", "Fruit Cups", "Fruit Snacks", "Granola Bars", "Mac & Cheese", "Meat Sticks", "Milk", "Miscellaneous", "Nuts", "Oatmeal", "Popcorn", "Poptarts", "Pudding/Jello", "Ramen Noodles", "Tuna/Chicken/Ham")

For i = 1 To 15
    QuantityList(i - 1) = i
Next i

ComboBox4.List = QuantityList
ComboBox5.List = QuantityList
ComboBox6.List = QuantityList
ComboBox7.List = Array("1", "2")
ComboBox4.Value = "1"
ComboBox5.Value = "1"
ComboBox6.Value = "1"
ComboBox7.Value = "1"
End Sub

Private Sub LineFill(currentBox As Integer, currentRows As Integer, itemCount As Integer)
        
For entryCounter = 1 To itemCount
    Worksheets(Sheets.Count).Range("A" & currentRows).Value = Date
    Worksheets(Sheets.Count).Range("B" & currentRows).Value = TextBox1.Value
    
    Select Case currentBox
        Case 1
            Worksheets(Sheets.Count).Range("C" & currentRows).Value = ComboBox1.Value
        Case 2
            Worksheets(Sheets.Count).Range("C" & currentRows).Value = ComboBox2.Value
        Case 3
            Worksheets(Sheets.Count).Range("C" & currentRows).Value = ComboBox3.Value
    End Select
    
    Worksheets(Sheets.Count).Range("D" & currentRows).Value = ComboBox7.Value
    Worksheets(Sheets.Count).Range("E" & currentRows).Value = Format(Time, "h:mm AM/PM")
    currentRows = currentRows + 1
    Next entryCounter
End Sub

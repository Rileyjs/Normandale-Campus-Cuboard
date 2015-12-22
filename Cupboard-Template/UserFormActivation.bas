Attribute VB_Name = "UserFormActivation"
' ActivateUserForm calls UserForm1, which is the primary form for day-today data entry.
Sub ActivateUserForm()
' Display form for inputing a person
    DataEntry.Show
End Sub

' ForReentryONLY calls a userform (Reentry) that is similar to UserForm1, but also includes boxes for Time and Date.
' This should only be used in instances of data loss.
Sub ForReentryONLY()
' Displays Reentry window
    Reentry.Show
End Sub

' ActivateUserForm3 calls ReportingForm, which is used to do weekly reporting. This sub is only used by the Reports Maintanence button.
Private Sub ActivateUserForm3()
    ReportingForm.Show
End Sub

' ActivateUserForm4 calls UserForm4, which is used for calulating unique students in the a range of sheets.
Private Sub ActivateUserForm4()
' Displays form for calculating unique IDs
    UniqueForm.Show
End Sub


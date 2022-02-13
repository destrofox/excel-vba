Attribute VB_Name = "Filters"
Option Explicit

Sub clearFilters()

On Error Resume Next
    ActiveSheet.ShowAllData

End Sub

Sub setFilterOnCurrentColumn()

' Variable list
Dim lastRow As Integer
Dim activeCol As Integer
Dim userInput As String
Dim filterFieldText As String

' Get sheet information
activeCol = ActiveCell.Column
lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

' The first row is assumed to be the title one
filterFieldText = ActiveSheet.Cells(1, activeCol).Value

' Get the filter token from the user
userInput = InputBox("Enter filter token for """ & filterFieldText & """.", "AutoFilter")

' Filter current column with user input, any previous filter is maintained
If Not Trim(userInput & vbNullString) = vbNullString Then
    ActiveSheet.Range(ActiveSheet.Cells(1, activeCol), ActiveSheet.Cells(lastRow, activeCol)).AutoFilter _
        Field:=activeCol, Criteria1:="*" & userInput & "*"
End If

End Sub

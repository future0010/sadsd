Sub GetEmployeeDetails()
    ' Declare variables
    Dim empID As Long
    Dim empName As String
    Dim managerID As Long
 
    ' Get the employee ID from the second sheet
    empID = InputBox("Enter employee ID:")
 
    ' Find the employee ID in the first sheet and get the corresponding name and manager ID
    With ThisWorkbook.Worksheets("Sheet1")
        For rowNum = 2 To .Range("A" & .Rows.Count).End(xlUp).Row
            If .Cells(rowNum, 1).Value = empID Then
                empName = .Cells(rowNum, 2).Value
                managerID = .Cells(rowNum, 3).Value
                place = .Cells(rowNum, 4).Value
                Exit For
            End If
        Next rowNum
    End With
 
    ' Write the employee name and manager ID to the second sheet
    ActiveSheet.Cells(ActiveCell.Row, 1).Value = empID
    ActiveSheet.Cells(ActiveCell.Row, 2).Value = empName
    ActiveSheet.Cells(ActiveCell.Row, 3).Value = managerID
    ActiveSheet.Cells(ActiveCell.Row, 4).Value = place
End Sub


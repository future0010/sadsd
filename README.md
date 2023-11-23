Sub GetEmployeeDetails()
 
    ' Declare variables
    Dim empID As Long
    Dim empName As String
    Dim managerID As Long
    Dim employeeIDs As String
    Dim place As String
 
    ' Get the employee ID from the second sheet
    empID = InputBox("Enter employee ID:")
 
    ' Initialize the employeeIDs variable
    employeeIDs = ""
    
    xID = ""
 
    ' Find the employee ID in the first sheet and get the corresponding name, manager ID, and place
    With ThisWorkbook.Worksheets("Sheet1")
        For rowNum = 2 To .Range("A" & .Rows.Count).End(xlUp).Row
 
                ' Update the empName, managerID, and place variables
                empName = .Cells(rowNum, 2).Value
                managerID = .Cells(rowNum, 3).Value
                place = .Cells(rowNum, 4).Value
 
                ' Exit the loop to avoid further searching
                Exit For
        Next rowNum
        

    End With
    
    With ThisWorkbook.Worksheets("Sheet1")
            For rowNum = 2 To .Range("C" & .Rows.Count).End(xlUp).Row
            If .Cells(rowNum, 1).Value = empID Then
                ' Append the employee ID to the employeeIDs variable
                xID = xID & .Cells(rowNum, 1).Value & ", "
                Exit For
            End If
        Next rowNum
    
    End With
 
    ' Write the employee name, manager ID, employee IDs, and place to the second sheet
    ActiveSheet.Cells(ActiveCell.Row, 1).Value = empID
    ActiveSheet.Cells(ActiveCell.Row, 2).Value = empName
    ActiveSheet.Cells(ActiveCell.Row, 3).Value = managerID
    ActiveSheet.Cells(ActiveCell.Row, 4).Value = place
    ActiveSheet.Cells(ActiveCell.Row, 5).Value = xID
 
End Sub

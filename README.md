Sub GetEmployeeDetails()

    ' Declare variables
    Dim empID As Long
    Dim empName As String
    Dim managerID As Long
    Dim employeeIDs As String
    Dim place As String
    Dim Name As String
    Dim i As Integer

    ' Get the employee ID from the second sheet
    empID = InputBox("Enter employee ID:")

    ' Initialize the employeeIDs variable
    employeeIDs = ""
    Name = ""
    i = 0

    ' Find the employee ID in the first sheet and get the corresponding name, manager ID, and place
    With ThisWorkbook.Worksheets("Sheet1")
        For rowNum = 2 To .Range("A" & .Rows.Count).End(xlUp).Row
            If .Cells(rowNum, 1).Value = empID Then

                ' Update the empName, managerID, and place variables
                empName = .Cells(rowNum, 2).Value
                managerID = .Cells(rowNum, 3).Value
                place = .Cells(rowNum, 4).Value

                ' Exit the loop to avoid further searching
                Exit For
            End If
        Next rowNum
    End With

    ' Write the employee name, manager ID, employee IDs, and place to the second sheet
    ActiveSheet.Cells(ActiveCell.Row, 1).Value = empID
    ActiveSheet.Cells(ActiveCell.Row, 2).Value = empName
    ActiveSheet.Cells(ActiveCell.Row, 3).Value = managerID
    ActiveSheet.Cells(ActiveCell.Row, 4).Value = place

    ' Loop through reportees and write their names to separate cells
    With ThisWorkbook.Worksheets("Sheet1")
        For rowNum = 2 To .Range("C" & .Rows.Count).End(xlUp).Row
            If .Cells(rowNum, 3).Value = empID Then

                'Set employeeIDs to reportees ID
                employeeIDs = employeeIDs & .Cells(rowNum, 1).Value & ", "

                'Set Name to reportees Name
                Name = Name & .Cells(rowNum, 2).Value & ", "

                i = i + 1

                ' Write reportee name to a separate cell
                ActiveSheet.Cells(ActiveCell.Row + i, 6).Value = .Cells(rowNum, 2).Value

                If rowNum = .Rows.Count Then
                    Exit For
                End If

            End If
        Next rowNum
    End With

End Sub

Sub GetEmployeeDetails()
    ' Declare variables
    Dim empID As Long
    Dim empName As String
    Dim managerID As Long
    Dim employeeIDs As String
    Dim place As String
    Dim Name As String

    ' Get the employee ID from the second sheet
    empID = InputBox("Enter employee ID:")

    ' Initialize the employeeIDs variable
    employeeIDs = ""
    Name = ""

    ' Reset the color of all reportees to "no fill"
    With ThisWorkbook.Worksheets("Sheet1")
        For rowNum = 2 To .Range("A" & .Rows.Count).End(xlUp).Row
            .Cells(rowNum, 1).Interior.Color = xlNone
            .Cells(rowNum, 2).Interior.Color = xlNone
            .Cells(rowNum, 3).Interior.Color = xlNone
            .Cells(rowNum, 4).Interior.Color = xlNone
        Next rowNum
    End With

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

    With ThisWorkbook.Worksheets("Sheet1")
        For rowNum = 2 To .Range("C" & .Rows.Count).End(xlUp).Row
            If .Cells(rowNum, 3).Value = empID Then
                ' Set employeeIDs to reportees ID
                employeeIDs = employeeIDs & .Cells(rowNum, 1).Value & ", "
                ' Color the reportees cells
                .Cells(rowNum, 1).Interior.Color = vbYellow
                .Cells(rowNum, 2).Interior.Color = vbYellow
                .Cells(rowNum, 3).Interior.Color = vbYellow
                .Cells(rowNum, 4).Interior.Color = vbYellow
                ' Set Name to reportees Name
                Name = Name & .Cells(rowNum, 2).Value & ", "
                If rowNum = .Rows.Count Then
                    ' Exit the loop after reaching the last row
                    Exit For
                End If
            End If
        Next rowNum
    End With

    ' Write the reportees ID and name
    ActiveSheet.Cells(ActiveCell.Row, 5).Value = employeeIDs
    ActiveSheet.Cells(ActiveCell.Row, 6).Value = Name
End Sub

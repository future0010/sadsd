Sub GetEmployeeDetails()

' Declare variables
Dim empID As Long
Dim empName As String
Dim managerID As Long
Dim employeeIDs As String
Dim place As String
Dim Name As String
Dim rowNum As Long

' Get the employee ID from the second sheet
empID = InputBox("Enter employee ID:")

' Initialize the employeeIDs variable
employeeIDs = ""
Name = ""

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

' Find the first empty row in the second sheet
rowNum = 2
While ActiveSheet.Cells(rowNum, 1).Value <> ""
    rowNum = rowNum + 1
Wend

' Write the employee ID, employee name, manager ID, employee IDs, and place to the second sheet
ActiveSheet.Cells(rowNum, 1).Value = empID
ActiveSheet.Cells(rowNum, 2).Value = empName
ActiveSheet.Cells(rowNum, 3).Value = managerID
ActiveSheet.Cells(rowNum, 4).Value = place

' Write the reportees ID
ActiveSheet.Cells(rowNum, 5).Value = employeeIDs
ActiveSheet.Cells(rowNum, 6).Value = Name

End Sub

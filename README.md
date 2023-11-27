Sub GetEmployeeDetails()

' Declare variables
Dim empID As Long
Dim empName As String
Dim managerID As Long
Dim reporteeIDs As Variant
Dim reporteeNames As Variant
Dim place As String
Dim rowNum As Long

' Get the employee ID from the second sheet
empID = InputBox("Enter employee ID:")

' Initialize the reporteeIDs and reporteeNames variables
reporteeIDs = Array()
reporteeNames = Array()

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

      ' Find the reportee IDs and names
      For i = 2 To .Range("C" & .Rows.Count).End(xlUp).Row
        If .Cells(i, 3).Value = empID Then
          ' Add reportee ID to the list
          ReDim Preserve reporteeIDs(1 To UBound(reporteeIDs) + 1)
          reporteeIDs(UBound(reporteeIDs)) = .Cells(i, 1).Value

          ' Add reportee name to the list
          ReDim Preserve reporteeNames(1 To UBound(reporteeNames) + 1)
          reporteeNames(UBound(reporteeNames)) = .Cells(i, 2).Value

          ' Color the reportees cells
          .Cells(i, 1).Interior.Color = vbYellow
          .Cells(i, 2).Interior.Color = vbYellow
          .Cells(i, 3).Interior.Color = vbYellow
          .Cells(i, 4).Interior.Color = vbYellow
        End If
      Next i

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

' Write the reportee IDs and names to separate cells
rowNum = 5
For i = 0 To UBound(reporteeIDs)
  ActiveSheet.Cells(ActiveCell.Row, rowNum).Value = reporteeIDs(i)
  rowNum = rowNum + 1
Next i

rowNum = 6
For i = 0 To UBound(reporteeNames)
  ActiveSheet.Cells(ActiveCell.Row, rowNum).Value = reporteeNames(i)
  rowNum = rowNum + 1
Next i

End Sub

Attribute VB_Name = "Module1"
Sub ClearExpenses()
    'Activate the "Expenses" sheet
    Sheets("Expenses").Activate
    
    'Select the range
    Range("A2:E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    'Clear the selected range
    Selection.ClearContents
    
    'Return to the "Output" sheet
    Sheets("Output").Activate
End Sub
Sub ClearIncome()
    'Activate the "Income" sheet
    Sheets("Incomes").Activate
    
    'Select the range
    Range("A2:E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    'Clear the selected range
    Selection.ClearContents
    
    'Return to the "Output" sheet
    Sheets("Output").Activate
End Sub
Sub ClearGoals()

    'Select the range
    Range("A2:G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    'Clear the selected range
    Selection.ClearContents

End Sub
Sub ClearOutputSheet()
    Dim ws As Worksheet
    
    ' Reference the "Output" sheet
    Set ws = ThisWorkbook.Sheets("Output")
    
    ' Clear specific cells and columns
    Application.ScreenUpdating = False
    ws.Range("A2, A4").ClearContents    ' Clear cells A2 and A4
    ws.Range("D2:M" & ws.Rows.Count).ClearContents ' Clear columns D to M starting from row 2
    Application.ScreenUpdating = True
End Sub


Sub ClearGraph()
    On Error GoTo ErrorHandler

    ' Declare worksheet variable
    Dim wsOutput As Worksheet
    Dim chartObj As chartObject

    ' Set the worksheet
    Set wsOutput = ThisWorkbook.Sheets("Output")

    ' Loop through all chart objects in the worksheet and delete them
    For Each chartObj In wsOutput.ChartObjects
        chartObj.Delete
    Next chartObj

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.description
End Sub

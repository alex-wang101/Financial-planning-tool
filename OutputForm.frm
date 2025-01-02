VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OutputForm 
   Caption         =   "Output"
   ClientHeight    =   8040
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   10056
   OleObjectBlob   =   "OutputForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OutputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SubmitBtn_Click()

    ' Variables declaration
    Dim startDate As Date
    Dim endDate As Date
    Dim startDay As String, startMonth As String, startYear As String
    Dim endDay As String, endMonth As String, endYear As String
    
    ' Get the start date from the textboxes
    startDay = txtDay1.Value
    startMonth = txtMonth1.Value
    startYear = txtYear1.Value
    
    ' Get the end date from the textboxes
    endDay = txtDay2.Value
    endMonth = txtMonth2.Value
    endYear = txtYear2.Value
    
    ' Validate if all fields are filled
    If startDay = "" Or startMonth = "" Or startYear = "" Or endDay = "" Or endMonth = "" Or endYear = "" Then
        MsgBox "Please fill in all date fields.", vbExclamation
        Exit Sub
    End If
    
    ' Create the date variables
    startDate = DateSerial(startYear, startMonth, startDay)
    endDate = DateSerial(endYear, endMonth, endDay)
    
    ' Check if the start date is before the end date
    If startDate > endDate Then
        MsgBox "Start date cannot be later than the end date.", vbExclamation
        Exit Sub
    End If
    
    ' Print the start date on A2 and end date on A4
    With ThisWorkbook.Sheets("Output")
        .Cells(2, 1).Value = startDate
        .Cells(2, 1).NumberFormat = "yyyy-mm-dd"
        .Cells(4, 1).Value = endDate
        .Cells(4, 1).NumberFormat = "yyyy-mm-dd"
    End With

    ' Call the procedure to retrieve the data based on the dates
    RetrieveData startDate, endDate

    ' Highlight extremes 
    HighlightExtremes

    Exit Sub

End Sub


Sub RetrieveData(startDate As Date, endDate As Date)
    ' Declare worksheet variables
    Dim wsIncomes As Worksheet
    Dim wsExpenses As Worksheet
    Dim wsOutput As Worksheet
    Dim incomeLastRow As Long
    Dim expenseLastRow As Long
    Dim outputRow As Long
    Dim i As Long, j As Long
    Dim incomeDate As Date
    Dim expenseDate As Date
    Dim totalIncome As Double
    Dim totalExpenses As Double

    ' Set the worksheets
    Set wsIncomes = ThisWorkbook.Sheets("Incomes")
    Set wsExpenses = ThisWorkbook.Sheets("Expenses")
    Set wsOutput = ThisWorkbook.Sheets("Output") 

    ' Clear the previous output (if any)
    wsOutput.Range("D2:M1000").ClearContents

    ' Get the last row in the Incomes and Expenses sheets
    incomeLastRow = wsIncomes.Cells(wsIncomes.Rows.Count, "A").End(xlUp).row
    expenseLastRow = wsExpenses.Cells(wsExpenses.Rows.Count, "A").End(xlUp).row

    ' Initialize output row (where data will start being displayed)
    outputRow = 2

    ' Initialize totals
    totalIncome = 0
    totalExpenses = 0

    ' Loop through the Incomes sheet and filter by date
    For i = 2 To incomeLastRow
        incomeDate = wsIncomes.Cells(i, 1).Value ' Assuming date is in column A
        If incomeDate >= startDate And incomeDate <= endDate Then
            ' Sum total income
            totalIncome = totalIncome + wsIncomes.Cells(i, 2).Value ' Total Income Value
            ' Output income data to the output table
            wsOutput.Cells(outputRow, 4).Value = incomeDate ' Date
            wsOutput.Cells(outputRow, 6).Value = wsIncomes.Cells(i, 2).Value ' Total Income Value
            wsOutput.Cells(outputRow, 7).Value = wsIncomes.Cells(i, 3).Value ' Category
            wsOutput.Cells(outputRow, 8).Value = wsIncomes.Cells(i, 4).Value ' Description
            outputRow = outputRow + 1
        End If
    Next i

    ' Place total income in E2
    wsOutput.Cells(2, 5).Value = totalIncome

    ' Reset output row for expenses
    outputRow = 2

    ' Loop through the Expenses sheet and filter by date
    For j = 2 To expenseLastRow
        On Error Resume Next ' Skip rows with invalid dates
        expenseDate = CDate(wsExpenses.Cells(j, 1).Value) ' Convert to date
        On Error GoTo 0 ' Resume normal error handling
        If expenseDate >= startDate And expenseDate <= endDate Then
            ' Sum total expenses
            totalExpenses = totalExpenses + wsExpenses.Cells(j, 2).Value 
            ' Output expense data to the output table
            wsOutput.Cells(outputRow, 13).Value = expenseDate ' Date
            wsOutput.Cells(outputRow, 10).Value = wsExpenses.Cells(j, 2).Value 
            wsOutput.Cells(outputRow, 11).Value = wsExpenses.Cells(j, 3).Value 
            wsOutput.Cells(outputRow, 12).Value = wsExpenses.Cells(j, 4).Value 
            outputRow = outputRow + 1
        End If
    Next j

    ' Place total expenses in I2
    wsOutput.Cells(2, 9).Value = totalExpenses

    MsgBox "Data retrieved successfully!"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.description
End Sub


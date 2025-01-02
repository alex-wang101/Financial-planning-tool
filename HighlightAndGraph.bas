Attribute VB_Name = "Module5"
Sub CompareAndGraphPivotTables_Debugged()
    Dim ws As Worksheet
    Dim ptExpenses As pivotTable
    Dim ptIncome As pivotTable
    Dim expensesRange As Range
    Dim incomeRange As Range
    Dim chartObj As chartObject
    Dim chartTitle As String
    Dim expenseValues() As Variant
    Dim incomeValues() As Variant
    Dim monthCategories() As Variant
    Dim colCount As Integer
    Dim i As Integer
    
    ' Define worksheet and pivot tables
    Set ws = ThisWorkbook.Sheets("Output")
    Set ptExpenses = ws.PivotTables("ExpensesPivot")
    Set ptIncome = ws.PivotTables("IncomePivot")
    
    ' Ensure pivot tables exist
    If ptExpenses Is Nothing Or ptIncome Is Nothing Then
        MsgBox "Both 'ExpensesPivot' and 'IncomePivot' must exist on the 'Output' sheet.", vbExclamation
        Exit Sub
    End If

    ' Get data ranges
    Set expensesRange = ptExpenses.DataBodyRange
    Set incomeRange = ptIncome.DataBodyRange

    ' Check if ranges exist
    If expensesRange Is Nothing Or incomeRange Is Nothing Then
        MsgBox "One or both pivot tables do not have data.", vbExclamation
        Exit Sub
    End If
    
    ' Initialize arrays to hold data
    colCount = ptExpenses.ColumnRange.Columns.Count
    ReDim monthCategories(1 To colCount)
    ReDim expenseValues(1 To colCount)
    ReDim incomeValues(1 To colCount)
    
    ' Collect data from pivot tables
    Dim monthOrCategory As String
    Dim expenseValue As Double
    Dim incomeValue As Double
    
    ' Gather column headers (months)
    For i = 1 To colCount
        monthCategories(i) = ptExpenses.ColumnRange.Cells(1, i).Value
    Next i
    
    ' Gather data for each month
    For i = 1 To colCount
        expenseValue = Application.Sum(ptExpenses.DataBodyRange.Columns(i))
        incomeValue = Application.Sum(ptIncome.DataBodyRange.Columns(i))
        expenseValues(i) = expenseValue
        incomeValues(i) = incomeValue
    Next i
    
    ' Clear previous chart if it exists
    On Error Resume Next
    ws.ChartObjects("PivotComparisonChart").Delete
    On Error GoTo 0
    
    ' Create the chart
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Range("E40").Left, _
                                       Top:=ws.Range("E40").Top, _
                                       Width:=400, Height:=250)
    chartObj.Name = "PivotComparisonChart"
    If Not chartObj Is Nothing Then
        With chartObj.chart
            .SeriesCollection.NewSeries
            .SeriesCollection(1).Name = "Expenses"
            .SeriesCollection(1).Values = expenseValues
            .SeriesCollection(1).XValues = monthCategories
            
            .SeriesCollection.NewSeries
            .SeriesCollection(2).Name = "Income"
            .SeriesCollection(2).Values = incomeValues
            .SeriesCollection(2).XValues = monthCategories
            
            .ChartType = xlColumnClustered
            
            ' Set chart title and format
            chartTitle = "Comparison of Expenses and Income"
            .HasTitle = True
            .chartTitle.Text = chartTitle
            .chartTitle.Font.Size = 12
            .chartTitle.Font.Bold = True
            
            ' Add axis titles
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = "Month"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "Amount ($)"
            
            ' Format legend
            .HasLegend = True
            .Legend.Position = xlBottom
        End With
    End If

    MsgBox "Comparison graph created successfully in 'Output' sheet!", vbInformation
End Sub


Sub HighlightExtremes()
    Dim ws As Worksheet
    Dim incomeRange As Range, expenseRange As Range
    Dim highestIncomeCell As Range, lowestIncomeCell As Range
    Dim highestExpenseCell As Range, lowestExpenseCell As Range
    Dim maxIncome As Double, minIncome As Double
    Dim maxExpense As Double, minExpense As Double
    Dim cell As Range

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Output")

    ' Define the income and expense ranges
    Set incomeRange = ws.Range("F2:F4") ' Adjust the range as needed
    Set expenseRange = ws.Range("J2:J12") ' Adjust the range as needed

    ' Initialize variables for income
    maxIncome = Application.WorksheetFunction.Max(incomeRange)
    minIncome = Application.WorksheetFunction.Min(incomeRange)

    ' Find the cells with the highest and lowest income
    For Each cell In incomeRange
        If cell.Value = maxIncome Then Set highestIncomeCell = cell
        If cell.Value = minIncome Then Set lowestIncomeCell = cell
    Next cell

    ' Initialize variables for expense
    maxExpense = Application.WorksheetFunction.Max(expenseRange)
    minExpense = Application.WorksheetFunction.Min(expenseRange)

    ' Find the cells with the highest and lowest expense
    For Each cell In expenseRange
        If cell.Value = maxExpense Then Set highestExpenseCell = cell
        If cell.Value = minExpense Then Set lowestExpenseCell = cell
    Next cell

    ' Apply formatting
    If Not highestIncomeCell Is Nothing Then
        highestIncomeCell.Interior.Color = RGB(0, 255, 0) ' Green
    End If
    If Not lowestIncomeCell Is Nothing Then
        lowestIncomeCell.Interior.Color = RGB(255, 0, 0) ' Red
    End If
    If Not highestExpenseCell Is Nothing Then
        highestExpenseCell.Interior.Color = RGB(0, 255, 0) ' Green
    End If
    If Not lowestExpenseCell Is Nothing Then
        lowestExpenseCell.Interior.Color = RGB(255, 0, 0) ' Red
    End If

End Sub



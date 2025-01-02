Attribute VB_Name = "Module4"
Sub CreateExpensesBySourceBarGraphCompact()
    Dim ws As Worksheet
    Dim pt As pivotTable
    Dim chartObj As chartObject
    Dim dataRange As Range
    Dim chartTitle As String
    
    ' Reference the worksheet containing the pivot table
    Set ws = ThisWorkbook.Sheets("Output")
    
    ' Reference the pivot table
    On Error Resume Next
    Set pt = ws.PivotTables("ExpensesPivot")
    If pt Is Nothing Then
        MsgBox "Pivot table 'ExpensesPivot' not found on the 'Output' sheet.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Get the data range of the pivot table
    Set dataRange = pt.TableRange2
    If dataRange Is Nothing Then
        MsgBox "Unable to determine the data range of the pivot table.", vbExclamation
        Exit Sub
    End If
    
    ' Delete existing chart if present
    For Each chartObj In ws.ChartObjects
        If chartObj.Name = "ExpensesBarGraph" Then chartObj.Delete
    Next chartObj
    
    ' Create a new chart
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(40, 23).Left, _
                                       Top:=ws.Cells(40, 23).Top, _
                                       Width:=400, Height:=250)
    chartObj.Name = "ExpensesBarGraph"
    
    ' Ensure the chart is properly initialized before referencing it
    If Not chartObj Is Nothing Then
        With chartObj.chart
            .SetSourceData Source:=dataRange
            .ChartType = xlBarStacked
            
            ' Set chart title and format
            chartTitle = "Monthly Expenses Distribution by Source"
            .HasTitle = True
            .chartTitle.Text = chartTitle
            .chartTitle.Font.Size = 12
            .chartTitle.Font.Bold = True
            
            ' Format axes
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = "Sources of Expenses"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "Total Expenses ($)"
            
            ' Add legend
            .HasLegend = True
            .Legend.Position = xlBottom
            
            ' Enable data labels for better visualization
            .ApplyDataLabels
        End With
    Else
        MsgBox "Failed to create chart.", vbExclamation
        Exit Sub
    End If
End Sub


Sub RefreshExpensesPivot()
    Dim pt As pivotTable
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim dataRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' Define worksheets
    Set wsData = ThisWorkbook.Sheets("Expenses")
    Set wsPivot = ThisWorkbook.Sheets("Output")
    
    ' Find the used range of the "Expenses" sheet
    With wsData
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        Set dataRange = .Range(.Cells(1, 1), .Cells(lastRow, lastCol))
    End With

    ' Reference the pivot table on the "Output" sheet
    On Error Resume Next
    Set pt = wsPivot.PivotTables("ExpensesPivot")
    If pt Is Nothing Then
        MsgBox "Pivot table 'ExpensesPivot' not found on the 'Output' sheet.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Refresh the pivot table
    pt.RefreshTable
    MsgBox "The pivot table 'ExpensesPivot' has been refreshed.", vbInformation
End Sub



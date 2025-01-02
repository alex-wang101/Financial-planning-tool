VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FinancialAdvice 
   Caption         =   "UserForm1"
   ClientHeight    =   2715
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8940.001
   OleObjectBlob   =   "FinancialAdvice.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FinancialAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim novemberIncome As Double
    Dim dateValue As Date
    Dim incomeValue As Double
    Dim savings As Double
    Dim necessities As Double
    Dim freeSpendings As Double
    
    ' Set the worksheet containing the data
    Set ws = ThisWorkbook.Sheets("Output") 

    ' Find the last row in column D (Date column)
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).row

    ' Initialize total November income
    novemberIncome = 0

    ' Loop through all rows of data
    For currentRow = 2 To lastRow ' Assuming headers are in row 1
        dateValue = ws.Cells(currentRow, "D").Value 
        incomeValue = ws.Cells(currentRow, "F").Value 

        ' Check if the date falls in November
        If Month(dateValue) = 11 Then
            novemberIncome = novemberIncome + incomeValue
        End If
    Next currentRow

    ' Check if income in November
    If novemberIncome > 0 Then
        ' Perform financial calculations
        savings = novemberIncome * 0.7
        necessities = novemberIncome * 0.2
        freeSpendings = novemberIncome * 0.1

        ' Display results
        MsgBox "November Income: $" & Format(novemberIncome, "#,##0.00") & vbCrLf & _
               "$" & Format(savings, "#,##0.00") & " should go into your savings" & vbCrLf & _
               "$" & Format(necessities, "#,##0.00") & " should go into your necessities" & vbCrLf & _
               "$" & Format(freeSpendings, "#,##0.00") & " should go into your free spendings", _
               vbInformation, "November Income Distribution"
    Else
        ' Handle for no November income
        MsgBox "No income data found for November in the table.", vbExclamation, "No Data"
    End If
End Sub


Private Sub GoalTracking_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Goals")
    
    Dim row As Long
    Dim goal As String
    Dim isOnTrack As Boolean
    Dim goalTime As Double
    Dim startPos As Long
    Dim endPos As Long
    Dim timeString As String
    
    isOnTrack = True
    
    ' Loop through each goal in column G 
    For row = 2 To ws.Cells(ws.Rows.Count, "G").End(xlUp).row
        goal = ws.Cells(row, 7).Value
        
        ' Check if the goal time 
        If InStr(goal, "years") > 0 Then
            startPos = InStr(goal, "It will take ") + Len("It will take ")
            endPos = InStr(goal, " years to reach this goal")
            timeString = Mid(goal, startPos, endPos - startPos)
            
            If IsNumeric(timeString) Then
                goalTime = CDbl(timeString)
                If goalTime > 1 Then
                    isOnTrack = False
                    MsgBox "Focus on the goal in row " & row & " that needs more than a year to complete.", vbExclamation, "Focus on Long-term Goal"
                    Exit Sub
                End If
            Else
                MsgBox "The goal time in row " & row & " is not a valid number.", vbExclamation, "Invalid Goal Time"
                Exit Sub
            End If
        End If
    Next row
    
    ' If all goals are less than a year
    If isOnTrack Then
        MsgBox "You are on track with your goals!", vbInformation, "On Track"
    End If
End Sub
Private Sub MonthlySpendingAdvice_Click()
    Dim ws As Worksheet
    Dim incomePivot As pivotTable
    Dim novemberIncome As Double
    Dim savings As Double
    Dim necessities As Double
    Dim freeSpendings As Double
    
    ' Set the worksheet containing the pivot table
    Set ws = ThisWorkbook.Sheets("Output")

    ' Set the IncomePivot table
    Set incomePivot = ws.PivotTables("IncomePivot") 

    ' Retrieve the November income using GetPivotData
    On Error Resume Next
    novemberIncome = incomePivot.GetPivotData("Value", "Category", "Nov")
    On Error GoTo 0

    ' Validate the retrieved value
    If IsNumeric(novemberIncome) And novemberIncome > 0 Then
        ' Perform calculations
        savings = novemberIncome * 0.7
        necessities = novemberIncome * 0.2
        freeSpendings = novemberIncome * 0.1

        ' Display the results
        MsgBox "November Income: $" & Format(novemberIncome, "#,##0.00") & vbCrLf & _
               "$" & Format(savings, "#,##0.00") & " should go into your savings" & vbCrLf & _
               "$" & Format(necessities, "#,##0.00") & " should go into your necessities" & vbCrLf & _
               "$" & Format(freeSpendings, "#,##0.00") & " should go into your free spendings", _
               vbInformation, "November Income Distribution"
    Else
        ' Handle invalid or missing data
        MsgBox "Unable to retrieve November income from the IncomePivot table. Ensure the pivot table is properly configured and has data for November.", _
               vbExclamation, "Invalid Data"
    End If
End Sub





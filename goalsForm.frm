VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} goalsForm 
   Caption         =   "UserForm1"
   ClientHeight    =   5388
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7500
   OleObjectBlob   =   "goalsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "goalsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SubmitButton_Click()
    Dim ws As Worksheet
    Dim incomeSheet As Worksheet
    Dim expenseSheet As Worksheet
    Dim monthlyIncome As Double
    Dim totalExpenses As Double
    Dim netMonthlyIncome As Double
   
    ' Set the worksheet where data will be stored
    Set ws = ThisWorkbook.Sheets("Goals")
    ' Set the worksheet where monthly income is stored
    Set incomeSheet = ThisWorkbook.Sheets("Incomes")
    ' Set the worksheet where total expenses are stored
    Set expenseSheet = ThisWorkbook.Sheets("Expenses")
   
    ' Get the monthly income from the specified cell
    monthlyIncome = incomeSheet.Range("F2").Value
    ' Get the total expenses from the specified cell
    totalExpenses = expenseSheet.Range("F2").Value
    ' Calculate net monthly income
    netMonthlyIncome = monthlyIncome - totalExpenses

    ' Find the next empty row in column A
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1

    ' Validate date input
    If IsDate(Me.YearTextBox.Value & "-" & Me.MonthTextBox.Value & "-" & Me.DateTextBox.Value) Then
        ' Write date into cell
        ws.Cells(nextRow, "A").Value = DateSerial(Me.YearTextBox.Value, Me.MonthTextBox.Value, Me.DateTextBox.Value)
        ' Format cell so Excel recognizes a date
        ws.Cells(nextRow, "A").NumberFormat = "yyyy-mm-dd;@"
    Else
        MsgBox "Please enter a valid date.", vbExclamation
        Exit Sub
    End If

    ' Transfer the time to complete from the UserForm to column B
    If IsNumeric(Me.TimeToCompleteTextBox.Value) Then
        ws.Cells(nextRow, 2).Value = Me.TimeToCompleteTextBox.Value
    Else
        MsgBox "Please enter a valid number for time to complete.", vbExclamation
        Exit Sub
    End If

    ' Transfer the unit of time from the UserForm to column C
    ws.Cells(nextRow, 3).Value = Me.dropDownBox.Value
    ' Transfer the money to save from the UserForm to column D
    If IsNumeric(Me.MoneyToSaveTextBox.Value) Then
        ws.Cells(nextRow, 4).Value = Me.MoneyToSaveTextBox.Value
    Else
        MsgBox "Please enter a valid number for money to save.", vbExclamation
        Exit Sub
    End If
    
    ' Transfer the description from the UserForm to column E
    ws.Cells(nextRow, 5).Value = Me.dropDownBox1.Value
    ' Transfer the net monthly income to column F
    ws.Cells(nextRow, 6).Value = netMonthlyIncome
   
    ' Calculate the required savings per week based on net monthly income and store the result in column G
    Dim timeToComplete As Double
    Dim result As String
    Dim weeksToSave As Double
    timeToComplete = CDbl(Me.TimeToCompleteTextBox.Value) ' Convert Time to Complete to double
   
    Select Case Me.dropDownBox.Value
        Case "weeks"
            weeksToSave = Me.MoneyToSaveTextBox.Value / (netMonthlyIncome / 4) ' Savings per week
            result = "It will take " & weeksToSave & " weeks to reach this goal"
        Case "months"
            weeksToSave = Me.MoneyToSaveTextBox.Value / netMonthlyIncome
            result = "It will take " & weeksToSave & " months to reach this goal"
        Case "years"
            weeksToSave = Me.MoneyToSaveTextBox.Value / (netMonthlyIncome * 12) ' Savings per month over years
            result = "It will take " & weeksToSave & " years to reach this goal"
        Case Else
            result = "Invalid Unit"
    End Select

    ' Print the result in column G
    ws.Cells(nextRow, 7).Value = result

    ' Clear the form inputs after submission
    Me.DateTextBox.Value = ""
    Me.YearTextBox.Value = ""
    Me.MonthTextBox.Value = ""
    Me.TimeToCompleteTextBox.Value = ""
    ' Clear the Unit of Time dropdown
    Me.dropDownBox.Value = ""
    Me.MoneyToSaveTextBox.Value = ""
    ' Clear the Description dropdown
    Me.dropDownBox1.Value = ""
End Sub

Private Sub UserForm_Initialize()
    ' Populate the ComboBox when the UserForm initializes
    With dropDownBox1
        .AddItem "Emergency Fund"
        .AddItem "Retirement Fund"
        .AddItem "Investment Fund"
        .AddItem "Vacation Fund"
    End With
    
    ' Populate the unit of time ComboBox
    With dropDownBox
        .AddItem "weeks"
        .AddItem "months"
        .AddItem "years"
    End With
End Sub


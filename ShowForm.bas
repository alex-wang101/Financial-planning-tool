Attribute VB_Name = "Module2"
Sub OpenOutput_Click()
    'Open the output userForm
    UserForm1.Show
End Sub
Sub Button3_Click()
    'Oped the 'Add Item' userForm
    AddItemFormIncome.Show
End Sub
Sub Button1_Click()
    'Open expenses 'Add Item' userform
    AddItemFormExpenses.Show
End Sub


Sub IncomeOptions_click()
    'Open the form output
    IncomeOptions.Show
End Sub

Sub ExpensesOptions_click()
    'Open the form output
    ExpensesOptions.Show
End Sub

Sub CompareOptions_click()
    'Open the form output
    CompareOptions.Show
End Sub

Sub openFinancialAdvice_click()
    'Open the form output
    FinancialAdvice.Show
End Sub

Sub GoToGoals()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets("Goals")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "The 'Goals' sheet does not exist.", vbExclamation
    Else
        ws.Activate
    End If
End Sub
Sub openGoalForm_Click()
    'Open the form output
     GoalsUserForm.Show
End Sub



VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddItemFormExpenses 
   Caption         =   "Add Expense/Income"
   ClientHeight    =   5664
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10008
   OleObjectBlob   =   "AddItemFormExpenses.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddItemFormExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    'Set up drop-down menu for form
    With Me.cboxCategory
        .AddItem "Schooling"
        .AddItem "Transportation"
        .AddItem "Groceries"
        .AddItem "Eating Out"
        .AddItem "Healthcare"
        .AddItem "Other"
    End With
End Sub

Private Sub SubmitBtn_Click()
    'Set workbook and sheet
    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Expenses")

    'Start on second row (headers are first row)
    intRow = 2

    'Test value of Item textbox
    If (txtItem.Value <> "") Then
        'Test value of date textboxes
        If (txtDay.Value <> "" And txtMonth.Value <> "" And txtYear.Value <> "") Then
            'Test value of Category combobox
            If (cboxCategory.Value <> "") Then
                'Test value of textValue textbox
                If (textValue.Value <> "" And IsNumeric(textValue.Value)) Then
                    'Go through rows, if they contain data, increment
                    Do While (ws.Cells(intRow, "A") <> "")
                        'Increment row counter
                        intRow = intRow + 1
                    Loop
                    'Write date into cell
                    ws.Cells(intRow, "A") = txtYear.Value + "-" + txtMonth.Value + "-" + txtDay.Value
                    'Format cell so Excel recognizes a date
                    ws.Cells(intRow, "A").NumberFormat = "yyyy-mm-dd;@"
                    'Write value into cell
                    ws.Cells(intRow, "B") = textValue.Value
                    'Write item into cell
                    ws.Cells(intRow, "C") = txtItem.Value
                    'Write category into cell
                    ws.Cells(intRow, "D") = cboxCategory.Value
                    'Write description into cell
                    ws.Cells(intRow, "E") = txtDescription.Value
                    
                    'Sort data by date
                    ws.Range("A2:E" & intRow).Sort Key1:=ws.Range("A2"), Order1:=xlAscending, Header:=xlNo
                    
                    'Clear the userform fields
                    txtItem.Value = ""
                    txtDay.Value = ""
                    txtMonth.Value = ""
                    txtYear.Value = ""
                    textValue.Value = ""
                    cboxCategory.Value = ""
                    txtDescription.Value = ""
                Else
                    'Give error message for no value or invalid value
                    MsgBox ("Please enter a valid value")
                End If
            Else
                'Give error for no category
                MsgBox ("Please select a category")
            End If
        Else
            'Give error message for no date
            MsgBox ("Please enter a valid date")
        End If
    Else
        'Give error message for no item
        MsgBox ("Please enter an item")
    End If
End Sub


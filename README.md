# Financial planning tool
## Overview of Project

This project was created to portray my interest in managing personal finances through data management. This worksheet allows the user to add monthly income/expenses, ask for financial advice, as well as automatically sorting the income/expenses entered by the user. The tool is developed through the use of VBA (Visual Basic Applications), which includes advanced functionalities like custom UserForms and modularized VBA code. 

## Repository Organizsation
This repository is broken into three branches:
## Main Branch
The **main** branch contains the core Excel file (financial_tool.xlsm) that serves as the financial planning tool, along with the aesthetics of the entire project. This file includes macros and VBA code for the tool's functionality.

- **`.frm` Files**: These files define the structure and layout of the UserForms, including controls such as buttons, text boxes, and labels.
- **`.frx` Files**: These files store binary data associated with the UserForms, including custom images, fonts, or formatting. 

#### How to Import UserForms:
1. Open Excel and press `Alt + F11` to open the VBA Editor.
2. Right-click on the project in the **Project Explorer**.
3. Select **Import File**.
4. Choose the `.frm` file from the `userforms/` folder and import it into the project.
5. The corresponding `.frx` file should be placed in the same directory to ensure all resources (e.g., images) are correctly linked.

### Modules Branch
The **modules** branch contains all the exported VBA modules used in the tool. These modules contain the VBA code responsible for core functionality such as stock analysis, budgeting calculations, and other automation.

- **`.bas` Files**: These files contain the VBA code for specific modules used in the tool. You can import these files into your VBA project to reuse the code.

#### How to Import VBA Modules:
1. Open Excel and press `Alt + F11` to open the VBA Editor.
2. Right-click on the project in the **Project Explorer**.
3. Select **Import File**.
4. Choose the `.bas` file from the `modules/` folder and import it into the project.

---

## How to Use the Financial Planning Tool

1. Download the `financial_tool.xlsm` file from the **main** branch.
2. Enable macros in Excel to run the VBA code.
3. Use the UserForms to input data and manage your financial portfolio.
4. Use the stock analysis and budgeting features to track investments and plan your finances.




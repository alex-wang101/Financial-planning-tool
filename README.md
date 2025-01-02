![Financial_planning_tool](https://github.com/user-attachments/assets/272f7c1e-6e3e-41f0-a192-d9dca9b4d065)


This repository was created over the span of 3 months for the purposes of organizing my personal finances. I wanted to create something that not only creates contains the source code and files for a financial planning tool built in Microsoft Excel, utilizing VBA (Visual Basic for Applications) to deliver a dynamic and interactive user experience.

## Project Structure

The repository is organized into three main types of files:

### 1. UserForms (`.frm` and `.frx`)
UserForms are custom graphical interfaces designed to make the application interactive and user-friendly. These forms allow users to input data, make selections, and view results.

- **`.frm` Files**: These are the plain text files containing the code and structure of the UserForms. They include the definitions for the controls (buttons, text boxes, dropdowns, etc.) and their associated event handlers.
- **`.frx` Files**: These are binary files that store the visual design aspects of the UserForms, such as layout, control positioning, and graphical elements. They work alongside the `.frm` files to render the UserForms correctly in the application.

UserForms in this repository:
- `AddItemFormExpenses.frm/.frx`: For adding expense data.
- `AddItemFormIncome.frm/.frx`: For adding income data.
- `FinancialAdvice.frm/.frx`: For displaying financial advice.
- `GoalsUserForm.frm/.frx`: For setting financial goals.
- `OutputForm.frm/.frx`: For generating and displaying outputs.
- `goalsForm.frm/.frx`: Another form for financial goal tracking.
- `UserForm1.frm/.frx`: A generic UserForm for additional features.

### 2. VBA Modules (`.bas`)
VBA modules are files containing the backend logic written in Visual Basic for Applications. These modules handle the functionality of the tool, such as calculations, data processing, and dynamic updates.

Modules in this repository:
- `ChartAndRefresh.bas`: Handles chart creation and refresh operations.
- `ClearModule.bas`: Provides functionality to reset or clear data.
- `CreateGraph.bas`: Includes logic for generating graphs based on financial data.
- `HighlightAndGraph.bas`: Implements features to highlight important data and create visual graphs.
- `ShowForm.bas`: Manages the display and navigation of UserForms.

### 3. Excel Workbook
- **`Financial planning tool.xlsm`**: The main Excel file containing the tool. This file integrates all the VBA modules and UserForms, providing the interface for users to interact with the financial planning tool.

## Features of the Tool
1. **Expense Tracking**: Input, categorize, and analyze expenses using intuitive forms.
2. **Income Management**: Track multiple sources of income efficiently.
3. **Goal Setting**: Set financial goals and visualize progress.
4. **Data Visualization**: Generate dynamic charts and graphs for better decision-making.
5. **Financial Advice**: Receive suggestions based on input data to improve financial health.

## Getting Started
1. Download or clone this repository to your local machine.
2. Open the `Financial planning tool.xlsm` file in Microsoft Excel.
3. Enable macros to allow the tool to function properly.
4. Start by navigating through the UserForms to input your data and explore features.

## Notes
- Ensure that `.frm` and `.frx` files are kept together to maintain the integrity of the UserForms.
- The `.bas` files can be imported into the VBA editor under "Modules" to customize or review the logic.

Feel free to reach out or create an issue for further questions or contributions.

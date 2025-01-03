![Alex's Financial_Planner](https://github.com/user-attachments/assets/4b67d23b-2125-48bd-9a25-71df600eae7d)




This repository was created over the span of 3 months for the purposes of organizing my personal finances. This project organizes personal finances but also illustrates personal interest and knowledge in excel. The following files contains the source code and files for a financial planning tool built in Microsoft Excel, utilizing VBA (Visual Basic for Applications) to deliver a dynamic and interactive user experience.

## Project Structure

The repository is organized into three main types of files:

### 1. UserForms (`.frm` and `.frx`)
UserForms are custom graphical interfaces designed to make the application interactive and user-friendly. These forms allow users to input data, make selections, and view results.

- **`.frm` Files**: These are the plain text files containing the code and structure of the UserForms. They include the definitions for the controls (buttons, text boxes, dropdowns, etc.) and their associated event handlers.
- **`.frx` Files**: These are binary files that store the visual design aspects of the UserForms, such as layout, control positioning, and graphical elements. They work alongside the `.frm` files to render the UserForms correctly in the application.

Features through UserForms:
+ Adding expenses?
  - Click on `Add Expenses` button to reveal userform: `AddItemFormExpenses.frm/.frx`.
<img width="100" alt="image" src="https://github.com/user-attachments/assets/a9891ebc-7b3b-4106-a733-70ec6da54818" />
<img width="300" alt="image" src="https://github.com/user-attachments/assets/5d4ea8f9-8e68-49b4-9cf6-37e11ab8bba4" />

+ Adding Income? 
  - Click on `Add Income` button to reveal userform: `AddItemFormIncome.frm/.frx`.
<img width="102" alt="image" src="https://github.com/user-attachments/assets/e0e116f2-e2b6-4aec-952a-28e544472f8b" />
<img width="300" alt="image" src="https://github.com/user-attachments/assets/35efc039-5ade-42be-becf-ad29d70cfb12" />
  
+ Financial advice?
  - Click on the button `Financial Advice` to reveal userform: `FinancialAdvice.frm/.frx`.
<img width="130" alt="image" src="https://github.com/user-attachments/assets/c17c81db-708c-485a-8126-f3528156ba3f" />
<img width="449" alt="image" src="https://github.com/user-attachments/assets/000d4424-5fb0-46e5-8f46-81f06e5d3677" />

+ Setting goals?
  - Click on the button `Goals` to reveal userform: `GoalsUserForm.frm/.frx`.
<img width="100" alt="image" src="https://github.com/user-attachments/assets/ab880136-335f-43d3-9a74-afae579842f6" />
<img width="300" alt="image" src="https://github.com/user-attachments/assets/aafe6211-e170-4453-b510-c650ef79a3c8" />

+ Output Income/Expenses data onto dashboard?
  - Click on the button `Output info' to reveal userform: `OutputForm.frm/.frx`.
<img width="100" alt="image" src="https://github.com/user-attachments/assets/f59bb620-f33c-4817-8dbc-6a1593a19b09" />
<img width="401" alt="image" src="https://github.com/user-attachments/assets/bb56b7c1-bbe8-4af9-b784-eb5062d6254d" />
  
+ Visualize income distribution?
  - Click on button `Show Income Distribution` to produce a bar chart
<img width="100" alt="image" src="https://github.com/user-attachments/assets/b8c1920d-e504-4d19-a770-a512c1e6c66b" />
<img width="300" alt="image" src="https://github.com/user-attachments/assets/b683e654-5a98-42b2-bcee-374742b355fe" />

+ Visualize expense distribution?
  - CLick on button `Show Expenses Distribution` to produce a bar chart
<img width="100" alt="image" src="https://github.com/user-attachments/assets/21d4e619-0202-4427-9710-ff461ee2a793" />
<img width="300" alt="image" src="https://github.com/user-attachments/assets/e5ada7a9-1efa-4fbe-a842-d15f5a46b290" />

+ Highlight max/min values?
  - Click on Button `Highlight extreme data`
<img width="100" alt="image" src="https://github.com/user-attachments/assets/93286e26-7632-4aca-99c2-5ab54ac3b466" />
<img width="300" alt="image" src="https://github.com/user-attachments/assets/22e837ce-8c26-4564-8325-1792e1e6dd86" />


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

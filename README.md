# QA_Data_Quality_Check_Excel_Automation
Using VBA and Python to do data quality check


## Background

### Project Overview

This repository combines the power of VBA (Visual Basic for Applications) and Python scripts to conduct comprehensive data quality checks. The primary objective is to identify and analyze null cells within each row of a dataset. I was tasked with the responsibility of meticulously inspecting all null cells, tallying their total count, and specifying the corresponding column names. These results are invaluable for future data analysis and decision-making processes.

### Data Quality Checks

The implemented solution utilizes VBA macros to navigate through rows and columns, pinpointing null cells efficiently. Subsequently, Python scripts are employed to process the collected data, enabling the calculation of the total number of null cells and associating them with their respective column names. This seamless integration of VBA and Python ensures a robust and automated data quality checking process.

### Project Scope

This project's scope encompasses the following key tasks:

1. **Identification of Null Cells:** VBA scripts meticulously scan each row, identifying null cells within the dataset.

2. **Counting Null Cells:** The number of null cells in each row is calculated to provide insights into the data's completeness.

3. **Column Identification:** For each null cell detected, the script determines the corresponding column's name, facilitating targeted analysis.

### Future Analysis

The data gathered from this process serves as a foundation for future analysis and decision-making. By understanding the distribution and location of null cells, stakeholders can make informed decisions to enhance data quality, leading to more accurate analyses and valuable insights.

---

### Code

#### VBA Script (Null Cell Count and Identification):

```vba
Sub CountBlankWithCountIfColumn()
    For i = 2 To 9
        Cells(i, 2).Value = WorksheetFunction.CountIf(Range("C" & i & ":AF" & i), "")
    Next i
End Sub

Sub NullCellIdentification()
Dim n As Integer
Dim j As Integer

For j = 2 To 9

For n = 3 To 32
    If Cells(j, n).Value = "" Then
    Cells(j, 1).Value = Cells(j, 1).Value & Cells(1, n) & ";"
    End If
    
Next n
    
Next j

End Sub

```

#### python Script (Null Cell Count and Identification):
```Python
import openpyxl

def count_null_cells_in_row(sheet, row_number, start_column):
    null_cell_count = 0
    first_row_content = []

    # Iterate over columns from the starting column to the end
    for column in range(start_column, sheet.max_column + 1):
        cell = sheet.cell(row=row_number, column=column)
        first_row_cell = sheet.cell(row=1, column=column).value  # Get the corresponding cell in the first row
        if cell.value is None:
            null_cell_count += 1
            first_row_content.append(first_row_cell)

    if null_cell_count > 1:
        return ';'.join(str(item) for item in first_row_content)
    elif null_cell_count == 1:
        return first_row_content[0]  # If there's only one null cell, return its content
    else:
        return None

def main(file_path, sheet_name, start_column):
    try:
        # Load the Excel workbook
        workbook = openpyxl.load_workbook(file_path)

        # Select the sheet by name
        sheet = workbook[sheet_name]

        # Iterate through all rows except the first row
        for row_number in range(2, sheet.max_row + 1):
            first_row_content = count_null_cells_in_row(sheet, row_number, start_column)
            sheet.cell(row=row_number, column=1, value=first_row_content)

        # Save the changes to the Excel file
        workbook.save(file_path)

    except Exception as e:
        print(f"An error occurred: {str(e)}")

# Example usage:
file_path = 'test.xlsx'  # Replace with the path to your Excel file
sheet_name = 'Sheet1'               # Replace with the name of your sheet
start_column = 3                   # Replace with the starting column (C)

main(file_path, sheet_name, start_column)
```
## Lessons Learned

### Scalability in Big Data

One of the significant takeaways from this project was understanding the scalability of the solution in handling vast volumes of data. Working with thousands of rows, I observed the efficiency of the implemented solution, especially in scenarios where traditional VBA scripts might struggle due to performance limitations. This experience highlighted the importance of developing solutions that can scale effectively, ensuring seamless operations even with extensive datasets.

### Hybrid Approach: VBA and Python

Recognizing the potential limitations of VBA when dealing with big data, I adopted a hybrid approach by integrating Python scripts into the solution. This decision proved instrumental in overcoming performance challenges. While VBA provided the initial data extraction capabilities, Python was leveraged to handle large-scale data processing efficiently. This hybrid integration not only enhanced the solution's performance but also showcased the versatility of combining different technologies to achieve optimal results.



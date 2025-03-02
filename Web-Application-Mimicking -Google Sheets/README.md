# Web Application Mimicking Google Sheets

## Objective

Develop a web application that closely mimics the user interface and core functionalities of Google Sheets, focusing on mathematical and data quality functions, data entry, and key UI interactions.

## Features

### 1. Spreadsheet Interface

- **Mimic Google Sheets UI**: Visual design and layout closely resemble Google Sheets, including the toolbar, formula bar, and cell structure.
- **Drag Functions**: Implement drag functionality for cell content, formulas, and selections to mirror Google Sheets' behavior.
- **Cell Dependencies**: Ensure formulas and functions accurately reflect cell dependencies and update accordingly when changes are made to related cells.
- **Basic Cell Formatting**: Support for bold, italics, font size, and color.
- **Row and Column Management**: Ability to add, delete, and resize rows and columns.

### 2. Mathematical Functions

- **SUM**: Calculates the sum of a range of cells.
- **AVERAGE**: Calculates the average of a range of cells.
- **MAX**: Returns the maximum value from a range of cells.
- **MIN**: Returns the minimum value from a range of cells.
- **COUNT**: Counts the number of cells containing numerical values in a range.

### 3. Data Quality Functions

- **TRIM**: Removes leading and trailing whitespace from a cell.
- **UPPER**: Converts the text in a cell to uppercase.
- **LOWER**: Converts the text in a cell to lowercase.
- **REMOVE\_DUPLICATES**: Removes duplicate rows from a selected range.
- **FIND\_AND\_REPLACE**: Allows users to find and replace specific text within a range of cells.

### 4. Data Entry and Validation

- Allow users to input various data types (numbers, text, dates).
- Implement basic data validation checks (e.g., ensuring numeric cells only contain numbers).

### 5. Testing

- Provide a means for users to test the implemented functions with their own data.
- Display the results of function execution clearly.

## Bonus Features

- Implement additional mathematical and data quality functions.
- Support more complex formulas and cell referencing (e.g., relative and absolute references).
- Allow users to save and load their spreadsheets.
- Incorporate data visualization capabilities (e.g., charts, graphs).
- Implement keyboard shortcuts (e.g., `Ctrl + S` for saving).

## Project Folder Structure

```
Web-Application-Mimicking-Google-Sheets/
│── index.html          # Main HTML file
│── styles.css          # CSS file for styling
│── script.js           # Main JavaScript file
│── functionalities.js  # Handles core spreadsheet functionalities
│── sheet.js            # Manages sheet-specific operations
│── README.md           # Project documentation
```

## Installation

1. Clone the repository:
   ```sh
   git clone https://github.com/your-username/Web-Application-Mimicking-Google-Sheets.git
   ```
2. Open `index.html` in a web browser.

## Technologies Used

- **HTML5** for structure
- **CSS3** for styling
- **JavaScript (ES6+)** for functionalities

## Future Enhancements

- Introduce more advanced spreadsheet functions.
- Improve performance with larger datasets.
- Implement real-time collaboration features.

## Contributors

- [Nagesh N]

##

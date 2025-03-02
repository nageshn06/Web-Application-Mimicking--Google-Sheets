const initialCellState = {
    fontFamily_data: 'Times New Roman',
    fontSize_data: '12',
    isBold: false,
    isItalic: false,
    textAlign: 'start',
    isUnderlined: false,
    color: '#000000',
    backgroundColor: '#ffffff',
    content: ''
};

let sheetsArray = [];
let activeSheetIndex = -1;
let activeSheetObject = false;
let activeCell = false;

// Functionality elements
let fontFamilyBtn = document.querySelector('.font-family');
let fontSizeBtn = document.querySelector('.font-size');
let boldBtn = document.querySelector('.bold');
let italicBtn = document.querySelector('.italic');
let underlineBtn = document.querySelector('.underline');
let leftBtn = document.querySelector('.start');
let centerBtn = document.querySelector('.center');
let rightBtn = document.querySelector('.end');
let colorBtn = document.querySelector('#color');
let bgColorBtn = document.querySelector('#bgcolor');
let addressBar = document.querySelector('.address-bar');
let formula = document.querySelector('.formula-bar');
let downloadBtn = document.querySelector(".download");
let openBtn = document.querySelector(".open");
let addColBtn = document.querySelector("#addColBtn");
let deleteColBtn = document.querySelector("#deleteColBtn");
let addRowBtn = document.querySelector("#addRowBtn");
let deleteRowBtn = document.querySelector("#deleteRowBtn");

// Grid header row element
let gridHeader = document.querySelector('.grid-header');

// Add header column
let bold = document.createElement('div');
bold.className = 'grid-header-col';
bold.innerText = 'SL. NO.';
gridHeader.append(bold);
for (let i = 65; i <= 90; i++) {
    let bold = document.createElement('div');
    bold.className = 'grid-header-col';
    bold.innerText = String.fromCharCode(i);
    bold.id = String.fromCharCode(i);
    gridHeader.append(bold);
}

for (let i = 1; i <= 100; i++) {
    let newRow = document.createElement('div')
    newRow.className = 'row';
    document.querySelector('.grid').append(newRow);

    let bold = document.createElement('div');
    bold.className = 'grid-cell';
    bold.innerText = i;
    bold.id = i;
    newRow.append(bold);

    for (let j = 65; j <= 90; j++) {
        let cell = document.createElement('div');
        cell.className = 'grid-cell cell-focus';
        cell.id = String.fromCharCode(j) + i;
        cell.contentEditable = true;

        cell.addEventListener('click', (event) => {
            event.stopPropagation();
        });
        cell.addEventListener('focus', cellFocus);
        cell.addEventListener('focusout', cellFocusOut);
        cell.addEventListener('input', cellInput);

        newRow.append(cell);
    }
}

function cellFocus(event) {
    let key = event.target.id;
    addressBar.innerHTML = event.target.id;
    activeCell = event.target;

    let activeBg = '#c9c8c8';
    let inactiveBg = '#ecf0f1';

    fontFamilyBtn.value = activeSheetObject[key].fontFamily_data;
    fontSizeBtn.value = activeSheetObject[key].fontSize_data;
    boldBtn.style.backgroundColor = activeSheetObject[key].isBold ? activeBg : inactiveBg;
    italicBtn.style.backgroundColor = activeSheetObject[key].isItalic ? activeBg : inactiveBg;
    underlineBtn.style.backgroundColor = activeSheetObject[key].isUnderlined ? activeBg : inactiveBg;
    setAlignmentBg(key, activeBg, inactiveBg);
    colorBtn.value = activeSheetObject[key].color;
    bgColorBtn.value = activeSheetObject[key].backgroundColor;

    formula.value = activeCell.innerText;

    document.getElementById(event.target.id.slice(0, 1)).classList.add('row-col-focus');
    document.getElementById(event.target.id.slice(1)).classList.add('row-col-focus');
}

function cellInput(event) {
    let key = activeCell.id;
    validateInput(activeCell); // Validate input before saving
    formula.value = activeCell.innerText;
    activeSheetObject[key].content = activeCell.innerText;

    // Update dependencies
    updateDependencies(key);

    // Move the cursor to the end of the cell's content
    moveCursorToEnd(activeCell);
}

// Function to move the cursor to the end of the content
function moveCursorToEnd(cell) {
    const range = document.createRange();
    const selection = window.getSelection();
    range.selectNodeContents(cell);
    range.collapse(false); // Move the cursor to the end
    selection.removeAllRanges();
    selection.addRange(range);
}

function setAlignmentBg(key, activeBg, inactiveBg) {
    leftBtn.style.backgroundColor = inactiveBg;
    centerBtn.style.backgroundColor = inactiveBg;
    rightBtn.style.backgroundColor = inactiveBg;
    if (key) {
        document.querySelector('.' + activeSheetObject[key].textAlign).style.backgroundColor = activeBg;
    } else {
        leftBtn.style.backgroundColor = activeBg;
    }
}

function cellFocusOut(event) {
    document.getElementById(event.target.id.slice(0, 1)).classList.remove('row-col-focus');
    document.getElementById(event.target.id.slice(1)).classList.remove('row-col-focus');
}

// Function to validate input
function validateInput(cell) {
    const value = cell.innerText.trim();

    // Validate numeric input
    if (!isNaN(value)) {
        cell.innerText = parseFloat(value).toString(); // Ensure numeric format
        return;
    }

    // Validate date input
    const date = Date.parse(value);
    if (!isNaN(date)) {
        cell.innerText = new Date(date).toLocaleDateString(); // Convert to localized date format
        return;
    }

    // If it's not a number or date, leave it as text
}

// Drag functionality
let isDragging = false;
let startCell = null;
let endCell = null;

document.querySelectorAll('.grid-cell').forEach(cell => {
    cell.addEventListener('mousedown', (event) => {
        isDragging = true;
        startCell = event.target;
    });

    cell.addEventListener('mousemove', (event) => {
        if (isDragging) {
            endCell = event.target;
        }
    });

    cell.addEventListener('mouseup', (event) => {
        if (isDragging && startCell && endCell) {
            handleDrag(startCell, endCell);
            isDragging = false;
            startCell = null;
            endCell = null;
        }
    });
});

function handleDrag(startCell, endCell) {
    const startId = startCell.id;
    const endId = endCell.id;

    const startCol = startId[0];
    const startRow = parseInt(startId.slice(1));
    const endCol = endId[0];
    const endRow = parseInt(endId.slice(1));

    const colStart = Math.min(startCol.charCodeAt(0), endCol.charCodeAt(0));
    const colEnd = Math.max(startCol.charCodeAt(0), endCol.charCodeAt(0));
    const rowStart = Math.min(startRow, endRow);
    const rowEnd = Math.max(startRow, endRow);

    for (let col = colStart; col <= colEnd; col++) {
        for (let row = rowStart; row <= rowEnd; row++) {
            const cellId = String.fromCharCode(col) + row;
            const cell = document.getElementById(cellId);
            if (cell) {
                cell.innerText = startCell.innerText;
                activeSheetObject[cellId].content = startCell.innerText;
            }
        }
    }
}

// Dependency graph
const dependencyGraph = {};

function updateDependencies(cellId) {
    if (dependencyGraph[cellId]) {
        dependencyGraph[cellId].forEach(dependentCellId => {
            const cell = document.getElementById(dependentCellId);
            if (cell) {
                evaluateFormula(cell);
            }
        });
    }
}

// Formula evaluation
function evaluateFormula(cell) {
    const formulaText = cell.innerText;
    if (formulaText.startsWith('=')) {
        const formula = formulaText.slice(1).trim();

        // Handle IF statements
        const ifMatch = formula.match(/IF\((.*),(.*),(.*)\)/i);
        if (ifMatch) {
            const [_, condition, trueValue, falseValue] = ifMatch;
            try {
                const conditionResult = eval(condition); // Evaluate the condition
                cell.innerText = conditionResult ? trueValue : falseValue;
            } catch (e) {
                cell.innerText = '#ERROR';
            }
            return;
        }

        // Handle VLOOKUP
        const vlookupMatch = formula.match(/VLOOKUP\("(.*)","(.*)",(\d+)\)/i);
        if (vlookupMatch) {
            const [_, lookupValue, range, colIndex] = vlookupMatch;
            try {
                cell.innerText = handleVLookup(lookupValue, range, parseInt(colIndex));
            } catch (e) {
                cell.innerText = '#ERROR';
            }
            return;
        }

        // Handle CONCATENATE
        const concatMatch = formula.match(/CONCATENATE\((.*)\)/i);
        if (concatMatch) {
            const args = concatMatch[1].split(',').map(arg => arg.trim().replace(/"/g, ''));
            try {
                cell.innerText = handleConcatenate(...args);
            } catch (e) {
                cell.innerText = '#ERROR';
            }
            return;
        }

        // Handle basic math functions (SUM, AVERAGE, MAX, MIN, COUNT)
        const mathMatch = formula.match(/(SUM|AVERAGE|MAX|MIN|COUNT)\(([A-Z]+\d+):([A-Z]+\d+)\)/i);
        if (mathMatch) {
            const [_, func, startCellId, endCellId] = mathMatch;
            try {
                const values = getRangeValues(startCellId, endCellId);
                switch (func.toUpperCase()) {
                    case 'SUM':
                        cell.innerText = values.reduce((a, b) => a + b, 0);
                        break;
                    case 'AVERAGE':
                        cell.innerText = values.reduce((a, b) => a + b, 0) / values.length;
                        break;
                    case 'MAX':
                        cell.innerText = Math.max(...values);
                        break;
                    case 'MIN':
                        cell.innerText = Math.min(...values);
                        break;
                    case 'COUNT':
                        cell.innerText = values.length;
                        break;
                }
            } catch (e) {
                cell.innerText = '#ERROR';
            }
            return;
        }

        // Handle simple arithmetic expressions
        try {
            const result = eval(formula); // Use eval for simple arithmetic (caution: security risk in production)
            cell.innerText = result;
        } catch (e) {
            cell.innerText = '#ERROR';
        }
    }
}

// Function to resolve cell references
function resolveReference(cellId, reference) {
    const isAbsoluteCol = reference.startsWith('$');
    const isAbsoluteRow = reference.match(/\$\d+/);

    let col = reference.replace(/[^A-Z]/g, '');
    let row = reference.replace(/\D/g, '');

    if (!isAbsoluteCol) {
        const currentCol = cellId[0];
        const currentColCode = currentCol.charCodeAt(0);
        col = String.fromCharCode(currentColCode + (col.charCodeAt(0) - 65));
    }

    if (!isAbsoluteRow) {
        const currentRow = parseInt(cellId.slice(1));
        row = currentRow + parseInt(row) - 1;
    }

    return col + row;
}

// Function to get values from a range
function getRangeValues(startCellId, endCellId) {
    const values = [];
    const startCol = startCellId[0];
    const startRow = parseInt(startCellId.slice(1));
    const endCol = endCellId[0];
    const endRow = parseInt(endCellId.slice(1));

    for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
        for (let row = startRow; row <= endRow; row++) {
            const cellId = String.fromCharCode(col) + row;
            const cell = document.getElementById(cellId);
            if (cell && !isNaN(cell.innerText)) {
                values.push(parseFloat(cell.innerText));
            }
        }
    }
    return values;
}

// Function to handle IF statements
function handleIfStatement(condition, trueValue, falseValue) {
    return condition ? trueValue : falseValue;
}

// Function to handle VLOOKUP
function handleVLookup(lookupValue, range, colIndex) {
    const [startCellId, endCellId] = range.split(':');
    const startCol = startCellId[0];
    const startRow = parseInt(startCellId.slice(1));
    const endCol = endCellId[0];
    const endRow = parseInt(endCellId.slice(1));

    for (let row = startRow; row <= endRow; row++) {
        const lookupCellId = startCol + row;
        const lookupCell = document.getElementById(lookupCellId);
        if (lookupCell && lookupCell.innerText === lookupValue) {
            const resultCellId = String.fromCharCode(startCol.charCodeAt(0) + colIndex - 1) + row;
            const resultCell = document.getElementById(resultCellId);
            return resultCell ? resultCell.innerText : '#N/A';
        }
    }
    return '#N/A';
}

// Function to handle CONCATENATE
function handleConcatenate(...args) {
    return args.join('');
}

// Function to add a column
function addColumn() {
    const colCount = gridHeader.children.length;
    const newColName = getColumnName(colCount);

    // Add column header
    const colHeader = document.createElement('div');
    colHeader.className = 'grid-header-col';
    colHeader.innerText = newColName;
    gridHeader.append(colHeader);

    // Add cells for the new column
    const rows = document.querySelectorAll('.row');
    rows.forEach((row, index) => {
        const cell = document.createElement('div');
        cell.className = 'grid-cell cell-focus';
        cell.id = newColName + (index + 1);
        cell.contentEditable = true;

        cell.addEventListener('click', (event) => event.stopPropagation());
        cell.addEventListener('focus', cellFocus);
        cell.addEventListener('focusout', cellFocusOut);
        cell.addEventListener('input', cellInput);

        row.append(cell);
        activeSheetObject[cell.id] = { ...initialCellState };
    });
}

// Function to delete a column
function deleteColumn() {
    const colCount = gridHeader.children.length;
    if (colCount > 1) {
        const colName = gridHeader.lastChild.innerText; // Get the name of the last column
        gridHeader.lastChild.remove(); // Remove the column header

        // Remove cells for the deleted column from all rows
        const rows = document.querySelectorAll('.row');
        rows.forEach((row) => {
            const cell = row.querySelector(`.grid-cell[id^="${colName}"]`); // Find the cell in this row with the matching column name
            if (cell) {
                cell.remove(); // Remove the cell from the DOM
                delete activeSheetObject[cell.id]; // Remove the cell data from activeSheetObject
            }
        });

        // Remove rows beyond the initial 100 rows if they exist
        const extraRows = document.querySelectorAll('.row');
        extraRows.forEach((row) => {
            const rowNumber = row.querySelector('.grid-cell').innerText; // Get the row number
            const cellId = colName + rowNumber; // Construct the cell ID (e.g., "AA101")
            const cell = document.getElementById(cellId);
            if (cell) {
                cell.remove(); // Remove the cell from the DOM
                delete activeSheetObject[cell.id]; // Remove the cell data from activeSheetObject
            }
        });
    }
}

// Function to add a row
function addRow() {
    const rowCount = document.querySelectorAll('.row').length;
    const newRow = document.createElement('div');
    newRow.className = 'row';
    document.querySelector('.grid').append(newRow);

    // Add row number cell
    const rowNumberCell = document.createElement('div');
    rowNumberCell.className = 'grid-cell';
    rowNumberCell.innerText = rowCount + 1;
    rowNumberCell.id = rowCount + 1;
    newRow.append(rowNumberCell);

    // Add cells for the new row
    for (let j = 65; j < 65 + gridHeader.children.length - 1; j++) {
        const cell = document.createElement('div');
        cell.className = 'grid-cell cell-focus';
        cell.id = String.fromCharCode(j) + (rowCount + 1);
        cell.contentEditable = true;

        cell.addEventListener('click', (event) => event.stopPropagation());
        cell.addEventListener('focus', cellFocus);
        cell.addEventListener('focusout', cellFocusOut);
        cell.addEventListener('input', cellInput);

        newRow.append(cell);
        activeSheetObject[cell.id] = { ...initialCellState };
    }
}

// Function to delete a row
function deleteRow() {
    const rows = document.querySelectorAll('.row');
    if (rows.length > 1) {
        const lastRow = rows[rows.length - 1];
        lastRow.remove();

        // Remove row data from activeSheetObject
        for (let key in activeSheetObject) {
            if (key.endsWith(rows.length)) {
                delete activeSheetObject[key];
            }
        }
    }
}

// Function to generate column names beyond Z
function getColumnName(index) {
    let name = '';
    while (index > 0) {
        index--;
        name = String.fromCharCode(65 + (index % 26)) + name;
        index = Math.floor(index / 26);
    }
    return name;
}

// Add event listeners for add/delete column and row buttons
addColBtn.addEventListener('click', addColumn);
deleteColBtn.addEventListener('click', deleteColumn);
addRowBtn.addEventListener('click', addRow);
deleteRowBtn.addEventListener('click', deleteRow);

// Chart functionality
document.querySelector('.chart-btn').addEventListener('click', generateChart);

function generateChart() {
    const selectedCells = Array.from(document.querySelectorAll('.grid-cell:focus'));
    const labels = selectedCells.map((cell, index) => `Data ${index + 1}`);
    const data = selectedCells.map(cell => parseFloat(cell.innerText) || 0);

    const ctx = document.createElement('canvas');
    ctx.id = 'chart';
    document.body.append(ctx);

    new Chart(ctx, {
        type: 'bar', // You can change this to 'line', 'pie', etc.
        data: {
            labels: labels,
            datasets: [{
                label: 'Cell Data',
                data: data,
                backgroundColor: 'rgba(75, 192, 192, 0.2)',
                borderColor: 'rgba(75, 192, 192, 1)',
                borderWidth: 1
            }]
        },
        options: {
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}

// Keyboard shortcuts
document.addEventListener('keydown', (event) => {
    if (event.ctrlKey || event.metaKey) {
        switch (event.key) {
            case 'c':
                document.querySelector('.copy').click();
                break;
            case 'v':
                document.querySelector('.paste').click();
                break;
            case 'z':
                // Implement undo functionality
                break;
            case 'y':
                // Implement redo functionality
                break;
        }
    }
});
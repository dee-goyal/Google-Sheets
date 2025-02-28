// ===== GLOBAL CONFIG =====
const DEFAULT_ROWS = 20;
const DEFAULT_COLS = 10; // Expand for more columns
let spreadsheetData = []; // 2D array storing { value, formula }
let isUpdating = false;   // Prevents recursive calls

document.addEventListener("DOMContentLoaded", () => {
  // 1. Generate the table
  generateTable(DEFAULT_ROWS, DEFAULT_COLS);

  // 2. Initialize data model
  initSpreadsheetData(DEFAULT_ROWS, DEFAULT_COLS);

  // 3. Attach additional events if needed
  attachEventListeners();
});

/* -------------------------------------------
 *            TABLE GENERATION
 * -----------------------------------------*/
function generateTable(rows, cols) {
  const thead = document.querySelector("#spreadsheet thead tr");
  const tbody = document.querySelector("#spreadsheet tbody");

  // Clear existing content
  thead.innerHTML = "<th></th>"; // top-left corner
  tbody.innerHTML = "";

  // Create column headers (A, B, C...)
  for (let c = 0; c < cols; c++) {
    const th = document.createElement("th");
    th.textContent = columnIndexToName(c); // e.g., 0->A, 1->B
    thead.appendChild(th);
  }

  // Create rows with row headers
  for (let r = 0; r < rows; r++) {
    const tr = document.createElement("tr");

    // Row header
    const rowHeader = document.createElement("th");
    rowHeader.textContent = r + 1; // 1-based
    tr.appendChild(rowHeader);

    // Create cells
    for (let c = 0; c < cols; c++) {
      const td = document.createElement("td");
      td.contentEditable = "true";
      td.addEventListener("input", () => onCellInput(r, c, td));
      td.addEventListener("focus", () => onCellFocus(r, c, td));
      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }
}

/* -------------------------------------------
 *           DATA MODEL & INIT
 * -----------------------------------------*/
function initSpreadsheetData(rows, cols) {
  spreadsheetData = [];
  for (let r = 0; r < rows; r++) {
    const rowData = [];
    for (let c = 0; c < cols; c++) {
      rowData.push({ value: "", formula: "" });
    }
    spreadsheetData.push(rowData);
  }
}

/* -------------------------------------------
 *        CELL INPUT & FORMULAS
 * -----------------------------------------*/
function onCellInput(row, col, cellElem) {
  const raw = cellElem.innerText.trim();
  if (raw.startsWith("=")) {
    // It's a formula
    spreadsheetData[row][col].formula = raw;
    spreadsheetData[row][col].value = "";
  } else {
    // Just text or number
    spreadsheetData[row][col].formula = "";
    spreadsheetData[row][col].value = raw;
  }
  recalcAll();
}

function onCellFocus(row, col, cellElem) {
  // Mark cell as active
  document.querySelectorAll("td").forEach(td => td.classList.remove("active-cell"));
  cellElem.classList.add("active-cell");

  // Show formula/value in formula bar
  const formulaBar = document.getElementById("formula-bar");
  const cellData = spreadsheetData[row][col];
  formulaBar.value = cellData.formula ? cellData.formula : cellData.value;
}

// Recalculate entire sheet
function recalcAll() {
  if (isUpdating) return;
  isUpdating = true;

  for (let r = 0; r < spreadsheetData.length; r++) {
    for (let c = 0; c < spreadsheetData[r].length; c++) {
      const cell = spreadsheetData[r][c];
      if (cell.formula.startsWith("=")) {
        const val = evaluateFormula(cell.formula.substring(1)); // remove '='
        cell.value = val;
        const td = getCellElement(r, c);
        if (td) td.innerText = val;
      }
    }
  }

  isUpdating = false;
}

// Evaluate formula like "A1+B2" or "SUM(A1:A5)"
function evaluateFormula(formula) {
  formula = formula.toUpperCase();

  // Built-in functions: SUM, AVERAGE, MAX, MIN, COUNT
  const funcMatch = formula.match(/^([A-Z]+)\(([^)]+)\)$/);
  if (funcMatch) {
    const funcName = funcMatch[1];
    const range = funcMatch[2];
    const values = parseRange(range).map(v => parseFloat(v) || 0);

    switch (funcName) {
      case "SUM":
        return values.reduce((a, b) => a + b, 0);
      case "AVERAGE":
        return values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0;
      case "MAX":
        return values.length ? Math.max(...values) : 0;
      case "MIN":
        return values.length ? Math.min(...values) : 0;
      case "COUNT":
        return values.filter(x => !isNaN(x)).length;
      default:
        return "#FUNC?";
    }
  }

  // If not a function, handle references like A1+B2
  formula = formula.replace(/([A-Z]+)(\d+)/g, (match, colLetters, rowNumber) => {
    const cIndex = colNameToIndex(colLetters);
    const rIndex = parseInt(rowNumber) - 1;
    if (spreadsheetData[rIndex] && spreadsheetData[rIndex][cIndex]) {
      const val = parseFloat(spreadsheetData[rIndex][cIndex].value);
      return isNaN(val) ? 0 : val;
    }
    return 0;
  });

  try {
    // eslint-disable-next-line no-new-func
    return new Function("return " + formula)();
  } catch (err) {
    return "#ERROR!";
  }
}

// Parse range like "A1:A5"
function parseRange(rangeStr) {
  const [start, end] = rangeStr.split(":");
  if (!end) {
    // single cell
    return [getCellValueByRef(start)];
  }

  const startCol = start.match(/[A-Z]+/)[0];
  const startRow = parseInt(start.match(/\d+/)[0]) - 1;
  const endCol = end.match(/[A-Z]+/)[0];
  const endRow = parseInt(end.match(/\d+/)[0]) - 1;
  const startIndex = colNameToIndex(startCol);
  const endIndex = colNameToIndex(endCol);

  const values = [];
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startIndex; c <= endIndex; c++) {
      if (spreadsheetData[r] && spreadsheetData[r][c]) {
        values.push(spreadsheetData[r][c].value);
      }
    }
  }
  return values;
}

// Convert col letters to index
function colNameToIndex(colLetters) {
  let index = 0;
  for (let i = 0; i < colLetters.length; i++) {
    index = index * 26 + (colLetters.charCodeAt(i) - 64);
  }
  return index - 1;
}

// Return the DOM <td> for row r, col c
function getCellElement(r, c) {
  const rowElem = document.querySelector("#spreadsheet tbody").rows[r];
  if (!rowElem) return null;
  return rowElem.cells[c + 1]; // +1 for row header
}

// Return cell value by reference "A1"
function getCellValueByRef(ref) {
  const colLetters = ref.match(/[A-Z]+/)[0];
  const rowNumber = parseInt(ref.match(/\d+/)[0]);
  const cIndex = colNameToIndex(colLetters);
  const rIndex = rowNumber - 1;
  if (spreadsheetData[rIndex] && spreadsheetData[rIndex][cIndex]) {
    return spreadsheetData[rIndex][cIndex].value;
  }
  return "";
}

/* -------------------------------------------
 *        FORMULA BAR (Enter Key)
 * -----------------------------------------*/
document.getElementById("formula-bar").addEventListener("keydown", function (event) {
  if (event.key === "Enter") {
    let formula = this.value.trim();
    if (formula.startsWith("=")) {
      // Convert references to numeric values
      let expression = formula.substring(1).toUpperCase();
      expression = expression.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
        let cIndex = colNameToIndex(col);
        let rIndex = parseInt(row) - 1;
        let val = parseFloat(spreadsheetData[rIndex][cIndex].value) || 0;
        return val;
      });

      try {
        let result = eval(expression);
        let activeCell = document.querySelector(".active-cell");
        if (activeCell) {
          activeCell.innerText = result;
          // Update data model
          const r = activeCell.parentElement.rowIndex - 1; // first row is col headers
          const c = activeCell.cellIndex - 1;              // first col is row headers
          spreadsheetData[r][c].formula = formula;
          spreadsheetData[r][c].value = result;
        } else {
          alert("Click a cell first!");
        }
      } catch (err) {
        alert("Invalid formula!");
      }
    }
  }
});

/* -------------------------------------------
 *       FORMATTING FUNCTIONS
 * -----------------------------------------*/
function applyFormat(format) {
  const active = document.activeElement;
  if (active && active.tagName === "TD") {
    switch (format) {
      case "bold":
        toggleStyle(active, "fontWeight", "bold", "normal");
        break;
      case "italic":
        toggleStyle(active, "fontStyle", "italic", "normal");
        break;
      case "underline":
        toggleStyle(active, "textDecoration", "underline", "none");
        break;
    }
  }
}

function toggleStyle(elem, styleProp, valOn, valOff) {
  elem.style[styleProp] = (elem.style[styleProp] === valOn) ? valOff : valOn;
}

function applyFontSize(size) {
  const active = document.activeElement;
  if (active && active.tagName === "TD") {
    active.style.fontSize = size;
  }
}

function applyColor(color) {
  const active = document.activeElement;
  if (active && active.tagName === "TD") {
    active.style.color = color;
  }
}

/* -------------------------------------------
 *    DATA QUALITY (TRIM, UPPER, LOWER, etc.)
 * -----------------------------------------*/
function performDataQuality(type) {
  const selectedCells = getSelectedCells();
  selectedCells.forEach(cell => {
    let text = cell.innerText;
    switch (type) {
      case "trim":
        cell.innerText = text.trim();
        break;
      case "upper":
        cell.innerText = text.toUpperCase();
        break;
      case "lower":
        cell.innerText = text.toLowerCase();
        break;
    }
    // Update data model
    const [r, c] = getCellRC(cell);
    spreadsheetData[r][c].value = cell.innerText;
    spreadsheetData[r][c].formula = cell.innerText.startsWith("=") ? cell.innerText : "";
  });
  recalcAll();
}

// Remove duplicates (by entire row) in selected range
function removeDuplicates() {
  const rows = getSelectedRowRange();
  const uniqueSet = new Set();
  rows.forEach(rowElem => {
    const rowText = Array.from(rowElem.cells).map(td => td.innerText).join("||");
    if (uniqueSet.has(rowText)) {
      // Duplicate row => clear it
      Array.from(rowElem.cells).forEach(td => td.innerText = "");
    } else {
      uniqueSet.add(rowText);
    }
  });
  recalcAll();
}

// Find & Replace
function findAndReplace() {
  const findText = prompt("Enter text to find:");
  if (findText === null) return;
  const replaceText = prompt("Enter replacement text:");
  if (replaceText === null) return;

  const selectedCells = getSelectedCells();
  selectedCells.forEach(cell => {
    cell.innerText = cell.innerText.replace(new RegExp(findText, "g"), replaceText);
    const [r, c] = getCellRC(cell);
    spreadsheetData[r][c].value = cell.innerText;
    spreadsheetData[r][c].formula = cell.innerText.startsWith("=") ? cell.innerText : "";
  });
  recalcAll();
}

/* -------------------------------------------
 *       MATH FUNCTIONS (SUM, AVERAGE, etc.)
 * -----------------------------------------*/
function performMath(type) {
  const selectedCells = getSelectedCells();
  const numbers = selectedCells.map(cell => parseFloat(cell.innerText) || 0);
  let result = 0;

  switch (type) {
    case "SUM":
      result = numbers.reduce((a, b) => a + b, 0);
      break;
    case "AVERAGE":
      result = numbers.length ? (numbers.reduce((a, b) => a + b, 0) / numbers.length) : 0;
      break;
    case "MAX":
      result = numbers.length ? Math.max(...numbers) : 0;
      break;
    case "MIN":
      result = numbers.length ? Math.min(...numbers) : 0;
      break;
    case "COUNT":
      result = numbers.filter(n => !isNaN(n)).length;
      break;
  }
  document.getElementById("result-display").innerText = `Result: ${result}`;
}

/* -------------------------------------------
 *     RANGE SELECTION (for Data Quality)
 * -----------------------------------------*/
function getSelectedCells() {
  const colInput = document.getElementById("col-select").value.toUpperCase();
  const rowStart = parseInt(document.getElementById("row-start").value, 10);
  const rowEnd = parseInt(document.getElementById("row-end").value, 10);
  if (!colInput || isNaN(rowStart) || isNaN(rowEnd)) return [];

  const columns = colInput.split(",").map(s => s.trim());
  const cells = [];
  const tbody = document.querySelector("#spreadsheet tbody");

  for (let r = rowStart; r <= rowEnd; r++) {
    const rowIndex = r - 1;
    const rowElem = tbody.rows[rowIndex];
    if (!rowElem) continue;

    columns.forEach(colLetter => {
      const cIndex = colNameToIndex(colLetter);
      const cellElem = rowElem.cells[cIndex + 1]; // offset for row header
      if (cellElem) cells.push(cellElem);
    });
  }
  return cells;
}

// Return array of row <tr> in selected range
function getSelectedRowRange() {
  const rowStart = parseInt(document.getElementById("row-start").value, 10);
  const rowEnd = parseInt(document.getElementById("row-end").value, 10);
  const tbody = document.querySelector("#spreadsheet tbody");
  const rows = [];
  for (let r = rowStart; r <= rowEnd; r++) {
    const rowIndex = r - 1;
    if (tbody.rows[rowIndex]) rows.push(tbody.rows[rowIndex]);
  }
  return rows;
}

// Return [rowIndex, colIndex] for a cell
function getCellRC(cellElem) {
  const rowIndex = cellElem.parentElement.rowIndex - 1; // thead row is 0
  const colIndex = cellElem.cellIndex - 1;              // row header is 0
  return [rowIndex, colIndex];
}

/* -------------------------------------------
 *    ROW & COLUMN MANAGEMENT
 * -----------------------------------------*/
function addRow() {
  const tbody = document.querySelector("#spreadsheet tbody");
  const newRowIndex = tbody.rows.length;
  const colCount = document.querySelector("#spreadsheet thead tr").cells.length - 1;
  const tr = document.createElement("tr");

  // Row header
  const th = document.createElement("th");
  th.textContent = newRowIndex + 1;
  tr.appendChild(th);

  // Create new cells
  for (let c = 0; c < colCount; c++) {
    const td = document.createElement("td");
    td.contentEditable = "true";
    td.addEventListener("input", () => onCellInput(newRowIndex, c, td));
    td.addEventListener("focus", () => onCellFocus(newRowIndex, c, td));
    tr.appendChild(td);
  }
  tbody.appendChild(tr);

  // Expand data model
  spreadsheetData.push(Array.from({ length: colCount }, () => ({ value: "", formula: "" })));
}

function deleteRow() {
  const tbody = document.querySelector("#spreadsheet tbody");
  if (tbody.rows.length > 0) {
    tbody.removeChild(tbody.lastChild);
    spreadsheetData.pop();
  }
}

function addColumn() {
  const thead = document.querySelector("#spreadsheet thead tr");
  const newColIndex = thead.cells.length - 1;
  const th = document.createElement("th");
  th.textContent = columnIndexToName(newColIndex);
  thead.appendChild(th);

  const tbody = document.querySelector("#spreadsheet tbody");
  for (let r = 0; r < tbody.rows.length; r++) {
    const td = document.createElement("td");
    td.contentEditable = "true";
    td.addEventListener("input", () => onCellInput(r, newColIndex, td));
    td.addEventListener("focus", () => onCellFocus(r, newColIndex, td));
    tbody.rows[r].appendChild(td);
    spreadsheetData[r].push({ value: "", formula: "" });
  }
}

function deleteColumn() {
  const thead = document.querySelector("#spreadsheet thead tr");
  const colCount = thead.cells.length - 1;
  if (colCount < 1) return;

  thead.removeChild(thead.lastChild);

  const tbody = document.querySelector("#spreadsheet tbody");
  for (let r = 0; r < tbody.rows.length; r++) {
    tbody.rows[r].removeChild(tbody.rows[r].lastChild);
    spreadsheetData[r].pop();
  }
}

function resizeRow() {
  const rowStart = parseInt(document.getElementById("row-start").value, 10);
  const rowEnd = parseInt(document.getElementById("row-end").value, 10);
  if (isNaN(rowStart) || isNaN(rowEnd)) return;

  for (let r = rowStart; r <= rowEnd; r++) {
    const rowElem = document.querySelector("#spreadsheet tbody").rows[r - 1];
    if (rowElem) {
      rowElem.style.height = "40px"; // or prompt user for a custom size
    }
  }
}

function resizeColumn() {
  const colInput = document.getElementById("col-select").value.toUpperCase();
  if (!colInput) return;
  const columns = colInput.split(",").map(s => s.trim());
  columns.forEach(colLetter => {
    const cIndex = colNameToIndex(colLetter);
    const theadCell = document.querySelector("#spreadsheet thead tr").cells[cIndex + 1];
    if (theadCell) {
      theadCell.style.width = "150px";
    }
    const tbodyRows = document.querySelectorAll("#spreadsheet tbody tr");
    tbodyRows.forEach(row => {
      const cell = row.cells[cIndex + 1];
      if (cell) {
        cell.style.width = "150px";
      }
    });
  });
}

/* -------------------------------------------
 *         SAVE/LOAD (LocalStorage)
 * -----------------------------------------*/
function saveSheet() {
  localStorage.setItem("spreadsheetData", JSON.stringify(spreadsheetData));
  alert("Spreadsheet saved to Local Storage!");
}

function loadSheet() {
  const dataStr = localStorage.getItem("spreadsheetData");
  if (!dataStr) {
    alert("No saved spreadsheet found!");
    return;
  }
  const data = JSON.parse(dataStr);
  if (!Array.isArray(data)) {
    alert("Invalid data!");
    return;
  }

  const rows = data.length;
  const cols = data[0].length;
  generateTable(rows, cols);
  initSpreadsheetData(rows, cols);

  // Fill the loaded data
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      spreadsheetData[r][c] = {
        value: data[r][c].value,
        formula: data[r][c].formula
      };
      const td = getCellElement(r, c);
      td.innerText = data[r][c].formula ? data[r][c].formula : data[r][c].value;
    }
  }
  recalcAll();
  alert("Spreadsheet loaded from Local Storage!");
}

/* -------------------------------------------
 *      EXTRA EVENT LISTENERS (Optional)
 * -----------------------------------------*/
function attachEventListeners() {
  // For example, handle clicks outside the table, etc.
}

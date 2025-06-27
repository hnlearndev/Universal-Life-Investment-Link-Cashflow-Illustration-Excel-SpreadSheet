function main(workbook: ExcelScript.Workbook) {
  let startTime = new Date();

  // Set the calculation mode to manual.
  workbook
    .getApplication()
    .setCalculationMode(ExcelScript.CalculationMode.manual);

  // Input validation
  let input_validation = validate_input(workbook);

  if (!input_validation) {
    return;
  }

  // Base cashflow
  baseCF(workbook);

  // Rider cashflow
  riderCF(workbook);

  // Return automatic calculation
  workbook
    .getApplication()
    .setCalculationMode(ExcelScript.CalculationMode.automatic);

  // Print time stamp
  _printTimeStamp(startTime, workbook.getWorksheet("Input").getRange("M1"));
}

// ------------------------------------------------------------------------------
// INPUT VALIDATION
// ------------------------------------------------------------------------------
function validate_input(workbook: ExcelScript.Workbook): Boolean {
  let wshInput = workbook.getWorksheet("Input");

  // Exit if validation is not met
  let validation = Boolean(wshInput.getRange("input_validation").getValue());

  if (!validation) {
    // Clear time stamp
    wshInput.getRange("M1:M3").clear(ExcelScript.ClearApplyTo.contents);

    // Print error message
    wshInput
      .getRange("M1")
      .setValue("Input validation failed. Please check the input data.");
  }

  return validation;
}

// ------------------------------------------------------------------------------
// BASE CASHFLOW
// ------------------------------------------------------------------------------

/**
 * Main function to run the Base Cash Flow calculations.
 * It normalizes the withdrawal table, generates scenario combinations,
 * adds rows to result tables, performs cash flow calculations, and checks results.
 */
function baseCF(workbook: ExcelScript.Workbook) {
  // Start time track
  let startTime = new Date();

  // Worksheet objects
  let wsh = workbook.getWorksheet("Base_CF");

  let tblBaseCF = wsh.getTable("tblBaseCF");
  let tblBaseCFSummary = wsh.getTable("tblBaseCFSummary");
  let tblBaseCFResult = wsh.getTable("tblBaseCFResult");

  let term = Number(wsh.getRange("term").getValue());
  let base = String(wsh.getRange("base").getValue());

  // Clear time stamp
  wsh.getRange("B1:B3").clear(ExcelScript.ClearApplyTo.contents);

  // Remove filter
  _removeFilters(wsh);

  // Normailize the withdrawal table
  _normalize_withdrawal_table(workbook);

  // Get scenario combinations
  let combinations = _scenarioCombination(base);

  // Add rows to result tables
  let totalBaseCFRows = combinations.length * term;
  _addRowsToTable(tblBaseCFResult, totalBaseCFRows);

  // Track of rows
  let baseCFResultRows = 0;

  for (let i = 0; i < combinations.length; i++) {
    // Single scenario
    let scenario = combinations[i];

    // Setup inputs
    wsh.getRange("int_rate_scenario").setValue(scenario.IntRate);
    wsh.getRange("risk_scenario").setValue(scenario.RiskType);
    wsh.getRange("prem_term_scenario").setValue(scenario.PremTerm);

    // Manual calculate output ranges and tables
    wsh.getRange("int_rate_scenario").getResizedRange(9, 0).calculate();
    tblBaseCF.getRangeBetweenHeaderAndTotal().calculate();
    tblBaseCFSummary.getRangeBetweenHeaderAndTotal().calculate();

    // Copy table Summary to table Result
    _copySummaryToResult(tblBaseCF, term, baseCFResultRows);

    // Update baseCFResultRows for the next iteration
    baseCFResultRows += term;
  }

  // ExitHere: Clear scenario input cells
  wsh.getRange("int_rate_scenario").setValue("");
  wsh.getRange("risk_scenario").setValue("");
  wsh.getRange("prem_term_scenario").setValue("");

  // Print time stamp
  _printTimeStamp(startTime, wsh.getRange("B1"));
}

// ----------------------------------
// _normalize_withdrawal_table function
// ----------------------------------

/**
 * Normalizes the withdrawal table by expanding the input rows into output rows.
 * Each input row with a range of years will create multiple output rows for each year.
 */
function _normalize_withdrawal_table(workbook: ExcelScript.Workbook) {
  // Get the input and output tables
  let wshInput = workbook.getWorksheet("Input");
  let tblInput = wshInput.getTable("tblWithdrawal");
  let tblOutput = wshInput.getTable("tblWithdrawal_Output");

  // Calculate total output rows by expanding each input row's year range
  let inputRows = tblInput.getRangeBetweenHeaderAndTotal().getValues();

  let totalOutputRows = inputRows.reduce((sum, row) => {
    if (row[0] === "" || row[0] === null || row[0] === undefined) return sum;
    return sum + (Number(row[1]) - Number(row[0]) + 1);
  }, 0);

  // Add rows to the output table if needed
  _addRowsToTable(tblOutput, totalOutputRows);

  // Normalize the withdrawal table
  let outputRows: (number | string)[][] = [];

  for (let row of inputRows) {
    // Quit if first cell is empty
    if (row[0] === "" || row[0] === null || row[0] === undefined) {
      break;
    }

    let startYear = Number(row[0]);
    let endYear = Number(row[1]);
    let amount = Number(row[2]);

    for (let year = startYear; year <= endYear; year++) {
      outputRows.push([year, amount]);
    }
  }

  // Add all output rows to the output table
  if (outputRows.length == 0) {
    return; // No output rows to add
  }

  // Clear existing output rows
  let outputRange = tblOutput
    .getRangeBetweenHeaderAndTotal()
    .getCell(0, 0)
    .getResizedRange(totalOutputRows - 1, 1);
  outputRange.setValues(outputRows);
}

// ----------------------------------
// _scenarioCombination function
// ----------------------------------

/**
 * Represents a scenario for running cash flow calculations.
 */
interface RunScenario {
  IntRate: string;
  RiskType: string;
  PremTerm: string;
}

/**
 * Generates scenario combinations based on the product in the "product" named range.
 * Returns an array of scenario objects.
 */
function _scenarioCombination(product: string): RunScenario[] {
  // Define arrays for interest rates, premium terms, and risk types
  let arrIntRate: string[] = [];
  let arrPremTerm: string[] = [];
  let arrRisk: string[] = ["Standard", "Subrisk"];

  switch (product) {
    case "UVL01":
    case "UVL02":
    case "UVL03":
      arrIntRate = ["High", "Low", "Guaranteed"];
      arrPremTerm = ["Opted term"];
      break;
    case "ILP01":
      arrIntRate = ["High", "Low"];
      arrPremTerm = ["Must-pay term", "Policy term", "Opted term"];
      break;
    default:
      return [];
  }

  let combinations: RunScenario[] = [];

  for (let IntRate of arrIntRate) {
    for (let PremTerm of arrPremTerm) {
      for (let RiskType of arrRisk) {
        combinations.push({
          IntRate,
          RiskType,
          PremTerm,
        });
      }
    }
  }

  return combinations;
}

// ------------------------------------------------------------------------------
// RIDER CASHFLOW
// ------------------------------------------------------------------------------

/**
 * Main function to run the Rider Cash Flow calculations.
 * It initializes the rider rows, adds rows to the result table, and runs the rider calculations.
 * Finally, it prints the time stamp and clears input cells.
 */
function riderCF(workbook: ExcelScript.Workbook) {
  let startTime = new Date();

  let wsh = workbook.getWorksheet("Rider_CF");

  // Clear time stamp
  wsh.getRange("B1:D1").clear(ExcelScript.ClearApplyTo.contents);

  // Remove filter
  _removeFilters(wsh);

  // Avoid running over table rider twice
  let riderRows = _riderRows(workbook);

  // Add rows to the Rider Result table
  _addRowsToTable(wsh.getTable("tblRiderCFResult"), riderRows[0]);

  // Track result row count
  let riderCFResultRowsCount = 0;

  // Run Non-WOP first and return row count
  riderCFResultRowsCount = _riderRun(
    workbook,
    riderRows[1],
    riderCFResultRowsCount
  );

  // Run WOP
  riderCFResultRowsCount = _riderRun(
    workbook,
    riderRows[2],
    riderCFResultRowsCount
  );

  // Clear input and risk type cells
  wsh.getRange("input").setValue("");
  wsh.getRange("risk").setValue("");

  // Print time stamp
  _printTimeStamp(startTime, wsh.getRange("B1"));
}

// ----------------------------------
// function _riderRun
// ----------------------------------

/**
 * Runs the Rider Cash Flow calculations for both risk types (Standard and Sub-standard).
 * It fills in the rider information, calculates the cash flow, and copies results to the result table.
 *
 * @param workbook - The Excel workbook object.
 * @param rows - An array of row indices to process.
 */
function _riderRun(
  workbook: ExcelScript.Workbook,
  rows: number[],
  idx: number
): number {
  let wshInput = workbook.getWorksheet("Input");
  let riderTable = wshInput.getTable("tblRider");
  let data = riderTable.getRangeBetweenHeaderAndTotal().getValues();

  let wsh = workbook.getWorksheet("Rider_CF");
  let tbl = wsh.getTable("tblRiderCF");
  let tblSummary = wsh.getTable("tblRiderCFSummary");

  let arrRisk = ["Standard", "Sub-standard"];

  let totalRows = idx;

  for (let i = 0; i < arrRisk.length; i++) {
    // Fill in risk type
    wsh.getRange("risk").setValue(arrRisk[i]);

    for (let j = 0; j < rows.length; j++) {
      let idx = rows[j];

      // Fill in rider info
      wsh.getRange("input").setValues([data[idx]]);

      // Caculate output ranges and tables
      // Note: There are cells not in the tables containing modal factors needs to be refresh for each rider
      wsh.getRange("gender").getResizedRange(3, 0).getEntireRow().calculate();
      tbl.getRangeBetweenHeaderAndTotal().calculate();
      tblSummary.getRangeBetweenHeaderAndTotal().calculate();

      // Identify rider term
      let term = Number(wsh.getRange("term").getValue());

      // Copy from Summary to Result table
      _copySummaryToResult(tbl, term, totalRows);

      totalRows += term;
    }
  }

  // Return row count
  return totalRows;
}

// ----------------------------------
// function _riderRows
// ----------------------------------
function _riderRows(
  workbook: ExcelScript.Workbook
): [number, number[], number[]] {
  // Rider_CF sheet
  let wsh = workbook.getWorksheet("Rider_CF");
  let input = wsh.getRange("input");

  // From Input sheet
  let tblRider = workbook.getWorksheet("Input").getTable("tblRider");
  let data = tblRider.getRangeBetweenHeaderAndTotal().getValues();

  let wopArr: number[] = [];
  let nonWopArr: number[] = [];
  let totalRows = 0;

  for (let i = 0; i < data.length; i++) {
    const rider = data[i][0];

    // Exit on the first empty rider row
    if (rider === "" || rider === null || rider === undefined) {
      break;
    }

    // Filling info in Rider_CF using Input info
    input.setValues([data[i]]);

    // Calculate inputs
    input.getResizedRange(0, 6).calculate();

    // Collect 1-based row index
    let isWOP = Boolean(wsh.getRange("is_waiver").getValue());

    if (isWOP) {
      wopArr.push(i);
    } else {
      nonWopArr.push(i);
    }

    // Calculate the rider term
    let term = Number(wsh.getRange("term").getValue());
    totalRows = totalRows + term;
  }

  // Multiply by 2 for both risk types
  return [totalRows * 2, nonWopArr, wopArr];
}

// ------------------------------------------------------------------------------
// OTHER HELPERS
// ------------------------------------------------------------------------------

/**
 * Adds empty rows to a specified result table to ensure it has the required number of rows.
 *
 * @param tblResult - The ExcelScript.Table object to which rows will be added.
 * @param totalRows - The total number of rows that the table should have after adding.
 */
function _addRowsToTable(tblResult: ExcelScript.Table, totalRows: number) {
  // Clear data rows in result tables
  tblResult
    .getRangeBetweenHeaderAndTotal()
    .clear(ExcelScript.ClearApplyTo.all);

  // Current rows in each table
  let currentRows = tblResult.getRangeBetweenHeaderAndTotal().getRowCount();

  // No need to add rows
  if (currentRows >= totalRows) {
    return;
  }

  // Efficiently add rows to Base CF Result table if needed
  let rowsToAdd = totalRows - currentRows;
  let emptyRow: string[] = new Array(
    tblResult.getHeaderRowRange().getColumnCount()
  ).fill("");
  let newRows = Array.from({ length: rowsToAdd }, () => emptyRow.slice());
  tblResult.addRows(-1, newRows);
}

/**
 * Removes any existing filters from the worksheet and its tables.
 *
 * @param wsh - The ExcelScript.Worksheet object from which filters will be removed.
 */
function _removeFilters(wsh: ExcelScript.Worksheet) {
  // Remove any existing filters from the worksheet
  let autoFilter = wsh.getAutoFilter();
  if (autoFilter) {
    autoFilter.remove();
  }
  // Clear any existing filters from the table
  let tables = wsh.getTables();
  for (let tbl of tables) {
    if (tbl.getAutoFilter()) {
      tbl.getAutoFilter().remove();
    }
  }
}
/**
 * Copies summary values from a calculation table to a result table for a specific term.
 *
 * @param tblCalculation - The ExcelScript.Table containing the calculation data -> use to infer the summary table and result table names.
 * @param term - The term for which summary values are copied.
 * @param idx - The index in the result table where the summary values will be placed.
 */
function _copySummaryToResult(
  tbl: ExcelScript.Table,
  term: number,
  idx: number
) {
  let wsh = tbl.getWorksheet();
  let tblName = tbl.getName();

  // Get the result and summary tables
  let summaryTable = wsh.getTable(tblName + "Summary");
  let resultTable = wsh.getTable(tblName + "Result");

  // Get summary values for the specified term
  let summaryRowCount = summaryTable.getRangeBetweenHeaderAndTotal().getRowCount();
  let summaryValues = summaryTable.getRangeBetweenHeaderAndTotal().getResizedRange(term - summaryRowCount,0).getValues();

  // Set values to the result table, starting from the first data row
  let resultColumnCount = resultTable.getHeaderRowRange().getColumnCount(); // Shape of summary and result tables should be the same
  let resultRange = resultTable
    .getRangeBetweenHeaderAndTotal()
    .getCell(idx, 0)
    .getResizedRange(term - 1, resultColumnCount - 1);

  resultRange.setValues(summaryValues);
}

/**
 * Prints the time stamp and duration of the script run to a specified cell.
 *
 * @param startTime - The start time of the script run.
 * @param cell - The ExcelScript.Range where the time stamp will be printed.
 */
function _printTimeStamp(startTime: Date, cell: ExcelScript.Range) {
  let endTime = new Date();

  let durationSeconds = (
    (endTime.getTime() - startTime.getTime()) /
    1000
  ).toFixed(2);

  cell.setValue("Started: " + startTime.toLocaleString());

  cell.getOffsetRange(1, 0).setValue("Finished: " + startTime.toLocaleString());

  cell
    .getOffsetRange(2, 0)
    .setValue("Duration: " + durationSeconds + " seconds");
}

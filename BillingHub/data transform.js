function splitGrossPitFromFormulas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const startRow = 2;
  const sourceCol = 7; // О
  const targetCol = 10; // J
  const numRows = sheet.getLastRow() - startRow + 1;
  if (numRows < 1) return;

  const sourceRange = sheet.getRange(startRow, sourceCol, numRows, 1);
  const formulas = sourceRange.getFormulas();
  const values = sourceRange.getValues();

  const output = [];

  for (let i = 0; i < numRows; i++) {
    const formula = formulas[i][0];
    const value = values[i][0];

    if (formula) {
      const expr = formula.replace(/^=/, "").replace(/\s+/g, "");
      const parts = expr.split("+");

      let pit = "";

      if (parts.length === 1) {
        pit = "";
      } else if (parts.length === 2) {
        pit = toNumber(parts[1]);
      } else if (parts.length >= 3) {
        const second = parts[1];
        const third = parts[2];

        // правило:
        // якщо 2-й елемент десятковий, а 3-й цілий -> PIT = 3-й
        // інакше PIT = 2-й
        if (isDecimalLiteral(second) && isIntegerLiteral(third)) {
          pit = toNumber(third);
        } else {
          pit = toNumber(second);
        }
      }

      const gross = pit === "" ? value : Number(value) - Number(pit);
      output.push([gross, pit]);
    } else {
      // не формула, просто значення
      if (value === "" || value === null) {
        output.push(["", ""]);
      } else if (Number(value) === 0) {
        output.push([0, ""]);
      } else {
        output.push([value, ""]);
      }
    }
  }

  sheet.getRange(1, targetCol, 1, 2).setValues([["Gross", "PIT"]]);
  sheet.getRange(startRow, targetCol, output.length, 2).setValues(output);
}

function isDecimalLiteral(str) {
  return /^-?\d+\.\d+$/.test(str);
}

function isIntegerLiteral(str) {
  return /^-?\d+$/.test(str);
}

function toNumber(str) {
  const n = Number(str);
  return isNaN(n) ? "" : n;
}
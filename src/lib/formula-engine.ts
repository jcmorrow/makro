import { run, parse, format } from "./ok";

interface CellData {
  formula: string;
  value: any;
  error?: string;
}

export function getCellLabel(col: number, row: number): string {
  let label = "";
  let c = col;
  while (c >= 0) {
    label = String.fromCharCode(65 + (c % 26)) + label;
    c = Math.floor(c / 26) - 1;
  }
  return label + (row + 1);
}

export function parseCellReference(
  ref: string,
): { col: number; row: number } | null {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;

  const colStr = match[1];
  const rowStr = match[2];

  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 65 + 1);
  }
  col -= 1;

  const row = parseInt(rowStr) - 1;

  return { col, row };
}

export function evaluateFormula(
  formula: string,
  cells: Record<string, CellData>,
  cellId: string,
): { value: any; error?: string } {
  // If it doesn't start with =, treat as literal
  if (!formula.startsWith("=")) {
    return { value: formula };
  }

  // Remove the = sign
  let expression = formula.slice(1);

  try {
    console.log("expression:", expression);
    expression = expression.replace(/\bUP\b/g, (match) => {
      const cellAddress = parseCellReference(cellId);
      return getCellLabel(cellAddress!.col, cellAddress!.row - 1);
    });
    console.log("expression2:", expression);
    // Replace cell references with their values
    expression = expression.replace(/[A-Z]+\d+/g, (match) => {
      const cell = cells[match];
      if (!cell) {
        return "undefined";
      }
      if (cell.error) {
        throw new Error(`Cell ${match} has error: ${cell.error}`);
      }
      // Convert value to JS literal
      return cell.value;
    });
    console.log("expression3:", expression);

    const result = format(run(parse(expression)));

    return { value: result };
  } catch (error) {
    console.log(error);
    console.error(error);
    return {
      value: null,
      error: error instanceof Error ? error.message : "Invalid formula",
    };
  }
}

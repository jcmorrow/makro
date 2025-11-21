import { run, parse, format } from "./ok";
import { getCellLabel, parseCellReference } from "./util";

interface CellData {
  formula: string;
  value: any;
  error?: string;
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
    expression = expression.replace(/\bUP\b/g, (match) => {
      const cellAddress = parseCellReference(cellId);
      return getCellLabel(cellAddress!.col, cellAddress!.row - 1);
    });
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

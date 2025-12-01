import { run, parse, format } from "./ok";
import { getCellLabel, iterateCellRange, parseCellReference } from "./util";

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
    // I think ideally this would actually also respect *shape*. Right now a 2x5
    // range becomes a 1x10 range.
    expression = expression.replace(
      /@\[([A-Z]+\d+):([A-Z]+\d+)]/g,
      (match, start, end) => {
        return [...iterateCellRange(start, end)]
          .map((cell) => cells[cell].value)
          .join(" ");
      },
    );
    // TODO: Implement RIGHT, DOWN
    expression = expression.replace(/\bUP\b/g, (match) => {
      const cellAddress = parseCellReference(cellId);
      return getCellLabel(cellAddress!.col, Math.max(0, cellAddress!.row - 1));
    });
    expression = expression.replace(/\bLEFT\b/g, (match) => {
      const cellAddress = parseCellReference(cellId);
      return getCellLabel(Math.max(0, cellAddress!.col - 1), cellAddress!.row);
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
    console.log("Error running expression: ", expression);
    console.error(error);
    return {
      value: null,
      error: error instanceof Error ? error.message : "Invalid formula",
    };
  }
}

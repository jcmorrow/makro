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
): { value: any; error?: string } {
  // If it doesn't start with =, treat as literal
  if (!formula.startsWith("=")) {
    // Try to parse as number
    const num = parseFloat(formula);
    if (!isNaN(num) && formula.trim() !== "") {
      return { value: num };
    }
    // Otherwise return as string
    return { value: formula };
  }

  // Remove the = sign
  let expression = formula.slice(1);

  try {
    // Replace cell references with their values
    expression = expression.replace(/\b([A-Z]+\d+)\b/g, (match) => {
      const cell = cells[match];
      if (!cell) {
        return "undefined";
      }
      if (cell.error) {
        throw new Error(`Cell ${match} has error: ${cell.error}`);
      }
      // Convert value to JS literal
      return JSON.stringify(cell.value);
    });

    // Add some useful functions
    const context = {
      r: (n: number) => Array.from({ length: n }, (_, i) => i),
      sum: (arr: number[]) => arr.reduce((a, b) => a + b, 0),
      avg: (arr: number[]) => arr.reduce((a, b) => a + b, 0) / arr.length,
      max: (arr: number[]) => Math.max(...arr),
      min: (arr: number[]) => Math.min(...arr),
      len: (arr: any[]) => arr.length,
      first: (arr: any[], n: number = 1) => arr.slice(0, n),
      skip: (arr: any[], n: number = 0) => arr.slice(n),
      Array,
      Math,
      String,
      Number,
      Boolean,
      JSON,
      Date,
      Object,
      parseInt,
      parseFloat,
      isNaN,
      isFinite,
    };

    // Create function with context
    const func = new Function(
      ...Object.keys(context),
      `"use strict"; return (${expression});`,
    );

    const result = func(...Object.values(context));

    return { value: result };
  } catch (error) {
    return {
      value: null,
      error: error instanceof Error ? error.message : "Invalid formula",
    };
  }
}

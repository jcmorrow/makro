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

export function getCellLabel(col: number, row: number): string {
  let label = "";
  let c = col;
  while (c >= 0) {
    label = String.fromCharCode(65 + (c % 26)) + label;
    c = Math.floor(c / 26) - 1;
  }
  return label + (row + 1);
}

export function* iterateCellRange(startString: string, endString: string) {
  const start = parseCellReference(startString)!;
  const end = parseCellReference(endString)!;

  const minCol = Math.min(start.col, end.col);
  const maxCol = Math.max(start.col, end.col);
  const minRow = Math.min(start.row, end.row);
  const maxRow = Math.max(start.row, end.row);

  for (let c = minCol; c <= maxCol; c++) {
    for (let r = minRow; r <= maxRow; r++) {
      yield getCellLabel(c, r);
    }
  }
}

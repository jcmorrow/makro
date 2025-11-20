import React, { useState, useCallback, useEffect, useRef } from "react";
import { Code2 } from "lucide-react";
import { evaluateFormula } from "../lib/formula-engine";

const ROWS = 30;
const COLS = 15;

interface CellData {
  formula: string;
  value: any;
  error?: string;
}

export function Spreadsheet() {
  const [cells, setCells] = useState<Record<string, CellData>>({});
  const [selectedCell, setSelectedCell] = useState<string | null>("A1");
  const [selectionStart, setSelectionStart] = useState<string | null>(null);
  const [selectionEnd, setSelectionEnd] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [editingCell, setEditingCell] = useState<string | null>(null);
  const [formulaBarValue, setFormulaBarValue] = useState("");
  const [mode, setMode] = useState<"normal" | "insert" | "visual">("normal");
  const inputRef = useRef<HTMLInputElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);

  // Initialize with some example data
  useEffect(() => {
    const savedCellsSerialized = localStorage.getItem("spreadsheet");
    if (savedCellsSerialized) {
      try {
        const savedCells = JSON.parse(savedCellsSerialized);
        console.log({ savedCells, savedCellsSerialized });
        if (!savedCells) {
          console.error("Invalid saved spreadsheet data");
          localStorage.setItem("spreadsheetBackup", savedCellsSerialized);
          localStorage.removeItem("spreadsheet");
          return;
        }
        setCells(savedCells);
        setFormulaBarValue(savedCells["A1"].formula || "");
      } catch (error) {
        console.error("Error parsing saved spreadsheet data:", error);
      }
    } else {
      const examples: Record<string, string> = {
        A1: "=!5",
        A2: "=+/A1",
        A3: "=!A2",
      };

      const initialCells: Record<string, CellData> = {};
      Object.entries(examples).forEach(([cellId, formula]) => {
        const result = evaluateFormula(formula, initialCells, cellId);
        initialCells[cellId] = {
          formula,
          value: result.value,
          error: result.error,
        };
      });
      setCells(initialCells);
      setFormulaBarValue(examples["A1"] || "");
    }

    // Focus container on mount
    setTimeout(() => containerRef.current?.focus(), 100);
  }, []);

  const recalculateAll = useCallback((newCells: Record<string, CellData>) => {
    const maxIterations = 100;
    let iteration = 0;
    let hasChanges = true;

    while (hasChanges && iteration < maxIterations) {
      hasChanges = false;
      iteration++;

      Object.keys(newCells).forEach((cellId) => {
        const cell = newCells[cellId];
        const result = evaluateFormula(cell.formula, newCells, cellId);

        if (result.value !== cell.value || result.error !== cell.error) {
          newCells[cellId] = {
            ...cell,
            value: result.value,
            error: result.error,
          };
          hasChanges = true;
        }
      });
    }

    return newCells;
  }, []);

  const updateCell = useCallback(
    (cellId: string, formula: string) => {
      setCells((prevCells: Record<string, CellData>) => {
        const newCells = { ...prevCells };

        if (formula === "") {
          delete newCells[cellId];
        } else {
          newCells[cellId] = {
            formula,
            value: null,
            error: undefined,
          };
        }

        const cells = recalculateAll(newCells);
        localStorage.setItem("spreadsheet", JSON.stringify(cells));
        return cells;
      });
    },
    [recalculateAll],
  );

  // Helper function to check if a cell is in the selection range
  const isCellInSelection = useCallback(
    (cellId: string) => {
      if (!selectionStart || !selectionEnd) return false;

      const startCol = selectionStart.charCodeAt(0) - 65;
      const startRow = parseInt(selectionStart.slice(1)) - 1;
      const endCol = selectionEnd.charCodeAt(0) - 65;
      const endRow = parseInt(selectionEnd.slice(1)) - 1;

      const cellCol = cellId.charCodeAt(0) - 65;
      const cellRow = parseInt(cellId.slice(1)) - 1;

      const minCol = Math.min(startCol, endCol);
      const maxCol = Math.max(startCol, endCol);
      const minRow = Math.min(startRow, endRow);
      const maxRow = Math.max(startRow, endRow);

      return (
        cellCol >= minCol &&
        cellCol <= maxCol &&
        cellRow >= minRow &&
        cellRow <= maxRow
      );
    },
    [selectionStart, selectionEnd],
  );

  const handleCellClick = useCallback(
    (cellId: string, _e: React.MouseEvent) => {
      setSelectedCell(cellId);
      setSelectionStart(null);
      setSelectionEnd(null);
      setEditingCell(null);
      setMode("normal");
      const cell = cells[cellId];
      setFormulaBarValue(cell?.formula || "");
      // Focus the container so it can receive keyboard events
      setTimeout(() => containerRef.current?.focus(), 0);
    },
    [cells],
  );

  const handleCellMouseDown = useCallback(
    (cellId: string) => {
      setSelectedCell(cellId);
      setSelectionStart(cellId);
      setSelectionEnd(cellId);
      setIsDragging(true);
      setEditingCell(null);
      setMode("visual");
      const cell = cells[cellId];
      setFormulaBarValue(cell?.formula || "");
      setTimeout(() => containerRef.current?.focus(), 0);
    },
    [cells],
  );

  const handleCellMouseEnter = useCallback(
    (cellId: string) => {
      if (isDragging) {
        setSelectionEnd(cellId);
      }
    },
    [isDragging],
  );

  const handleMouseUp = useCallback(() => {
    if (isDragging) {
      setIsDragging(false);
      // If only one cell selected, go back to normal mode
      if (selectionStart === selectionEnd) {
        setMode("normal");
        setSelectionStart(null);
        setSelectionEnd(null);
      }
    }
  }, [isDragging, selectionStart, selectionEnd]);

  useEffect(() => {
    document.addEventListener("mouseup", handleMouseUp);
    return () => {
      document.removeEventListener("mouseup", handleMouseUp);
    };
  }, [handleMouseUp]);

  const handleCellDoubleClick = useCallback((cellId: string) => {
    setEditingCell(cellId);
    setSelectedCell(cellId);
    setMode("insert");
  }, []);

  const handleCellEdit = useCallback(
    (cellId: string, value: string) => {
      updateCell(cellId, value);
      setEditingCell(null);
      setMode("normal");
      setTimeout(() => containerRef.current?.focus(), 0);
    },
    [updateCell],
  );

  const handleFormulaBarChange = useCallback((value: string) => {
    setFormulaBarValue(value);
  }, []);

  const handleFormulaBarSubmit = useCallback(() => {
    if (selectedCell) {
      updateCell(selectedCell, formulaBarValue);
      setTimeout(() => containerRef.current?.focus(), 0);
    }
  }, [selectedCell, formulaBarValue, updateCell]);

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent) => {
      if (!selectedCell) return;

      const [col, row] = [
        selectedCell.charCodeAt(0) - 65,
        parseInt(selectedCell.slice(1)) - 1,
      ];
      const isShiftPressed = e.shiftKey;

      // Handle Shift for visual selection
      const handleNavigation = (newCol: number, newRow: number) => {
        const newCell = String.fromCharCode(65 + newCol) + (newRow + 1);

        if (isShiftPressed) {
          // Enter visual mode if not already
          if (mode !== "visual") {
            setMode("visual");
            setSelectionStart(selectedCell);
          }
          if (!selectionStart) {
            setSelectionStart(selectedCell);
          }
          setSelectionEnd(newCell);
        } else {
          // Normal navigation - exit visual mode
          if (mode === "visual") {
            setMode("normal");
            setSelectionStart(null);
            setSelectionEnd(null);
          }
        }

        setSelectedCell(newCell);
        setFormulaBarValue(cells[newCell]?.formula || "");
      };

      // Vim keybindings in normal mode
      if ((mode === "normal" || mode === "visual") && !editingCell) {
        // Navigation: h, j, k, l (with optional Shift for visual mode)
        if (e.key === "h" || (e.key === "H" && isShiftPressed)) {
          e.preventDefault();
          const newCol = Math.max(col - 1, 0);
          handleNavigation(newCol, row);
          return;
        } else if (e.key === "j" || (e.key === "J" && isShiftPressed)) {
          e.preventDefault();
          const newRow = Math.min(row + 1, ROWS - 1);
          handleNavigation(col, newRow);
          return;
        } else if (e.key === "k" || (e.key === "K" && isShiftPressed)) {
          e.preventDefault();
          const newRow = Math.max(row - 1, 0);
          handleNavigation(col, newRow);
          return;
        } else if (e.key === "l" || (e.key === "L" && isShiftPressed)) {
          e.preventDefault();
          const newCol = Math.min(col + 1, COLS - 1);
          handleNavigation(newCol, row);
          return;
        }
        // Arrow keys with Shift
        else if (e.key === "ArrowLeft") {
          e.preventDefault();
          const newCol = Math.max(col - 1, 0);
          handleNavigation(newCol, row);
          return;
        } else if (e.key === "ArrowRight") {
          e.preventDefault();
          const newCol = Math.min(col + 1, COLS - 1);
          handleNavigation(newCol, row);
          return;
        } else if (e.key === "ArrowUp") {
          e.preventDefault();
          const newRow = Math.max(row - 1, 0);
          handleNavigation(col, newRow);
          return;
        } else if (e.key === "ArrowDown") {
          e.preventDefault();
          const newRow = Math.min(row + 1, ROWS - 1);
          handleNavigation(col, newRow);
          return;
        }

        // Only in normal mode (not visual)
        if (mode === "normal") {
          // Enter insert mode: i
          if (e.key === "=") {
            e.preventDefault();
            updateCell(selectedCell, "=");
            setFormulaBarValue("=");
            setEditingCell(selectedCell);
            setMode("insert");
            return;
          }
          if (e.key === "i") {
            e.preventDefault();
            setEditingCell(selectedCell);
            setMode("insert");
            return;
          }
          // Enter insert mode at end: a
          else if (e.key === "a") {
            e.preventDefault();
            setEditingCell(selectedCell);
            setMode("insert");
            return;
          }
          // Delete cell content: dd or x
          else if (e.key === "d" || e.key === "x") {
            e.preventDefault();
            updateCell(selectedCell, "");
            setFormulaBarValue("");
            return;
          }
          // Yank value: Y
          else if (e.key === "Y") {
            e.preventDefault();
            const cell = cells[selectedCell];
            if (cell) {
              navigator.clipboard.writeText(cell.value);
            }
            return;
          }
          // Yank formula: y
          else if (e.key === "y") {
            e.preventDefault();
            const cell = cells[selectedCell];
            if (cell) {
              navigator.clipboard.writeText(cell.formula);
            }
            return;
          }
          // Paste: p
          else if (e.key === "p") {
            e.preventDefault();
            navigator.clipboard.readText().then((text) => {
              updateCell(selectedCell, text);
              setFormulaBarValue(text);
            });
            return;
          }
          // Go to first column: 0
          else if (e.key === "0") {
            e.preventDefault();
            const newCell = "A" + (row + 1);
            setSelectedCell(newCell);
            setFormulaBarValue(cells[newCell]?.formula || "");
            return;
          }
          // Go to last column: $
          else if (e.key === "$") {
            e.preventDefault();
            const newCell = String.fromCharCode(65 + COLS - 1) + (row + 1);
            setSelectedCell(newCell);
            setFormulaBarValue(cells[newCell]?.formula || "");
            return;
          }
          // Go to first row: g
          else if (e.key === "g") {
            e.preventDefault();
            const newCell = String.fromCharCode(65 + col) + "1";
            setSelectedCell(newCell);
            setFormulaBarValue(cells[newCell]?.formula || "");
            return;
          }
          // Go to last row: G
          else if (e.key === "G") {
            e.preventDefault();
            const newCell = String.fromCharCode(65 + col) + ROWS;
            setSelectedCell(newCell);
            setFormulaBarValue(cells[newCell]?.formula || "");
            return;
          }
        }

        // Visual mode specific commands
        if (mode === "visual") {
          // Escape to exit visual mode
          if (e.key === "Escape") {
            e.preventDefault();
            setMode("normal");
            setSelectionStart(null);
            setSelectionEnd(null);
            return;
          }
          // Delete selection: d or x
          else if (e.key === "d" || e.key === "x") {
            e.preventDefault();
            // Delete all cells in selection
            if (selectionStart && selectionEnd) {
              const startCol = selectionStart.charCodeAt(0) - 65;
              const startRow = parseInt(selectionStart.slice(1)) - 1;
              const endCol = selectionEnd.charCodeAt(0) - 65;
              const endRow = parseInt(selectionEnd.slice(1)) - 1;

              const minCol = Math.min(startCol, endCol);
              const maxCol = Math.max(startCol, endCol);
              const minRow = Math.min(startRow, endRow);
              const maxRow = Math.max(startRow, endRow);

              for (let c = minCol; c <= maxCol; c++) {
                for (let r = minRow; r <= maxRow; r++) {
                  const cellId = String.fromCharCode(65 + c) + (r + 1);
                  updateCell(cellId, "");
                }
              }
            }
            setMode("normal");
            setSelectionStart(null);
            setSelectionEnd(null);
            setFormulaBarValue("");
            return;
          }
          // Yank values: Y
          else if (e.key === "Y") {
            e.preventDefault();
            // Copy all cells in selection
            if (selectionStart && selectionEnd) {
              const startCol = selectionStart.charCodeAt(0) - 65;
              const startRow = parseInt(selectionStart.slice(1)) - 1;
              const endCol = selectionEnd.charCodeAt(0) - 65;
              const endRow = parseInt(selectionEnd.slice(1)) - 1;

              const minCol = Math.min(startCol, endCol);
              const maxCol = Math.max(startCol, endCol);
              const minRow = Math.min(startRow, endRow);
              const maxRow = Math.max(startRow, endRow);

              let copyText = "";
              for (let r = minRow; r <= maxRow; r++) {
                for (let c = minCol; c <= maxCol; c++) {
                  const cellId = String.fromCharCode(65 + c) + (r + 1);
                  const cell = cells[cellId];
                  console.log({ cell });
                  copyText += (cell?.value || "") + "\t";
                }
                copyText += "\n";
              }
              navigator.clipboard.writeText(copyText);
            }
            setMode("normal");
            setSelectionStart(null);
            setSelectionEnd(null);
            return;
          }
          // Yank selection: y
          else if (e.key === "y") {
            e.preventDefault();
            // Copy all cells in selection
            if (selectionStart && selectionEnd) {
              const startCol = selectionStart.charCodeAt(0) - 65;
              const startRow = parseInt(selectionStart.slice(1)) - 1;
              const endCol = selectionEnd.charCodeAt(0) - 65;
              const endRow = parseInt(selectionEnd.slice(1)) - 1;

              const minCol = Math.min(startCol, endCol);
              const maxCol = Math.max(startCol, endCol);
              const minRow = Math.min(startRow, endRow);
              const maxRow = Math.max(startRow, endRow);

              let copyText = "";
              for (let r = minRow; r <= maxRow; r++) {
                for (let c = minCol; c <= maxCol; c++) {
                  const cellId = String.fromCharCode(65 + c) + (r + 1);
                  const cell = cells[cellId];
                  copyText += (cell?.formula || "") + "\t";
                }
                copyText += "\n";
              }
              navigator.clipboard.writeText(copyText);
            }
            setMode("normal");
            setSelectionStart(null);
            setSelectionEnd(null);
            return;
          }
        }
      }

      // In insert mode, ESC exits to normal mode
      if (mode === "insert" && editingCell && e.key === "Escape") {
        e.preventDefault();
        setEditingCell(null);
        setMode("normal");
        setFormulaBarValue(cells[selectedCell]?.formula || "");
        setTimeout(() => containerRef.current?.focus(), 0);
        return;
      }

      // Standard keybindings
      if (e.key === "Enter" && !isShiftPressed) {
        e.preventDefault();
        if (editingCell) {
          setMode("normal");
        }
        // Move down
        const newRow = Math.min(row + 1, ROWS - 1);
        const newCell = String.fromCharCode(65 + col) + (newRow + 1);
        setSelectedCell(newCell);
        setFormulaBarValue(cells[newCell]?.formula || "");
        setMode("normal");
        setSelectionStart(null);
        setSelectionEnd(null);
        setTimeout(() => containerRef.current?.focus(), 0);
      } else if (e.key === "Tab") {
        e.preventDefault();
        // Move right
        const newCol = Math.min(col + 1, COLS - 1);
        const newCell = String.fromCharCode(65 + newCol) + (row + 1);
        setSelectedCell(newCell);
        setFormulaBarValue(cells[newCell]?.formula || "");
        setEditingCell(null);
        setMode("normal");
        setSelectionStart(null);
        setSelectionEnd(null);
      } else if (e.key === "Escape" && !editingCell && mode !== "visual") {
        setEditingCell(null);
        setMode("normal");
        setFormulaBarValue(cells[selectedCell]?.formula || "");
      } else if (
        (e.key === "Delete" || e.key === "Backspace") &&
        !editingCell &&
        mode !== "insert"
      ) {
        e.preventDefault();
        if (mode === "visual" && selectionStart && selectionEnd) {
          // Delete all cells in selection
          const startCol = selectionStart.charCodeAt(0) - 65;
          const startRow = parseInt(selectionStart.slice(1)) - 1;
          const endCol = selectionEnd.charCodeAt(0) - 65;
          const endRow = parseInt(selectionEnd.slice(1)) - 1;

          const minCol = Math.min(startCol, endCol);
          const maxCol = Math.max(startCol, endCol);
          const minRow = Math.min(startRow, endRow);
          const maxRow = Math.max(startRow, endRow);

          for (let c = minCol; c <= maxCol; c++) {
            for (let r = minRow; r <= maxRow; r++) {
              const cellId = String.fromCharCode(65 + c) + (r + 1);
              updateCell(cellId, "");
            }
          }
        } else {
          updateCell(selectedCell, "");
        }
        setFormulaBarValue("");
      }
    },
    [
      selectedCell,
      editingCell,
      mode,
      cells,
      handleFormulaBarSubmit,
      updateCell,
      selectionStart,
      selectionEnd,
    ],
  );

  const columns = Array.from({ length: COLS }, (_, i) =>
    String.fromCharCode(65 + i),
  );

  return (
    <div
      className="flex flex-col h-screen"
      onKeyDown={handleKeyDown}
      ref={containerRef}
      tabIndex={0}
      style={{ outline: "none" }}
    >
      {/* Header */}
      <div className="bg-slate-900 border-b border-slate-700 p-4">
        <div className="flex items-center justify-between mb-4">
          <div className="flex items-center gap-2">
            <Code2 className="size-6 text-emerald-400" />
            <h1 className="text-slate-100">Makro</h1>
            <span className="text-slate-500 text-sm font-serif">
              A spreadsheet with vim-keybindings and formulas written in K
            </span>
          </div>
          <span className="text-slate-500 text-sm font-serif">
            Made by{" "}
            <a
              target="_blank"
              href="https://jcmorrow.com"
              className="text-emerald-400 hover:underline"
            >
              @jcmorrow
            </a>{" "}
            with support from{" "}
            <a
              target="_blank"
              href="https://github.com/JohnEarnest/ok"
              className="text-emerald-400 hover:underline"
            >
              the oK interpreter
            </a>
          </span>
        </div>

        {/* Formula Bar */}
        <div className="flex items-center gap-2">
          <div className="text-slate-400 text-sm font-mono w-32 flex items-center gap-2">
            <span>{selectedCell || ""}</span>
            <span
              className={`text-xs px-1.5 py-0.5 rounded ${mode === "normal" ? "bg-emerald-600 text-white" : mode === "visual" ? "bg-purple-600 text-white" : "bg-blue-600 text-white"}`}
            >
              {mode === "normal"
                ? "NORMAL"
                : mode === "visual"
                  ? "VISUAL"
                  : "INSERT"}
            </span>
          </div>
          <div className="flex-1 flex items-center gap-2">
            <input
              ref={inputRef}
              type="text"
              value={formulaBarValue}
              onChange={(e) => handleFormulaBarChange(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  handleFormulaBarSubmit();
                  setEditingCell(null);
                }
              }}
              onFocus={() => {
                // When formula bar is focused, don't use vim mode
                setMode("insert");
              }}
              onBlur={() => {
                // Return focus to container
                setTimeout(() => containerRef.current?.focus(), 0);
              }}
              placeholder="Enter formula (e.g., =!5 or =(</){(x*x;y*2)}/\(2;100)"
              className="flex-1 bg-slate-800 text-slate-100 px-3 py-2 rounded border border-slate-700 focus:outline-none focus:ring-2 focus:ring-emerald-500 font-mono text-sm"
            />
          </div>
        </div>
      </div>

      {/* Spreadsheet Grid */}
      <div className="flex-1 overflow-auto bg-slate-900">
        <div className="inline-block min-w-full">
          <table className="border-collapse">
            <thead>
              <tr>
                <th className="sticky top-0 left-0 z-20 bg-slate-800 border border-slate-700 w-12 h-8 text-slate-400 text-xs"></th>
                {columns.map((col) => (
                  <th
                    key={col}
                    className="sticky top-0 z-10 bg-slate-800 border border-slate-700 min-w-[120px] h-8 text-slate-400 text-xs font-mono font-normal"
                  >
                    {col}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {Array.from({ length: ROWS }, (_, rowIndex) => (
                <tr key={rowIndex}>
                  <td className="sticky left-0 z-10 bg-slate-800 border border-slate-700 w-12 h-8 text-center text-slate-400 text-xs font-mono px-2">
                    {rowIndex + 1}
                  </td>
                  {columns.map((col) => {
                    const cellId = `${col}${rowIndex + 1}`;
                    const cell = cells[cellId];
                    const isSelected = selectedCell === cellId;
                    const isEditing = editingCell === cellId;
                    const isInSelection = isCellInSelection(cellId);

                    return (
                      <Cell
                        key={cellId}
                        cell={cell}
                        isSelected={isSelected}
                        isEditing={isEditing}
                        isInSelection={isInSelection}
                        onClick={(e) => handleCellClick(cellId, e)}
                        onDoubleClick={() => handleCellDoubleClick(cellId)}
                        onEdit={(value) => handleCellEdit(cellId, value)}
                        onMouseDown={() => handleCellMouseDown(cellId)}
                        onMouseEnter={() => handleCellMouseEnter(cellId)}
                        setFormulaBarValue={setFormulaBarValue}
                      />
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Help Text */}
      <div className="bg-slate-900 border-t border-slate-700 px-4 py-2 text-xs text-slate-500">
        <span className="text-emerald-400">Vim Mode:</span> h/j/k/l to navigate
        • Shift+hjkl for VISUAL mode • i to edit • ESC to exit mode • d/x to
        delete • y to yank • p to paste • g/G for first/last row • 0/$ for
        first/last col
      </div>
    </div>
  );
}

interface CellProps {
  cell?: CellData;
  isSelected: boolean;
  isEditing: boolean;
  isInSelection: boolean;
  onClick: (e: React.MouseEvent) => void;
  onDoubleClick: () => void;
  onEdit: (value: string) => void;
  onMouseDown: () => void;
  onMouseEnter: () => void;
  setFormulaBarValue: (value: string) => void;
}

function Cell({
  cell,
  isSelected,
  isEditing,
  isInSelection,
  onClick,
  onDoubleClick,
  onEdit,
  onMouseDown,
  onMouseEnter,
  setFormulaBarValue,
}: CellProps) {
  const [editValue, setEditValue] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (isEditing) {
      setEditValue(cell?.formula || "");
      setTimeout(() => inputRef.current?.focus(), 0);
    }
  }, [isEditing, cell?.formula]);

  const displayValue = cell?.error
    ? `#ERROR: ${cell.error}`
    : cell?.value !== undefined && cell?.value !== null
      ? formatValue(cell.value)
      : "";

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter") {
      e.preventDefault();
      e.stopPropagation();
      onEdit(editValue);
    } else if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      onEdit(cell?.formula || "");
    }
  };

  return (
    <td
      onClick={onClick}
      onDoubleClick={onDoubleClick}
      onMouseDown={onMouseDown}
      onMouseEnter={onMouseEnter}
      className={`
        border border-slate-700 min-w-[120px] h-8 px-2 text-sm font-mono
        ${isSelected ? "max-w-96 text-wrap" : " max-w-64"}
        ${isSelected ? "ring-2 ring-emerald-500 ring-inset bg-slate-800" : isInSelection ? "bg-purple-900/40" : "bg-slate-900"}
        ${cell?.error ? "text-red-400" : "text-slate-200"}
        cursor-cell hover:bg-slate-800 transition-colors
      `}
    >
      {isEditing ? (
        <input
          ref={inputRef}
          type="text"
          value={editValue}
          onChange={(e) => {
            setEditValue(e.target.value);
            setFormulaBarValue(e.target.value);
          }}
          onKeyDown={handleKeyDown}
          onBlur={() => onEdit(editValue)}
          className="w-full h-full bg-transparent outline-none text-slate-100"
          onClick={(e) => e.stopPropagation()}
        />
      ) : (
        <div className={isSelected ? "" : "truncate"}>{displayValue}</div>
      )}
    </td>
  );
}

function formatValue(value: any): string {
  if (Array.isArray(value)) {
    return `[${value.map((v) => formatValue(v)).join(", ")}]`;
  }
  if (typeof value === "object" && value !== null) {
    return JSON.stringify(value);
  }
  if (typeof value === "string") {
    return value;
  }
  if (typeof value === "number") {
    return value.toString();
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  return String(value);
}

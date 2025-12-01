import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, AlertCircle } from 'lucide-react';

function App() {
  const [grandTotals, setGrandTotals] = useState([]);
  const [error, setError] = useState(null);
  const [fileName, setFileName] = useState(null);
  const [selectedCells, setSelectedCells] = useState(new Set());
  const [lastSelectedCell, setLastSelectedCell] = useState(null);
  const [isDragging, setIsDragging] = useState(false);

  const formatNumber = (value) => typeof value === 'number' && !isNaN(value) ? value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : value;

  const formatDepositNumber = (value) => {
    if (value === 'N/A') return 'N/A';
    if (typeof value === 'string' || typeof value === 'number') {
      return `DS-${value}`;
    }
    return value;
  };

  const getCellKey = (row, col) => `${row}-${col}`;

  const handleCellMouseDown = (row, col, e) => {
    if (e.button !== 0) return; // left click only
    setIsDragging(true);

    if (e.ctrlKey || e.metaKey) {
      // Ctrl+Click: toggle cell
      setSelectedCells((prev) => {
        const newSet = new Set(prev);
        const key = getCellKey(row, col);
        if (newSet.has(key)) {
          newSet.delete(key);
        } else {
          newSet.add(key);
        }
        return newSet;
      });
    } else if (e.shiftKey && lastSelectedCell) {
      // Shift+Click: range select
      const [lastRow, lastCol] = lastSelectedCell;
      const newSet = new Set();
      const minRow = Math.min(row, lastRow);
      const maxRow = Math.max(row, lastRow);
      const minCol = Math.min(col, lastCol);
      const maxCol = Math.max(col, lastCol);

      for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
          newSet.add(getCellKey(r, c));
        }
      }
      setSelectedCells(newSet);
    } else {
      // Single click: select only this cell
      setSelectedCells(new Set([getCellKey(row, col)]));
      setLastSelectedCell([row, col]);
    }
  };

  const handleCellMouseEnter = (row, col) => {
    if (!isDragging) return;
    // Drag to select
    if (lastSelectedCell) {
      const [startRow, startCol] = lastSelectedCell;
      const newSet = new Set();
      const minRow = Math.min(row, startRow);
      const maxRow = Math.max(row, startRow);
      const minCol = Math.min(col, startCol);
      const maxCol = Math.max(col, startCol);

      for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
          newSet.add(getCellKey(r, c));
        }
      }
      setSelectedCells(newSet);
    }
  };

  const handleMouseUp = () => {
    setIsDragging(false);
  };

  const isCellSelected = (row, col) => selectedCells.has(getCellKey(row, col));

  useEffect(() => {
    const handleKeyDown = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
        if (selectedCells.size === 0) return;

        // Build text from selected cells
        const cellArray = Array.from(selectedCells).map((key) => {
          const [row, col] = key.split('-').map(Number);
          const item = grandTotals[row];
          if (col === 0) return item.sheetName;
          if (col === 1) return formatDepositNumber(item.depositNumber);
          if (col === 2) return typeof item.grandTotal === 'number' ? formatNumber(item.grandTotal) : item.grandTotal;
          return '';
        });

        // Copy to clipboard
        navigator.clipboard.writeText(cellArray.join('\n'));

        // Deselect cells after copy
        setSelectedCells(new Set());
        setLastSelectedCell(null);

        e.preventDefault();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [selectedCells, grandTotals]);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) {
      setError('No file selected.');
      setFileName(null);
      return;
    }

    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        const totals = workbook.SheetNames.map((sheetName) => ({
          sheetName,
          date: workbook.Sheets[sheetName]['B3']?.v ?? 'N/A',
          depositNumber: workbook.Sheets[sheetName]['F8']?.v ?? 'N/A',
          grandTotal: workbook.Sheets[sheetName]['I54']?.v ?? 'N/A',
        }));
        setGrandTotals(totals);
        setError(null);
      } catch (error) {
        console.error(error);
        setError('Error reading file.');
        setFileName(null);
      }
    };
    reader.onerror = () => {
      setError('Error reading file.');
      setFileName(null);
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <div className="h-screen bg-gray-50 flex flex-col">
      <div className="flex-1 flex flex-col max-w-4xl mx-auto w-full p-4">
        {/* Header */}
        <div className="text-center mb-4">
          <div className="flex items-center justify-center mb-2">
            <FileSpreadsheet className="w-6 h-6 text-blue-600 mr-2" />
            <h1 className="text-2xl font-bold text-gray-900">LTO Grand Total</h1>
          </div>
          <p className="text-sm text-gray-600">Upload your Excel file to extract deposit numbers (F8) and grand totals (I54)</p>
        </div>

        {/* Upload Section */}
        <div className="bg-white rounded-lg shadow-sm border-2 border-dashed border-gray-300 hover:border-blue-400 transition-colors duration-200 mb-4">
          <label htmlFor="file-upload" className="cursor-pointer block p-4 text-center">
            <Upload className="w-8 h-8 text-gray-400 mx-auto mb-2" />
            <span className="text-base font-medium text-gray-700 block mb-1">
              Choose Excel File
            </span>
            <span className="text-xs text-gray-500">
              Supports .xlsx and .xls files
            </span>
            <input
              id="file-upload"
              type="file"
              accept=".xlsx, .xls"
              onChange={handleFileUpload}
              className="hidden"
            />
          </label>
        </div>

        {/* Error Message */}
        {error && (
          <div className="bg-red-50 border border-red-200 rounded-lg p-3 mb-4 flex items-center">
            <AlertCircle className="w-4 h-4 text-red-500 mr-2 flex-shrink-0" />
            <span className="text-sm text-red-800">{error}</span>
          </div>
        )}

        {/* Results Table */}
        {grandTotals.length > 0 && (
          <div className="bg-white rounded-lg shadow-sm overflow-hidden flex-1 flex flex-col min-h-0">
            <div className="px-4 py-3 bg-gray-50 border-b border-gray-200 flex-shrink-0">
              <div className="flex items-center justify-between">
                <h2 className="text-base font-semibold text-gray-900">
                  Results ({grandTotals.length} sheet{grandTotals.length !== 1 ? 's' : ''})
                </h2>
                {fileName && (
                  <span className="text-sm text-gray-600 truncate ml-4 max-w-xs">
                    📄 {fileName}
                  </span>
                )}
              </div>
            </div>
            <div className="flex-1 overflow-auto">
              <table className="w-full border-collapse select-none" onMouseUp={handleMouseUp} onMouseLeave={handleMouseUp}>
                <thead className="bg-gray-50 sticky top-0">
                  <tr>
                    <th className="border border-gray-300 px-4 py-2 text-left text-xs font-medium text-gray-700 bg-gray-100">
                      Sheet Name
                    </th>
                    <th className="border border-gray-300 px-4 py-2 text-left text-xs font-medium text-gray-700 bg-gray-100">
                      Deposit Number (F8)
                    </th>
                    <th className="border border-gray-300 px-4 py-2 text-right text-xs font-medium text-gray-700 bg-gray-100">
                      Total (I54)
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {grandTotals.map((item, rowIdx) => (
                    <tr key={rowIdx} className="hover:bg-gray-100">
                      <td
                        className={`border border-gray-300 px-4 py-2 text-sm text-gray-900 font-medium cursor-cell ${isCellSelected(rowIdx, 0) ? 'bg-blue-200' : 'bg-white'
                          }`}
                        onMouseDown={(e) => handleCellMouseDown(rowIdx, 0, e)}
                        onMouseEnter={() => handleCellMouseEnter(rowIdx, 0)}
                      >
                        {item.sheetName}
                      </td>
                      <td
                        className={`border border-gray-300 px-4 py-2 text-sm text-gray-900 font-mono cursor-cell ${isCellSelected(rowIdx, 1) ? 'bg-blue-200' : 'bg-white'
                          }`}
                        onMouseDown={(e) => handleCellMouseDown(rowIdx, 1, e)}
                        onMouseEnter={() => handleCellMouseEnter(rowIdx, 1)}
                      >
                        {formatDepositNumber(item.depositNumber)}
                      </td>
                      <td
                        className={`border border-gray-300 px-4 py-2 text-sm text-gray-900 text-right font-mono cursor-cell ${isCellSelected(rowIdx, 2) ? 'bg-blue-200' : 'bg-white'
                          }`}
                        onMouseDown={(e) => handleCellMouseDown(rowIdx, 2, e)}
                        onMouseEnter={() => handleCellMouseEnter(rowIdx, 2)}
                      >
                        {typeof item.grandTotal === 'number' ? formatNumber(item.grandTotal) : item.grandTotal}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;
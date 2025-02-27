import React, { useState, useRef } from 'react';
import './Spreadsheet.css';
import * as XLSX from 'xlsx';

const numRows = 10;
const initialNumCols = 5;

const Spreadsheet = () => {
  const [data, setData] = useState(
    Array.from({ length: numRows }, () => Array(initialNumCols).fill(''))
  );
  const [numCols, setNumCols] = useState(initialNumCols);
  const [cellStyles, setCellStyles] = useState({});
  const [selectedCell, setSelectedCell] = useState(null);
  const [formula, setFormula] = useState('');
  const [dragStart, setDragStart] = useState(null);
  const tableRef = useRef(null);

  const numberToColumnName = (num) => {
    let columnName = '';
    while (num >= 0) {
      columnName = String.fromCharCode((num % 26) + 65) + columnName;
      num = Math.floor(num / 26) - 1;
    }
    return columnName;
  };

  const columnNameToIndex = (colName) => {
    let index = 0;
    for (let i = 0; i < colName.length; i++) {
      index = index * 26 + (colName.charCodeAt(i) - 65);
    }
    return index;
  };

  const getCellValue = (col, row, currentData) => {
    const colIndex = columnNameToIndex(col);
    const rowIndex = parseInt(row, 10) - 1;
    return currentData[rowIndex]?.[colIndex] || '';
  };

  const applyStyle = (styleKey, value) => {
    if (selectedCell) {
      setCellStyles((prevStyles) => {
        const key = `${selectedCell.row}-${selectedCell.col}`;
        return {
          ...prevStyles,
          [key]: {
            ...prevStyles[key],
            [styleKey]: prevStyles[key]?.[styleKey] === value ? 'normal' : value,
          },
        };
      });
    }
  };

  const evaluateFormula = (formula, currentData) => {
    try {
      if (!formula.startsWith('=')) return formula;
      const formulaContent = formula.slice(1).trim();

      const getRangeValues = (range) => {
        const [start, end] = range.split(':');
        const startCol = start.match(/[A-Z]+/i)[0];
        const startRow = start.match(/\d+/)[0];
        const endCol = end.match(/[A-Z]+/i)[0];
        const endRow = end.match(/\d+/)[0];
        const colStartIndex = columnNameToIndex(startCol.toUpperCase());
        const colEndIndex = columnNameToIndex(endCol.toUpperCase());
        const rowStartIndex = parseInt(startRow, 10) - 1;
        const rowEndIndex = parseInt(endRow, 10) - 1;

        let values = [];
        for (let r = rowStartIndex; r <= rowEndIndex; r++) {
          for (let c = colStartIndex; c <= colEndIndex; c++) {
            let value = currentData[r][c];
            if (!isNaN(value)) {
              values.push(parseFloat(value));
            } else {
              values.push(value);
            }
          }
        }
        return values;
      };

      if (/^(SUM|AVERAGE|MAX|MIN|COUNT|MEDIAN|MODE)\((.*?)\)$/i.test(formulaContent)) {
        const [, func, range] = formulaContent.match(/^(SUM|AVERAGE|MAX|MIN|COUNT|MEDIAN|MODE)\((.*?)\)$/i);
        const values = getRangeValues(range).filter((val) => typeof val === 'number');

        switch (func.toUpperCase()) {
          case 'SUM':
            return values.reduce((acc, val) => acc + val, 0);
          case 'AVERAGE':
            return values.length ? values.reduce((acc, val) => acc + val, 0) / values.length : 0;
          case 'MAX':
            return Math.max(...values);
          case 'MIN':
            return Math.min(...values);
          case 'COUNT':
            return values.length;
          case 'MEDIAN':
            values.sort((a, b) => a - b);
            const mid = Math.floor(values.length / 2);
            return values.length % 2 !== 0 ? values[mid] : (values[mid - 1] + values[mid]) / 2;
          case 'MODE':
            const freq = {};
            values.forEach((val) => {
              freq[val] = (freq[val] || 0) + 1;
            });
            const maxFreq = Math.max(...Object.values(freq));
            return values.filter((val) => freq[val] === maxFreq)[0];
          default:
            return 'ERROR';
        }
      }

      if (/^(TRIM|UPPER|LOWER)\((.*?)\)$/i.test(formulaContent)) {
        const [, func, arg] = formulaContent.match(/^(TRIM|UPPER|LOWER)\((.*?)\)$/i);
        const match = arg.match(/([A-Z]+)(\d+)/i);
        if (!match) return 'ERROR';

        let cellValue = getCellValue(match[1].toUpperCase(), match[2], currentData);
        if (typeof cellValue !== 'string') cellValue = String(cellValue);

        switch (func.toUpperCase()) {
          case 'TRIM':
            return cellValue.trim();
          case 'UPPER':
            return cellValue.toUpperCase();
          case 'LOWER':
            return cellValue.toLowerCase();
          default:
            return 'ERROR';
        }
      }

      if (/^REMOVE_DUPLICATES\((.*?)\)$/i.test(formulaContent)) {
        const [, range] = formulaContent.match(/^REMOVE_DUPLICATES\((.*?)\)$/i);
        const [start, end] = range.split(':').map((cell) => cell.trim());

        const matchStart = start.match(/([A-Z]+)(\d+)/i);
        const matchEnd = end.match(/([A-Z]+)(\d+)/i);

        if (!matchStart || !matchEnd) return 'ERROR';

        const colStart = columnNameToIndex(matchStart[1].toUpperCase());
        const rowStart = parseInt(matchStart[2], 10) - 1;
        const colEnd = columnNameToIndex(matchEnd[1].toUpperCase());
        const rowEnd = parseInt(matchEnd[2], 10) - 1;

        let seen = new Set();
        let newData = currentData.map((row) => [...row]);

        for (let r = rowStart; r <= rowEnd; r++) {
          let rowData = '';
          for (let c = colStart; c <= colEnd; c++) {
            rowData += getCellValue(numberToColumnName(c), r + 1, currentData) + '|';
          }
          if (seen.has(rowData)) {
            for (let c = colStart; c <= colEnd; c++) {
              newData[r][c] = '';
            }
          } else {
            seen.add(rowData);
          }
        }

        setData(newData);
        return 'Duplicates Removed';
      }

      if (/^FIND_AND_REPLACE\((.*?)\)$/i.test(formulaContent)) {
        const [, args] = formulaContent.match(/^FIND_AND_REPLACE\((.*?)\)$/i);
        let [range, findText, replaceText] = args.split(',').map((arg) => arg.trim().replace(/"/g, ''));

        const [start, end] = range.split(':').map((cell) => cell.trim());
        const matchStart = start.match(/([A-Z]+)(\d+)/i);
        const matchEnd = end.match(/([A-Z]+)(\d+)/i);

        if (!matchStart || !matchEnd) return 'ERROR';

        const colStart = columnNameToIndex(matchStart[1].toUpperCase());
        const rowStart = parseInt(matchStart[2], 10) - 1;
        const colEnd = columnNameToIndex(matchEnd[1].toUpperCase());
        const rowEnd = parseInt(matchEnd[2], 10) - 1;

        let newData = currentData.map((row) => [...row]);

        for (let r = rowStart; r <= rowEnd; r++) {
          for (let c = colStart; c <= colEnd; c++) {
            let val = getCellValue(numberToColumnName(c), r + 1, currentData);
            if (val === findText) val = replaceText;
            newData[r][c] = val;
          }
        }

        setData(newData);
        return 'Find and Replace Completed';
      }

      // Handle basic arithmetic operations with support for relative and absolute references
      const arithmeticRegex = /^([A-Z]+\d+|\$[A-Z]+\d+|[A-Z]+\$?\d+|\$?[A-Z]+\$?\d+)\s*([+\-*/])\s*([A-Z]+\d+|\$[A-Z]+\d+|[A-Z]+\$?\d+|\$?[A-Z]+\$?\d+)$/i;
      if (arithmeticRegex.test(formulaContent)) {
        const [, cell1, operator, cell2] = formulaContent.match(arithmeticRegex);
        const value1 = parseFloat(getCellValue(cell1.match(/[A-Z]+/i)[0], cell1.match(/\d+/)[0], currentData));
        const value2 = parseFloat(getCellValue(cell2.match(/[A-Z]+/i)[0], cell2.match(/\d+/)[0], currentData));
        if (isNaN(value1) || isNaN(value2)) return 'ERROR';

        switch (operator) {
          case '+':
            return value1 + value2;
          case '-':
            return value1 - value2;
          case '*':
            return value1 * value2;
          case '/':
            return value2 !== 0 ? value1 / value2 : 'ERROR';
          default:
            return 'ERROR';
        }
      }

      return 'ERROR';
    } catch {
      return 'ERROR';
    }
  };

  const handleFormulaChange = (e) => {
    setFormula(e.target.value);
  };

  const handleFormulaKeyDown = (e) => {
    if (e.key === 'Enter' && selectedCell) {
      const { row, col } = selectedCell;
      const evaluatedFormula = evaluateFormula(formula, data);
      handleChange(row, col, evaluatedFormula);
    }
  };

  const handleFormulaDrag = (row, col, targetRow, targetCol) => {
    const originalFormula = data[row][col];

    if (typeof originalFormula !== 'string' || !originalFormula.startsWith('=')) {
      return;
    }

    const adjustedFormula = originalFormula.replace(/([A-Z]+)(\d+)/gi, (match, colName, rowNumber) => {
      const newCol = numberToColumnName(columnNameToIndex(colName.toUpperCase()) + (targetCol - col));
      const newRow = parseInt(rowNumber) + (targetRow - row);
      return `${newCol}${newRow}`;
    });

    setData((prevData) => {
      const newData = [...prevData];
      newData[targetRow][targetCol] = adjustedFormula;
      return newData;
    });
  };

  const handleDragFill = (startRow, startCol, endRow, endCol) => {
    const startValue = data[startRow][startCol];
    const isFormula = typeof startValue === 'string' && startValue.startsWith('=');
    const isNumber = !isNaN(parseFloat(startValue));

    if (isNumber) {
      const baseValue = parseFloat(startValue);
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          const newValue = baseValue + (row - startRow);
          handleChange(row, col, newValue.toString());
        }
      }
    } else if (isFormula) {
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          handleFormulaDrag(startRow, startCol, row, col);
        }
      }
    }
  };

  const handleChange = (row, col, value) => {
    const newData = [...data];
    newData[row][col] = value;
    setData(newData);
  };

  const handleBlur = (row, col) => {
    const newData = [...data];
    let cellValue = newData[row][col];

    if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
      const evaluatedValue = evaluateFormula(cellValue, newData);
      newData[row][col] = evaluatedValue.toString();
      setData(newData);
    }
  };

  const handleKeyDown = (e, row, col) => {
    if (e.key === 'Enter') {
      handleBlur(row, col);
      const nextRow = row + 1 < data.length ? row + 1 : row;
      document.getElementById(`cell-${nextRow}-${col}`)?.focus();
    }
  };

  const handleCellSelect = (row, col) => {
    setSelectedCell({ row, col });
    setFormula(data[row][col] || '');
  };

  const handleDragStart = (row, col) => {
    setDragStart({ row, col });
  };

  const handleDragEnd = (row, col) => {
    if (dragStart) {
      handleDragFill(dragStart.row, dragStart.col, row, col);
    }
    setDragStart(null);
  };

  const handleDragOver = (row, col) => {
    // No operation needed for drag over
  };

  const addRow = () => {
    setData([...data, Array(numCols).fill('')]);
  };

  const deleteRow = (rowIndex) => {
    if (rowIndex != null && rowIndex >= 0 && rowIndex < data.length) {
      const newData = data.filter((_, index) => index !== rowIndex);
      setData(newData);
    }
  };

  const addColumn = () => {
    const newData = data.map((row) => [...row, '']);
    setData(newData);
    setNumCols(numCols + 1);
  };

  const deleteColumn = (colIndex) => {
    const newData = data.map((row) => row.filter((_, index) => index !== colIndex));
    setData(newData);
    setNumCols(numCols - 1);
  };

  const saveSpreadsheet = () => {
    alert('Save functionality is not yet implemented.');
  };

  const loadSpreadsheet = (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
      const binaryStr = e.target.result;
      const workbook = XLSX.read(binaryStr, { type: 'binary' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const newData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      setData(newData);
    };
    reader.readAsBinaryString(file);
  };

  const createChart = () => {
    alert('Create Chart functionality is not yet implemented.');
  };

  return (
    <div className="spreadsheet-container">
      <div className="title-bar">
        <span role="img" aria-label="spreadsheet">
          📄
        </span>
        <input type="text" placeholder="Untitled Spreadsheet" />
      </div>
      <div className="toolbar">
        <button onClick={() => applyStyle('fontWeight', 'bold')}>
          <b>B</b>
        </button>
        <button onClick={() => applyStyle('fontStyle', 'italic')}>
          <i>I</i>
        </button>
        <select onChange={(e) => applyStyle('fontSize', e.target.value)}>
          {['12px', '14px', '16px', '18px', '20px', '22px'].map((size) => (
            <option key={size} value={size}>
              {size}
            </option>
          ))}
        </select>
        <input type="color" onChange={(e) => applyStyle('color', e.target.value)} />
        <button onClick={addRow}>Add Row</button>
        <button onClick={() => deleteRow(selectedCell?.row)}>Delete Row</button>
        <button onClick={addColumn}>Add Column</button>
        <button onClick={() => deleteColumn(selectedCell?.col)}>Delete Column</button>
        <button onClick={saveSpreadsheet}>Save</button>
        <input type="file" onChange={loadSpreadsheet} />
        <button onClick={createChart}>Create Chart</button>
      </div>
      <div className="formula-bar">
        <input
          type="text"
          value={formula}
          onChange={handleFormulaChange}
          onKeyDown={handleFormulaKeyDown}
          placeholder="Enter formula"
        />
      </div>
      <div className="spreadsheet-wrapper">
        <table ref={tableRef}>
          <thead>
            <tr>
              <th className="row-header"></th>
              {[...Array(numCols)].map((_, colIndex) => (
                <th key={colIndex} className="column-header">
                  {numberToColumnName(colIndex)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.map((row, rowIndex) => (
              <tr key={rowIndex}>
                <td className="row-header">{rowIndex + 1}</td>
                {row.map((cell, colIndex) => (
                  <td key={colIndex}>
                    <input
                      id={`cell-${rowIndex}-${colIndex}`}
                      type="text"
                      value={cell ?? ''}
                      onChange={(e) => handleChange(rowIndex, colIndex, e.target.value)}
                      onBlur={() => handleBlur(rowIndex, colIndex)}
                      onKeyDown={(e) => handleKeyDown(e, rowIndex, colIndex)}
                      onFocus={() => handleCellSelect(rowIndex, colIndex)}
                      onDragStart={() => handleDragStart(rowIndex, colIndex)}
                      onDragEnd={() => handleDragEnd(rowIndex, colIndex)}
                      onDragOver={() => handleDragOver(rowIndex, colIndex)}
                      draggable
                      style={cellStyles[`${rowIndex}-${colIndex}`] || {}}
                    />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default Spreadsheet;
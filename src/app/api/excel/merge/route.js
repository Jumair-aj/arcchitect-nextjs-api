import { NextRequest, NextResponse } from 'next/server';
import { join } from 'path';
import fs from 'fs/promises';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { writeFile } from 'fs/promises';
import { v4 as uuidv4 } from 'uuid';

const excelToJSON = (buffer, sheetIndex) => {
  const workbook = XLSX.read(buffer);
  const sheetName = workbook.SheetNames[sheetIndex];
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
};

const autoFitColumns = (worksheet) => {
  worksheet.columns.forEach(column => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const columnLength = cell.value ? cell.value.toString().length : 10;
      if (columnLength > maxLength) {
        maxLength = columnLength;
      }
    });
    column.width = maxLength < 10 ? 10 : maxLength;
  });
};

export async function POST(request) {
  try {
    const formData = await request.formData();
    const file1 = formData.get('file1') ;
    const file2 = formData.get('file2') ;
    const file3 = formData.get('file3') ;

    if (!file1 || !file2 || !file3) {
      return NextResponse.json({ error: 'All three files are required' }, { status: 400 });
    }

    const [excel1, excel2, excel3] = await Promise.all([
      excelToJSON(await file1.arrayBuffer(), 0),
      excelToJSON(await file2.arrayBuffer(), 0),
      excelToJSON(await file3.arrayBuffer(), 0),
    ]);

    // Your merging logic (same as in the original code)
    const data = excel1.map((drawing, i) => {
      const matchingTagWithMer = excel2.find((tag) => tag['MER Tag Number'] === drawing['EPE Tag Number'] && tag['Drawing Number'] === drawing['Drawing Number']);
      const matchingTagWithSap = excel3.find((tag) => tag['SAP TAG '] === drawing['EPE Tag Number']);
      return {
        "SL.NO": i + 1,
        "Drawing Number": drawing['Drawing Number'] || '',
        "EPE Tag Number": drawing['EPE Tag Number'] || '',
        "MER Tag No": matchingTagWithMer ? matchingTagWithMer['MER Tag Number'] : "NOT IN MER",
        "Site Markup Tag No": matchingTagWithMer ? matchingTagWithMer['site chainges'] || 'NO CHANGE IN SITE' : 'NO CHANGE IN SITE',
        "SAP tag": matchingTagWithSap ? (matchingTagWithMer ? "NO CHANGE IN SAP" : "AVAILABLE IN SAP") : "NOT IN SAP",
        "Equipment Description From SAP": matchingTagWithSap ? matchingTagWithSap['SAP DISCRIPTION'] || '' : '',
        "Equipment Type - New": matchingTagWithMer ? matchingTagWithMer['Equipment Type-New'] || '' : '',
        "Size - Old": matchingTagWithMer ? matchingTagWithMer['Size - Old'] || matchingTagWithMer['size new'] ? "" : "NOT AVAILABLE" : "NOT AVAILABLE",
        "Size - New": matchingTagWithMer ? matchingTagWithMer['size new'] || '' : '',
        "Drawing No.": drawing['Drawing Number'] || '',
        "Rev": '',
        "PCR / Project No.": matchingTagWithMer ? matchingTagWithMer['PCR/PROJECT'] || '' : '',
        "Additional Information": matchingTagWithMer ? matchingTagWithMer['REMARK'] || '' : '',
        "DRAWING LINK": '',
        "ECM LINK": '',
        "OAO LINK": '',
      };
    });
    excel2.forEach(tag => {
        if (!excel1.some(drawing => drawing['EPE Tag Number'] === tag['MER Tag Number'] && tag['Drawing Number'] === drawing['Drawing Number'])) {
            const matchingTagWithSap = excel3.find(sap => sap['SAP TAG '] === tag['MER Tag Number']);
            data.push({
                "SL.NO": data.length + 1,
                "Drawing Number": tag['Drawing Number'],
                "EPE Tag Number": "NOT IN EPE",
                "MER Tag No": tag['MER Tag Number'] || 'NOT IN MER',
                "Site Markup Tag No": tag['site chainges'] || 'NO CHANGE IN SITE',
                "SAP tag": matchingTagWithSap ? "AVAILABLE IN SAP" : "NOT IN SAP",
                "Equipment Description From SAP": matchingTagWithSap ? matchingTagWithSap['SAP DISCRIPTION'] || '' : '',
                "Equipment Type - New": tag ? tag['Equipment Type-New'] || '' : '',
                "Size - Old": tag ? tag['Size - Old'] ? tag['Size - Old'] : tag['size new'] ? "" : "NOT AVAILABLE" : "NOT AVAILABLE",
                "Size - New": tag ? tag['size new'] || '' : '',
                "Drawing No.": tag['Drawing Number'] || '',
                "Rev": '',
                "PCR / Project No.": tag ? tag['PCR/PROJECT'] || '' : '',
                "Additional Information": tag ? tag['REMARK'] || '' : '',
                "DRAWING LINK": '',
                "ECM LINK": '',
                "OAO LINK": '',
            });
        }
    });

    // ... (rest of the merging logic)

    // Create a new workbook and worksheet using ExcelJS
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Architect');
    worksheet.getRow(1).height = 30;

    // Define the columns for the worksheet
    worksheet.columns = [
      { header: 'SL.NO', key: 'SL.NO', width: 10 },
      { header: 'Drawing Number', key: 'Drawing Number', width: 20 },
      { header: 'EPE Tag Number', key: 'EPE Tag Number', width: 20 },
      { header: 'MER Tag No', key: 'MER Tag No', width: 20 },
      { header: 'Site Markup Tag No', key: 'Site Markup Tag No', width: 20 },
      { header: 'SAP tag', key: 'SAP tag', width: 20 },
      { header: 'Equipment Description From SAP', key: 'Equipment Description From SAP', width: 30 },
      { header: 'Equipment Type - New', key: 'Equipment Type - New', width: 20 },
      { header: 'Size - Old', key: 'Size - Old', width: 20 },
      { header: 'Size - New', key: 'Size - New', width: 20 },
      { header: 'Drawing No.', key: 'Drawing No.', width: 20 },
      { header: 'Rev', key: 'Rev', width: 10 },
      { header: 'PCR / Project No.', key: 'PCR / Project No.', width: 20 },
      { header: 'Additional Information', key: 'Additional Information', width: 30 },
      { header: 'DRAWING LINK', key: 'DRAWING LINK', width: 20 },
      { header: 'ECM LINK', key: 'ECM LINK', width: 20 },
      { header: 'OAO LINK', key: 'OAO LINK', width: 20 },
    ];

    // Add data rows to the worksheet
    worksheet.addRows(data);

    // Style the header row
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '4472C4' } };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

   // Set height for the header row
   worksheet.getRow(1).height = 30;

   // Adjust row height for all other rows
   worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
       if (rowNumber > 1) {
           row.height = 50;
       }
       row.eachCell((cell) => {
           cell.alignment = { vertical: 'middle', horizontal: 'center' };
           cell.border = {
               top: { style: 'thin' },
               left: { style: 'thin' },
               bottom: { style: 'thin' },
               right: { style: 'thin' }
           };
       });
   });

   // Auto-fit columns based on content
   worksheet.columns.forEach(column => {
       let maxLength = 0;
       column.eachCell({ includeEmpty: true }, (cell) => {
           const columnLength = cell.value ? cell.value.toString().length : 10;
           if (columnLength > maxLength) {
               maxLength = columnLength;
           }
       });
       column.width = maxLength + 2; // Add some padding
   });
   worksheet.eachRow((row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
});

    // Auto-fit columns
    autoFitColumns(worksheet);

    const sortAndMergeColumns = () => {
        const EPE_COL_INDEX = 3;
        const MER_COL_INDEX = 4;
        const SL_NO_COL_INDEX = 1;
        const DRAWING_NUMBER_COL_INDEX = 2; // Adjust this if the Drawing Number is in a different column
        const columnsToMerge = [1, 5, 6, 7, 9, 10]; // SL.NO, Site Markup Tag No, SAP tag, Equipment Description From SAP, Size - Old, Size - New

        // Function to safely get cell value
        const getCellValue = (row, col) => {
            try {
                return worksheet.getCell(row, col).value;
            } catch (error) {
                console.error(`Error getting cell value at row ${row}, column ${col}: ${error.message}`);
                return null;
            }
        };

        // Function to safely set cell value
        const setCellValue = (row, col, value) => {
            try {
                worksheet.getCell(row, col).value = value;
            } catch (error) {
                console.error(`Error setting cell value at row ${row}, column ${col}: ${error.message}`);
            }
        };

        // Step 1: Collect all row data and merge information
        const rowData = [];
        const mergeInfo = new Map();
        for (let i = 2; i <= worksheet.rowCount; i++) {
            const row = {};
            for (let j = 1; j <= worksheet.columnCount; j++) {
                const cell = worksheet.getCell(i, j);
                row[j] = cell.value;
                if (cell.isMerged) {
                    const master = cell.master;
                    if (master.address === cell.address) {
                        mergeInfo.set(`${i},${j}`, {
                            start: { row: master.row, col: master.col },
                            end: { row: master.row + master.rowSpan - 1, col: master.col + master.colSpan - 1 }
                        });
                    }
                }
            }
            rowData.push(row);
        }


         const getCombinedTag = (row) => {
             const epeTag = (row[EPE_COL_INDEX] || '').toString().trim();
             const merTag = (row[MER_COL_INDEX] || '').toString().trim();
 
             if (epeTag === "NOT IN EPE" && merTag !== "NOT IN MER") {
                 return merTag;
             } else if (merTag === "NOT IN MER" && epeTag !== "NOT IN EPE") {
                 return epeTag;
             } else if (epeTag !== "NOT IN EPE" && merTag !== "NOT IN MER") {
                 return epeTag || merTag; // Use EPE if available, otherwise MER
             } else {
                 return epeTag + merTag; // Both are "NOT IN", combine them
             }
         };
 
         // Step 3: Sort the row data based on the combined tag
         rowData.sort((a, b) => {
             const tagA = getCombinedTag(a);
             const tagB = getCombinedTag(b);
             return tagA.localeCompare(tagB);
         });
 
 
        // Step 3: Group rows by merged cells
        const groupedRows = [];
        let currentGroup = [];
        rowData.forEach((row, index) => {
            const isNotIn = row[EPE_COL_INDEX] === "NOT IN EPE" && row[MER_COL_INDEX] === "NOT IN MER";

            if (isNotIn) {
                // If the current row is a "NOT IN" row, add it as a single-row group
                if (currentGroup.length > 0) {
                    groupedRows.push(currentGroup);
                    currentGroup = [];
                }
                groupedRows.push([{ row, originalIndex: index + 2 }]);
            } else {
                currentGroup.push({ row, originalIndex: index + 2 });
                const nextRow = rowData[index + 1];
                if (!nextRow || (row[EPE_COL_INDEX] !== nextRow[EPE_COL_INDEX] && row[MER_COL_INDEX] !== nextRow[MER_COL_INDEX])) {
                    groupedRows.push(currentGroup);
                    currentGroup = [];
                }
            }
        });

        // Add any remaining rows in the current group
        if (currentGroup.length > 0) {
            groupedRows.push(currentGroup);
        }

        // Step 4: Sort groups by drawing number
        groupedRows.sort((a, b) => {
            const drawingA = a[0].row[DRAWING_NUMBER_COL_INDEX] || '';
            const drawingB = b[0].row[DRAWING_NUMBER_COL_INDEX] || '';
            return drawingA.localeCompare(drawingB);
        });

        // Step 5: Flatten groups and rewrite data to worksheet
        let newRowIndex = 2;
        groupedRows.forEach(group => {
            group.forEach(({ row, originalIndex }) => {
                for (let j = 1; j <= worksheet.columnCount; j++) {
                    setCellValue(newRowIndex, j, row[j]);
                }
                newRowIndex++;
            });
        });

        // Function to safely merge cells
        const safeMergeCells = (startRow, endRow, colIndex) => {
            if (startRow === endRow) return; // No need to merge a single cell
            try {
                worksheet.mergeCells(startRow, colIndex, endRow, colIndex);
            } catch (error) {
                // console.error(`Error merging cells from row ${startRow} to ${endRow}, column ${colIndex}: ${error.message}`);
            }
        };

        // Function to merge cells for a specific column and related columns
        const mergeCellsForColumn = (colIndex) => {
            let currentValue = null;
            let startRow = null;
            let endRow = null;
            let slNo = 1;

            for (let i = 2; i <= worksheet.rowCount; i++) {
                const cellValue = getCellValue(i, colIndex);
                const epeValue = getCellValue(i, EPE_COL_INDEX);

                // For MER column, ignore "NOT IN EPE" in EPE column
                const shouldMerge = colIndex === MER_COL_INDEX
                    ? cellValue === currentValue && cellValue && cellValue !== "NOT IN MER"
                    : cellValue === currentValue && cellValue && cellValue !== "NOT IN EPE" && cellValue !== "NOT IN MER";

                if (shouldMerge) {
                    endRow = i;
                } else {
                    if (startRow !== null && endRow !== null) {
                        if (startRow !== endRow) {
                            safeMergeCells(startRow, endRow, colIndex);

                            columnsToMerge.forEach(relatedColIndex => {
                                let shouldMerge = true;

                                // Skip content comparison for the first column (e.g., SL.NO column)
                                if (relatedColIndex !== SL_NO_COL_INDEX) {
                                    // Loop through each row in the range and check if the contents are the same
                                    for (let row = startRow; row <= endRow; row++) {
                                        const currentCellValue = worksheet[XLSX.utils.encode_cell({ r: row, c: relatedColIndex })]?.v;
                                        const firstCellValue = worksheet[XLSX.utils.encode_cell({ r: startRow, c: relatedColIndex })]?.v;

                                        // If any cell in the column has a different value, skip merging for this column
                                        if (currentCellValue !== firstCellValue) {
                                            shouldMerge = false;
                                            break;
                                        }
                                    }
                                }

                                // Don't merge if the contents are different, but always merge SL.NO column
                                if (shouldMerge || relatedColIndex === SL_NO_COL_INDEX) {
                                    // For MER column, don't merge EPE column if it contains "NOT IN EPE"
                                    if (!(colIndex === MER_COL_INDEX && relatedColIndex === EPE_COL_INDEX && epeValue === "NOT IN EPE")) {
                                        safeMergeCells(startRow, endRow, relatedColIndex);
                                    }
                                }
                            });


                            // Set SL.NO for the merged group
                            setCellValue(startRow, SL_NO_COL_INDEX, slNo++);
                        } else {
                            // Set SL.NO for single row
                            setCellValue(startRow, SL_NO_COL_INDEX, slNo++);
                        }
                    }

                    currentValue = cellValue;
                    startRow = i;
                    endRow = i;
                }
            }

            // Handle the last group
            if (startRow !== null && endRow !== null) {
                if (startRow !== endRow) {
                    safeMergeCells(startRow, endRow, colIndex);

                    // Merge related columns for the last range including SL.NO
                    columnsToMerge.forEach(relatedColIndex => {
                        let shouldMerge = true;

                        // Skip content comparison for the first column (e.g., SL.NO column)
                        if (relatedColIndex !== SL_NO_COL_INDEX) {
                            // Loop through each row in the range and check if the contents are the same
                            for (let row = startRow; row <= endRow; row++) {
                                const currentCellValue = worksheet[XLSX.utils.encode_cell({ r: row, c: relatedColIndex })]?.v;
                                const firstCellValue = worksheet[XLSX.utils.encode_cell({ r: startRow, c: relatedColIndex })]?.v;

                                // If any cell in the column has a different value, skip merging for this column
                                if (currentCellValue !== firstCellValue) {
                                    shouldMerge = false;
                                    break;
                                }
                            }
                        }

                        // Don't merge if the contents are different, but always merge SL.NO column
                        if (shouldMerge || relatedColIndex === SL_NO_COL_INDEX) {
                            // For MER column, don't merge EPE column if it contains "NOT IN EPE"
                            if (!(colIndex === MER_COL_INDEX && relatedColIndex === EPE_COL_INDEX && epeValue === "NOT IN EPE")) {
                                safeMergeCells(startRow, endRow, relatedColIndex);
                            }
                        }
                    });


                    // Set SL.NO for the last merged group
                    setCellValue(startRow, SL_NO_COL_INDEX, slNo++);
                } else {
                    // Set SL.NO for the last single row
                    setCellValue(startRow, SL_NO_COL_INDEX, slNo++);
                }
            }
        };

        // Step 5: Merge cells for EPE column
        mergeCellsForColumn(EPE_COL_INDEX);

        // Step 6: Merge cells for MER column
        mergeCellsForColumn(MER_COL_INDEX);



        // Step 7: Update SL.NO
        let currentSlNo = 1;
        for (let i = 2; i <= worksheet.rowCount; i++) {
            const cell = worksheet.getCell(i, SL_NO_COL_INDEX);
            if (cell._mergeCount) {
                i = i + cell._mergeCount - 1
            }
            else {
                setCellValue(i, SL_NO_COL_INDEX, currentSlNo++);
            }

        }
    };

    // Apply the sorting and merging logic
    sortAndMergeColumns();

    // Generate a unique filename
    const filename = `merged_output_${uuidv4()}.xlsx`;
    const outputFilePath = join(process.cwd(), 'tmp', filename);

    // Ensure the tmp directory exists
    await fs.mkdir(join(process.cwd(), 'tmp'), { recursive: true });

    // Write to file
    await workbook.xlsx.writeFile(outputFilePath);

    // Read the file
    const fileBuffer = await fs.readFile(outputFilePath);

    // Delete the temporary file
    await fs.unlink(outputFilePath);

    // Set response headers
    const headers = new Headers();
    headers.set('Content-Disposition', `attachment; filename=${filename}`);
    headers.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    headers.set('Access-Control-Allow-Origin', '*');
    headers.set('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    headers.set('Access-Control-Allow-Headers', 'Content-Type, Authorization');

    // Return the file as a response
    return new NextResponse(fileBuffer, {
      status: 200,
      headers: headers,
    });
  } catch (error) {
    console.error('Error processing request:', error);
    return NextResponse.json({ error: 'Internal Server Error' }, { status: 500 });
  }
}

export async function OPTIONS() {
  const headers = new Headers();
  headers.set('Access-Control-Allow-Origin', '*');
  headers.set('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  headers.set('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  return new NextResponse(null, { status: 204, headers: headers });
}

export const maxDuration = 5

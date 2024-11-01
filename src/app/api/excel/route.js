import { NextResponse } from 'next/server';
import { join } from 'path';
import fs from 'fs/promises';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
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
    column.width = maxLength < 10 ? 10 : maxLength + 2;
  });
};

export async function POST(request) {
  try {
    const formData = await request.formData();
    const file1 = formData.get('file1');
    const file2Files = formData.getAll('file2');
    const file3 = formData.get('file3');

    if (!file1 || file2Files.length === 0 || !file3) {
      return NextResponse.json({ error: 'All three files are required' }, { status: 400 });
    }

    const excel1 = excelToJSON(await file1.arrayBuffer(), 1);
    const excel3 = excelToJSON(await file3.arrayBuffer(), 0);

    let excel2 = [];
    for (const file of file2Files) {
      const fileData = excelToJSON(await file.arrayBuffer(), 0);
      const fileDataWithFilename = fileData.map(row => ({
        ...row,
        filename: file.name.split('.')[0].split('-').join('')
      }));
      excel2 = excel2.concat(fileDataWithFilename);
    }

    const Dname = file2Files.map(file => file.name.split('.')[0].split('-').join(''));

    const data = excel1.map((drawing, i) => {
      const matchingTagWithMer = excel2.find(tag => tag['MER TAG NO'] === drawing['Tag Number']);
      const matchingTagWithSap = excel3.find(tag => tag['SAP-TAGS'] === drawing['Tag Number']);
      let dngNo = '';
      if (drawing['Drawing no'].split('-').length > 3 || drawing['Drawing no'].length != 14) {
        dngNo = drawing['Drawing no'];
      } else {
        dngNo = drawing['Drawing no'].slice(0, 4) + '-' + drawing['Drawing no'].slice(4, 5) + '-' + drawing['Drawing no'].slice(5, 7) + '-' + drawing['Drawing no'].slice(7, 11) + '-' + drawing['Drawing no'].slice(11);
      }
      return {
        "SL.NO": i + 1,
        "Drawing Number": dngNo || '',
        "EPE Tag Number": drawing['Tag Number'] || '',
        "MER Tag No": matchingTagWithMer ? matchingTagWithMer['MER TAG NO'] : Dname.includes(drawing['Drawing no']) ? "NOT IN MER" : '',
        "Site Markup Tag No": matchingTagWithMer ? matchingTagWithMer['SITE CHANGE'] : '',
        "Final MER Tag No": matchingTagWithMer ? matchingTagWithMer['SITE CHANGE'] && matchingTagWithMer['SITE CHANGE'].includes("NO") ? matchingTagWithMer['MER TAG NO'] : matchingTagWithMer['SITE CHANGE'] : '',
        "SAP tag": matchingTagWithSap ? (matchingTagWithMer ? "NO CHANGE IN SAP" : "AVAILABLE IN SAP") : "NOT IN SAP",
        "MERRemarks": '',
        "Equipment Description From SAP": matchingTagWithSap ? matchingTagWithSap['DESCRIPTION'] || '' : '',
        "Equipment Type - New": '',
        "Size - Old": matchingTagWithMer ? matchingTagWithMer['Size - Old'] ? matchingTagWithMer['Size - Old'] : matchingTagWithMer['Size - New'] ? "" : "NOT AVAILABLE" : 'NOT AVAILABLE',
        "Size From SAP": '',
        "Size - New": matchingTagWithMer ? matchingTagWithMer['Size - New'] : "",
        "Drawing No.": drawing['Drawing no'] || '',
        "Rev": '',
        "PCR / Project No.": matchingTagWithMer ? matchingTagWithMer['PCR / Project No.'] : '',
        "Additional Information": '',
        "DRAWING LINK": "",
        "ECM LINK": '',
        "OAO LINK": '',
      };
    });

    excel2.forEach(tag => {
      if (!excel1.some(drawing => drawing['Tag Number'] === tag['MER TAG NO'])) {
        let dngNo = '';
        if (tag.filename.split('-').length > 3 || tag.filename.length != 14) {
          dngNo = tag.filename;
        } else {
          dngNo = tag.filename.slice(0, 4) + '-' + tag.filename.slice(4, 5) + '-' + tag.filename.slice(5, 7) + '-' + tag.filename.slice(7, 11) + '-' + tag.filename.slice(11);
        }
        data.push({
          "SL.NO": data.length + 1,
          "Drawing Number": dngNo,
          "EPE Tag Number": "NOT IN EPE",
          "MER Tag No": tag['MER TAG NO'] || '',
          "Site Markup Tag No": tag ? tag['SITE CHANGE'] : '',
          "Final MER Tag No": tag ? tag['SITE CHANGE'] && tag['SITE CHANGE'].includes("NO") ? tag['MER TAG NO'] : tag['SITE CHANGE'] : '',
          "SAP tag": 'NOT IN SAP',
          "MERRemarks": '',
          "Equipment Description From SAP": "",
          "Equipment Type - New": '',
          "Size - Old": tag ? tag['Size - Old'] || tag['Size - New'] ? "" : "NOT AVAILABLE" : '',
          "Size From SAP": '',
          "Size - New": tag ? tag['Size - New'] : "",
          "Drawing No.": tag.filename,
          "Rev": '',
          "PCR / Project No.": tag ? tag['PCR / Project No.'] : '',
          "Additional Information": '',
          "DRAWING LINK": "",
          "ECM LINK": '',
          "OAO LINK": '',
        });
      }
    });

    data.sort((a, b) => (a['Drawing Number'] || '').localeCompare(b['Drawing Number'] || ''));
    data.forEach((item, index) => { item["SL.NO"] = index + 1; });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Architect');
    worksheet.getRow(1).height = 50;

    worksheet.columns = [
      { header: 'SL.NO', key: 'SL.NO', width: 10 },
      { header: 'Drawing Number', key: 'Drawing Number', width: 20 },
      { header: 'EPE Tag Number', key: 'EPE Tag Number', width: 20, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC000' } } },
      { header: 'MER Tag No', key: 'MER Tag No', width: 20 },
      { header: 'Site Markup Tag No', key: 'Site Markup Tag No', width: 20 },
      { header: 'Final MER Tag No', key: 'Final MER Tag No', width: 20 },
      { header: 'SAP tag', key: 'SAP tag', width: 20 },
      { header: 'Equipment Description From SAP', key: 'Equipment Description From SAP', width: 30 },
      { header: 'Equipment Type - New', key: 'Equipment Type - New', width: 20 },
      { header: 'Size - Old', key: 'Size - Old', width: 20, },
      { header: 'Size From SAP', key: 'Size From SAP', width: 20 },
      { header: 'Size - New', key: 'Size - New', width: 20 },
      { header: 'Drawing No.', key: 'Drawing No.', width: 20 },
      { header: 'Rev', key: 'Rev', width: 10 },
      { header: 'PCR / Project No.', key: 'PCR / Project No.', width: 20 },
      { header: 'Additional Information', key: 'Additional Information', width: 30 },
      { header: 'DRAWING LINK', key: 'DRAWING LINK', width: 20 },
      { header: 'ECM LINK', key: 'ECM LINK', width: 20 },
      { header: 'OAO LINK', key: 'OAO LINK', width: 20 },
    ];
    worksheet.addRows(data);

    const headerRow = worksheet.getRow(1);
    const rows = [3, 5, 8, 11, 17, 18, 19];
    headerRow.font = { bold: true, color: { argb: '000000' }, size: 11 };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C0C0C0' } };
    rows.forEach((row) => headerRow.getCell(row).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC000' } });
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    autoFitColumns(worksheet);

    const buffer = await workbook.xlsx.writeBuffer();

    const headers = new Headers();
    headers.set('Content-Disposition', 'attachment; filename="final-excel.xlsx"');
    headers.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    headers.set('Access-Control-Allow-Origin', '*');
    headers.set('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    headers.set('Access-Control-Allow-Headers', 'Content-Type, Authorization');

    return new NextResponse(buffer, {
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
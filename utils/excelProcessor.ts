import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { Box1Data, ProcessedRow } from '../types';

const normalize = (s: any): string => {
  if (s === null || s === undefined) return "";
  return String(s).replace(/[^A-Za-z0-9]/g, '').toUpperCase();
};

export const processFiles = async (
  fileReference: File,
  fileOe: File
): Promise<ProcessedRow[]> => {
  // 1. 使用 ExcelJS 处理参考数据库以提取图片
  const refBuffer = await fileReference.arrayBuffer();
  const refWorkbook = new ExcelJS.Workbook();
  await refWorkbook.xlsx.load(refBuffer);
  const refWorksheet = refWorkbook.worksheets[0];

  // 寻找表头行 (通常是第1或第2行，含 'OEM')
  let headerRowIndex = 1;
  for (let i = 1; i <= 10; i++) {
    const row = refWorksheet.getRow(i);
    if (row.values.includes('OEM')) {
      headerRowIndex = i;
      break;
    }
  }

  const headers = refWorksheet.getRow(headerRowIndex).values as any[];
  const colMap: Record<string, number> = {};
  let priceColIndex = -1;
  let pictureColIndex = -1;

  headers.forEach((h, i) => {
    if (h) {
      const headerName = String(h).trim();
      colMap[headerName] = i;
      if (priceColIndex === -1 && (headerName.includes('广州') || headerName.includes('Price') || headerName.includes('价格'))) {
        priceColIndex = i;
      }
      if (pictureColIndex === -1 && (headerName.includes('图片') || headerName.includes('Picture') || headerName.includes('Image'))) {
        pictureColIndex = i;
      }
    }
  });

  // 提取图片映射
  const imageMap: Record<number, { buffer: ArrayBuffer; extension: string }> = {};
  refWorksheet.getImages().forEach((image) => {
    const img = refWorkbook.model.media.find((m: any, idx: number) => idx === (image as any).imageId || m.index === (image as any).imageId);
    if (img && image.range.tl.nativeRow + 1) {
      // 记录图片所在的行（注意 ExcelJS 索引从0开始，nativeRow+1 对应行号）
      imageMap[image.range.tl.nativeRow + 1] = {
        buffer: img.buffer,
        extension: img.extension
      };
    }
  });

  const mapRef: Record<string, Box1Data> = {};

  refWorksheet.eachRow((row, rowNumber) => {
    if (rowNumber <= headerRowIndex) return;

    const oemRaw = row.getCell(colMap['OEM']).value;
    if (!oemRaw) return;

    const tokens = String(oemRaw).split(/[\s\n]+/);
    const rowPrice = priceColIndex !== -1 ? row.getCell(priceColIndex).value : null;
    const imageData = imageMap[rowNumber] || null;

    for (const token of tokens) {
      const norm = normalize(token);
      if (norm) {
        mapRef[norm] = {
          xxCode: String(row.getCell(colMap['XX CODE']).value || ""),
          application: String(row.getCell(colMap['Application']).value || ""),
          year: String(row.getCell(colMap['Year']).value || ""),
          oem: String(oemRaw),
          drive: String(row.getCell(colMap['Drive']).value || ""),
          picture: "已提取图片",
          price: rowPrice as any,
          imageData: imageData
        };
      }
    }
  });

  // 2. 处理 OE 输入列表 (由于输入列表通常较简单，继续使用 XLSX 快速读取)
  const bufOe = await fileOe.arrayBuffer();
  const wbOe = XLSX.read(bufOe, { type: 'array' });
  const wsOe = wbOe.Sheets[wbOe.SheetNames[0]];
  const dataOe = XLSX.utils.sheet_to_json<any>(wsOe);

  const results: ProcessedRow[] = [];

  for (const row of dataOe) {
    const keys = Object.keys(row);
    const inputKey = keys.find(k => k.toUpperCase() === 'OE') || keys[0];
    const inputOE = row[inputKey];
    
    const normInput = normalize(inputOE);
    const match = mapRef[normInput];

    const newRow: ProcessedRow = {
      '输入 OE': inputOE,
      'XX 编码': null,
      '适用车型': null,
      '年份': null,
      'OEM': null,
      '驱动': null,
      '图片': null,
      '图片数据': null,
      '广州价': null,
    };

    if (match) {
      newRow['XX 编码'] = match.xxCode;
      newRow['适用车型'] = match.application;
      newRow['年份'] = match.year;
      newRow['OEM'] = match.oem;
      newRow['驱动'] = match.drive;
      newRow['图片'] = match.imageData ? "匹配成功" : "无图片";
      newRow['图片数据'] = match.imageData;
      newRow['广州价'] = match.price;
    }

    results.push(newRow);
  }

  return results;
};

export const exportToExcel = async (data: ProcessedRow[], fileName: string) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('匹配结果');

  const columns = ['输入 OE', 'XX 编码', '适用车型', '年份', 'OEM', '驱动', '图片', '广州价'];
  
  // 设置列定义
  worksheet.columns = columns.map(c => ({ 
    header: c, 
    key: c, 
    width: c === 'OEM' || c === '适用车型' ? 35 : (c === '图片' ? 20 : 15)
  }));

  // 设置表头样式
  worksheet.getRow(1).font = { bold: true };
  worksheet.getRow(1).height = 25;
  worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
  worksheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFF1F5F9' }
  };

  for (let i = 0; i < data.length; i++) {
    const rowData = data[i];
    const excelRow = worksheet.addRow({});
    const rowNumber = i + 2; // +1 表头，+1 索引
    
    // 设置行高以便显示图片
    excelRow.height = 80;
    excelRow.alignment = { vertical: 'middle', horizontal: 'center' };

    columns.forEach((col, colIdx) => {
      const cell = excelRow.getCell(colIdx + 1);
      const val = rowData[col];

      if (col === '图片' && rowData['图片数据']) {
        const imgData = rowData['图片数据'];
        const imageId = workbook.addImage({
          buffer: imgData.buffer,
          extension: imgData.extension as any,
        });

        worksheet.addImage(imageId, {
          tl: { col: colIdx, row: rowNumber - 1 },
          ext: { width: 100, height: 100 },
          editAs: 'oneCell'
        });
        cell.value = ""; // 清空文字，只显示图片
      } else if (col === 'OEM' && val && rowData['输入 OE']) {
        const oemStr = String(val);
        const inputNorm = normalize(rowData['输入 OE']);
        const parts = oemStr.split(/([\s\n]+)/);
        const richText: any[] = [];

        parts.forEach(part => {
           if (!part) return;
           if (/^[\s\n]+$/.test(part)) {
             richText.push({ text: part });
           } else {
             if (normalize(part) === inputNorm) {
               richText.push({ 
                 text: part, 
                 font: { color: { argb: 'FFFF0000' }, bold: true } 
               });
             } else {
               richText.push({ text: part });
             }
           }
        });
        cell.value = { richText };
      } else {
        cell.value = val;
      }
    });
  }

  // 写入并下载
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  window.URL.revokeObjectURL(url);
};
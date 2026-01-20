import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { Box1Data, ProcessedRow } from '../types';

const normalize = (s: any): string => {
  if (s === null || s === undefined) return "";
  const str = typeof s === 'object' ? (s.text || s.result || String(s)) : String(s);
  return str.replace(/[^A-Za-z0-9]/g, '').toUpperCase();
};

/**
 * 安全提取 ExcelJS 单元格的值
 * 处理公式、超链接、富文本等复杂情况
 */
const getCellValue = (cell: ExcelJS.Cell): any => {
  const val = cell.value;
  if (val === null || val === undefined) return "";
  
  // 处理公式单元格
  if (typeof val === 'object' && 'result' in val) {
    return val.result ?? "";
  }
  
  // 处理超链接或富文本
  if (typeof val === 'object' && 'text' in val) {
    return val.text ?? "";
  }

  // 处理富文本数组
  if (typeof val === 'object' && 'richText' in val) {
    return (val as any).richText.map((t: any) => t.text).join("");
  }

  return val;
};

// 辅助函数：根据可能的名称寻找列索引
const findColIndex = (headers: any[], possibleNames: string[]): number => {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "").trim().toUpperCase();
    if (possibleNames.some(name => h.includes(name.toUpperCase()))) {
      return i;
    }
  }
  return -1;
};

export const processFiles = async (
  fileReference: File,
  fileOe: File
): Promise<ProcessedRow[]> => {
  const refBuffer = await fileReference.arrayBuffer();
  const refWorkbook = new ExcelJS.Workbook();
  await refWorkbook.xlsx.load(refBuffer);
  const refWorksheet = refWorkbook.worksheets[0];

  let headerRowIndex = 1;
  let oemColIdx = -1;
  
  for (let i = 1; i <= 20; i++) {
    const rowValues = refWorksheet.getRow(i).values as any[];
    oemColIdx = findColIndex(rowValues, ['OEM', 'OE', '原厂编号', '零件号']);
    if (oemColIdx !== -1) {
      headerRowIndex = i;
      break;
    }
  }

  if (oemColIdx === -1) throw new Error("参考数据库中未找到包含 'OEM' 的列。");

  const headers = refWorksheet.getRow(headerRowIndex).values as any[];
  
  const xxColIdx = findColIndex(headers, ['XX CODE', 'XX编码', '公司编号']);
  const appColIdx = findColIndex(headers, ['Application', '适用车型', '车型']);
  const yearColIdx = findColIndex(headers, ['Year', '年份', '年度']);
  const driveColIdx = findColIndex(headers, ['Drive', '驱动', '左/右']);
  const priceColIdx = findColIndex(headers, ['广州', 'Price', '价格', '单价']);

  const imageMap: Record<number, { buffer: ArrayBuffer; extension: string }> = {};
  refWorksheet.getImages().forEach((image) => {
    const img = refWorkbook.model.media.find((m: any, idx: number) => idx === (image as any).imageId || m.index === (image as any).imageId);
    if (img && image.range.tl.nativeRow + 1) {
      imageMap[image.range.tl.nativeRow + 1] = {
        buffer: img.buffer,
        extension: img.extension
      };
    }
  });

  const mapRef: Record<string, Box1Data> = {};

  refWorksheet.eachRow((row, rowNumber) => {
    if (rowNumber <= headerRowIndex) return;

    const oemRaw = getCellValue(row.getCell(oemColIdx));
    if (!oemRaw) return;

    const tokens = String(oemRaw).split(/[\s\n,;:/|，；、]+/);
    const imageData = imageMap[rowNumber] || null;

    for (const token of tokens) {
      const norm = normalize(token);
      if (norm && norm.length > 2) {
        mapRef[norm] = {
          xxCode: xxColIdx !== -1 ? String(getCellValue(row.getCell(xxColIdx)) || "") : "",
          application: appColIdx !== -1 ? String(getCellValue(row.getCell(appColIdx)) || "") : "",
          year: yearColIdx !== -1 ? String(getCellValue(row.getCell(yearColIdx)) || "") : "",
          oem: String(oemRaw),
          drive: driveColIdx !== -1 ? String(getCellValue(row.getCell(driveColIdx)) || "") : "",
          picture: "已提取",
          price: priceColIdx !== -1 ? getCellValue(row.getCell(priceColIdx)) : null,
          imageData: imageData
        };
      }
    }
  });

  const bufOe = await fileOe.arrayBuffer();
  const wbOe = XLSX.read(bufOe, { type: 'array' });
  const wsOe = wbOe.Sheets[wbOe.SheetNames[0]];
  const dataOeRaw = XLSX.utils.sheet_to_json<any[]>(wsOe, { header: 1 });
  
  let oeInputCol = 0;
  if (dataOeRaw.length > 0) {
    const firstRow = dataOeRaw[0];
    const detectedIdx = findColIndex(firstRow, ['OE', 'OEM', '查询', '输入']);
    if (detectedIdx !== -1) oeInputCol = detectedIdx - 1; 
    if (oeInputCol < 0) oeInputCol = 0;
  }

  const results: ProcessedRow[] = [];
  const isFirstRowHeader = findColIndex(dataOeRaw[0], ['OE', 'OEM', '零件']) !== -1;
  const startIdx = isFirstRowHeader ? 1 : 0;

  for (let i = startIdx; i < dataOeRaw.length; i++) {
    const row = dataOeRaw[i];
    if (!row || row.length === 0) continue;

    const inputOE = String(row[oeInputCol] || "").trim();
    if (!inputOE) continue;
    
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

  // 尺寸与居中计算 (96 DPI):
  // 1 磅 (Point) = 1/72 英寸
  // 2.15cm = 81 像素 (图片高度)
  // 2.15cm = 61 磅 (Excel 行高)
  // 6.0cm = 227 像素 (图片宽度)
  // 列宽 34 约为 251 像素
  
  // 水平偏移计算: (单元格宽 251px - 图片宽 227px) / 2 = 12px
  // 1 像素 = 9525 EMU (English Metric Units)
  const horizOffsetEMU = 12 * 9525; 
  // 垂直偏移计算: 行高 61pt 刚好对应 81px 图片高度，偏移设为 0 或微调
  const vertOffsetEMU = 0;

  const columns = ['输入 OE', 'XX 编码', '适用车型', '年份', 'OEM', '驱动', '图片', '广州价'];
  worksheet.columns = columns.map(c => ({ 
    header: c, 
    key: c, 
    width: c === 'OEM' || c === '适用车型' ? 35 : (c === '图片' ? 34 : 15)
  }));

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
    const rowNumber = i + 2;
    
    // 设置行高为 61 磅 (2.15cm)
    excelRow.height = 61;
    excelRow.alignment = { vertical: 'middle', horizontal: 'center' };

    columns.forEach((col, colIdx) => {
      const cell = excelRow.getCell(colIdx + 1);
      const val = rowData[col];

      if (col === '图片' && rowData['图片数据']) {
        const imgData = rowData['图片数据'];
        try {
          const imageId = workbook.addImage({
            buffer: imgData.buffer,
            extension: imgData.extension as any,
          });
          
          // 设置图片居中插入
          worksheet.addImage(imageId, {
            tl: { 
              col: colIdx, 
              row: rowNumber - 1,
              nativeColOff: horizOffsetEMU,
              nativeRowOff: vertOffsetEMU
            },
            ext: { width: 227, height: 81 },
            editAs: 'oneCell'
          });
        } catch (e) { console.error(e); }
        cell.value = "";
      } else if (col === 'OEM') {
        // 设置 OEM 列自动换行
        cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
        
        if (val && rowData['输入 OE']) {
          const oemStr = String(val);
          const inputNorm = normalize(rowData['输入 OE']);
          const parts = oemStr.split(/([\s\n,;:/|，；、]+)/);
          const richText: any[] = [];
          parts.forEach(part => {
             if (!part) return;
             if (/^[\s\n,;:/|，；、]+$/.test(part)) {
               richText.push({ text: part });
             } else {
               if (normalize(part) === inputNorm) {
                 richText.push({ text: part, font: { color: { argb: 'FFFF0000' }, bold: true } });
               } else {
                 richText.push({ text: part });
               }
             }
          });
          cell.value = { richText };
        } else {
          cell.value = val;
        }
      } else {
        cell.value = val;
      }
    });
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  window.URL.revokeObjectURL(url);
};
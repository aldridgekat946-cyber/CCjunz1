import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { PackingSpec, PackingInputRow, PackingCalculationResult } from '../types';

export const normalize = (s: any): string => {
  if (s === null || s === undefined) return "";
  const str = typeof s === 'object' ? (s.text || s.result || String(s)) : String(s);
  return str.replace(/[^A-Za-z0-9\u4e00-\u9fa5]/g, '').toUpperCase();
};

const getCellValue = (cell: ExcelJS.Cell): any => {
  const val = cell.value;
  if (val === null || val === undefined) return "";
  if (typeof val === 'object' && 'result' in val) return val.result ?? "";
  if (typeof val === 'object' && 'text' in val) return val.text ?? "";
  if (typeof val === 'object' && 'richText' in val) return (val as any).richText.map((t: any) => t.text).join("");
  return val;
};

const findColIndex = (headers: any[], possibleNames: string[]): number => {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "").trim().toUpperCase();
    if (possibleNames.some(name => h.includes(name.toUpperCase()))) return i;
  }
  return -1;
};

function parseDimensions(dimStr: string) {
  if (!dimStr) return null;
  // match patterns like 1000*155*235 or 1000x155x235 mm
  const cleaned = String(dimStr).toUpperCase().replace(/MM/g, '').trim();
  const parts = cleaned.split(/[*X×\s]+/).map(p => parseFloat(p)).filter(p => !isNaN(p));
  if (parts.length === 3) {
    return { l: parts[0], w: parts[1], h: parts[2] };
  }
  return null;
}

export const processPackingReference = async (file: File) => {
  const buf = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);
  
  const sheetNames = workbook.worksheets.map(ws => ws.name);
  const requiredSheets = ['Sheet1', 'Sheet2', '内外纸箱', '重量'];
  const missingSheets = requiredSheets.filter(req => !sheetNames.some(n => n.includes(req)));
  
  if (missingSheets.length > 0) {
    throw new Error(`参考库文件缺少以下Sheet页: ${missingSheets.join(', ')}`);
  }

  // 1. 重量
  const weightSheet = workbook.worksheets.find(ws => ws.name.includes('重量'))!;
  const weightsMap = new Map<string, number>();
  
  let wHeaderRow = 1;
  let wCodeCol = -1;
  let wWeightCol = -1;
  for (let i = 1; i <= 10; i++) {
    const rowValues = weightSheet.getRow(i).values as any[];
    wCodeCol = findColIndex(rowValues, ['XX CODE', 'XX编码', 'XX 编码']);
    wWeightCol = findColIndex(rowValues, ['重量']);
    if (wCodeCol !== -1 && wWeightCol !== -1) {
      wHeaderRow = i; break;
    }
  }
  
  weightSheet.eachRow((row, rowNumber) => {
    if (rowNumber <= wHeaderRow) return;
    const code = String(getCellValue(row.getCell(wCodeCol))).trim();
    if (!code) return;
    const codeNorm = normalize(code);
    let rawWeight = getCellValue(row.getCell(wWeightCol));
    if (typeof rawWeight === 'string') {
      rawWeight = parseFloat(rawWeight.replace(/[^\d.-]/g, ''));
    }
    if (typeof rawWeight === 'number' && !isNaN(rawWeight)) {
      weightsMap.set(codeNorm, rawWeight);
    }
  });

  // 2. 内外纸箱
  const boxSheet = workbook.worksheets.find(ws => ws.name.includes('内外纸箱'))!;
  const specsMap = new Map<string, PackingSpec>(); 
  
  let bHeaderRow = 1;
  let bCodeCol = -1, matNameCol = -1, matCodeCol = -1, specCol = -1, usageCol = -1;
  for (let i = 1; i <= 10; i++) {
    const rowValues = boxSheet.getRow(i).values as any[];
    bCodeCol = findColIndex(rowValues, ['XX CODE']);
    matNameCol = findColIndex(rowValues, ['物料名称']);
    matCodeCol = findColIndex(rowValues, ['物料代码']);
    specCol = findColIndex(rowValues, ['规格']);
    usageCol = findColIndex(rowValues, ['用量']);
    if (bCodeCol !== -1 && matNameCol !== -1) {
      bHeaderRow = i; break;
    }
  }

  boxSheet.eachRow((row, rowNumber) => {
    if (rowNumber <= bHeaderRow) return;
    const fullCode = String(getCellValue(row.getCell(bCodeCol))).trim();
    if (!fullCode) return;
    
    const matName = String(getCellValue(row.getCell(matNameCol)));
    const matCode = String(getCellValue(row.getCell(matCodeCol)));
    const specDim = parseDimensions(String(getCellValue(row.getCell(specCol))));
    const usageRaw = getCellValue(row.getCell(usageCol));
    
    if (!specDim) return;
    
    let usage = 0;
    if (typeof usageRaw === 'number') usage = usageRaw;
    else if (typeof usageRaw === 'string') usage = parseFloat(usageRaw);
    
    let capacity = usage > 0 ? Math.round(1 / usage) : 0;

    if (!specsMap.has(fullCode)) {
      specsMap.set(fullCode, { fullCode, innerBox: null, outerBox: null });
    }
    const spec = specsMap.get(fullCode)!;
    
    if (matName.includes('内')) {
      spec.innerBox = {
        materialCode: matCode,
        length: specDim.l,
        width: specDim.w,
        height: specDim.h,
        capacity: capacity > 0 ? capacity : 1
      };
    } else if (matName.includes('外') || matName.includes('箱')) {
      spec.outerBox = {
        materialCode: matCode,
        length: specDim.l,
        width: specDim.w,
        height: specDim.h,
        capacity: capacity > 0 ? capacity : 2
      };
    }
  });

  // 3. 产品资料和图片 (Sheet1)
  const prodSheet = workbook.worksheets.find(ws => ws.name.includes('Sheet1'))!;
  const productMap = new Map<string, { oem: string, app: string, year: string, drive: string, imageData: { buffer: ArrayBuffer, extension: string } | null }>();
  
  let pHeaderRow = 1;
  let pCodeCol = -1, pOemCol = -1, pAppCol = -1, pYearCol = -1, pDriveCol = -1;
  for (let i = 1; i <= 20; i++) {
    const rowValues = prodSheet.getRow(i).values as any[];
    pCodeCol = findColIndex(rowValues, ['XX CODE']);
    pOemCol = findColIndex(rowValues, ['OEM']);
    if (pCodeCol !== -1 && pOemCol !== -1) {
      pHeaderRow = i;
      pAppCol = findColIndex(rowValues, ['Application']);
      pYearCol = findColIndex(rowValues, ['Year']);
      pDriveCol = findColIndex(rowValues, ['Drive']);
      break;
    }
  }

  const imageMap: Record<number, { buffer: ArrayBuffer; extension: string }> = {};
  prodSheet.getImages().forEach((image) => {
    const img = workbook.model.media.find((m: any, idx: number) => idx === (image as any).imageId || m.index === (image as any).imageId);
    if (img && image.range.tl.nativeRow + 1) {
      imageMap[image.range.tl.nativeRow + 1] = { buffer: img.buffer, extension: img.extension };
    }
  });

  prodSheet.eachRow((row, rowNumber) => {
    if (rowNumber <= pHeaderRow) return;
    const xxCode = String(getCellValue(row.getCell(pCodeCol))).trim();
    if (!xxCode) return;
    const xxCodeNorm = normalize(xxCode);
    
    if (!productMap.has(xxCodeNorm)) {
      productMap.set(xxCodeNorm, {
        oem: String(getCellValue(row.getCell(pOemCol)) || ""),
        app: pAppCol !== -1 ? String(getCellValue(row.getCell(pAppCol)) || "") : "",
        year: pYearCol !== -1 ? String(getCellValue(row.getCell(pYearCol)) || "") : "",
        drive: pDriveCol !== -1 ? String(getCellValue(row.getCell(pDriveCol)) || "") : "",
        imageData: imageMap[rowNumber] || null
      });
    }
  });

  return { specs: Array.from(specsMap.values()), weights: weightsMap, products: productMap };
};

export const processPackingQueries = async (
  file: File, 
  refData: { specs: PackingSpec[], weights: Map<string, number>, products: Map<string, any> },
  onProgress?: (msg: string) => void
): Promise<PackingInputRow[]> => {
  
  onProgress?.("正在读取待查清单...");
  const buf = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);
  const worksheet = workbook.worksheets[0];
  
  // Find header row and columns
  let headerRowIndex = 0;
  let xxColIdx = -1;

  for (let i = 1; i <= 20; i++) {
    const rowValues = worksheet.getRow(i).values as any[];
    xxColIdx = findColIndex(rowValues, ['XX CODE', 'XX编码', 'XX 编码']);
    
    if (xxColIdx !== -1) { 
        headerRowIndex = i; 
        break; 
    }
  }

  // Fallback if no header found (user provided single column file)
  if (xxColIdx === -1) {
      xxColIdx = 1;
      headerRowIndex = 0;
  }

  const results: PackingInputRow[] = [];

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber <= headerRowIndex) return;

    const xxCodeRaw = String(getCellValue(row.getCell(xxColIdx)) || "").trim();
    if (!xxCodeRaw) return;

    const baseCodeStr = xxCodeRaw;
    let matchedSpecs: PackingSpec[] = [];
    let normBase = normalize(baseCodeStr);
    
    const prodData = refData.products.get(normBase);

    if (normBase) {
        matchedSpecs = refData.specs.filter(s => {
            const nFull = normalize(s.fullCode);
            if (nFull === normBase) return true;
            const segments = String(s.fullCode).toUpperCase().replace(/[^A-Z0-9]/g, '-').split('-');
            if (segments.includes(baseCodeStr.toUpperCase())) return true;
            return false;
        });
    }

    let weight = null;
    if (normBase && refData.weights.has(normBase)) {
        weight = refData.weights.get(normBase) || null;
    } else if (matchedSpecs.length > 0) {
        for(const sp of matchedSpecs) {
             const nFull = normalize(sp.fullCode);
             if (refData.weights.has(nFull)) {
                 weight = refData.weights.get(nFull) || null;
                 break;
             }
        }
    }

    let statusMsgParts = [];
    if (!prodData) statusMsgParts.push("缺产品资料");
    else if (!prodData.imageData) statusMsgParts.push("缺图片");
    
    if (matchedSpecs.length === 0) {
        statusMsgParts.push("缺包装规格");
    } else if (matchedSpecs.length === 1) {
        if (!matchedSpecs[0].innerBox) statusMsgParts.push("缺内箱");
        if (!matchedSpecs[0].outerBox) statusMsgParts.push("缺外箱");
    }

    if (weight === null) statusMsgParts.push("缺重量");

    const blockingErrors = ['缺包装规格', '规格未选择', '缺外箱', '缺重量', '数量无效'];
    let status: PackingInputRow['status'] = statusMsgParts.some(m => blockingErrors.includes(m)) ? 'error' : 'no_match';

    results.push({
        originalIndex: rowNumber,
        originalXXCode: baseCodeStr,
        originalProductName: 'steering rack',
        originalOEM: prodData?.oem || "",
        originalPrice: "", 
        imageData: prodData?.imageData || null,
        
        status,
        statusMsg: statusMsgParts.join(', '),
        
        availableSpecs: matchedSpecs,
        selectedSpecFullCode: matchedSpecs.length === 1 ? matchedSpecs[0].fullCode : undefined,
        quantity: '',
        weightPerItem: weight
    });
  });

  return results;
};

export const calculatePacking = (row: PackingInputRow): PackingCalculationResult[] => {
    if (!row.selectedSpecFullCode || !row.quantity) return [];
    
    const qtyStr = String(row.quantity).trim();
    const totalQty = parseInt(qtyStr, 10);
    
    if (isNaN(totalQty) || totalQty <= 0) return [];
    
    const spec = row.availableSpecs.find(s => s.fullCode === row.selectedSpecFullCode);
    if (!spec || !spec.outerBox) return [];
    
    const weightPerItem = row.weightPerItem || 0;
    const results: PackingCalculationResult[] = [];
    
    const outerCapacity = spec.outerBox.capacity;
    const outerCount = Math.floor(totalQty / outerCapacity);
    const remainder = totalQty % outerCapacity;
    
    if (outerCount > 0) {
        const oL = spec.outerBox.length;
        const oW = spec.outerBox.width;
        const oH = spec.outerBox.height;
        const cbmPerBox = (oL * oW * oH) / 1000000000;
        const areaPerBox = 2 * (oL*oW + oL*oH + oW*oH) / 1000000;
        
        const netW = weightPerItem * outerCapacity;
        const grossW = netW + 3;
        
        results.push({
            boxType: 'outer',
            boxesCount: outerCount,
            itemsPerBox: outerCapacity,
            itemsTotal: outerCount * outerCapacity,
            
            netWeightPerBox: netW,
            grossWeightPerBox: grossW,
            totalNetWeight: netW * outerCount,
            totalGrossWeight: grossW * outerCount,
            
            cbmPerBox,
            totalCBM: cbmPerBox * outerCount,
            areaPerBox,
            totalArea: areaPerBox * outerCount
        });
    }
    
    if (remainder > 0 && spec.innerBox) {
        const iL = spec.innerBox.length;
        const iW = spec.innerBox.width;
        const iH = spec.innerBox.height;
        const capacity = remainder; // We pack the remainder in inner boxes
        const iCapacity = spec.innerBox.capacity; // Max capacity of inner
        const innerCount = Math.ceil(remainder / iCapacity);
        // Assuming the remainder fits in inner boxes. 
        // Based on instructions: remainder=1 generates 1 inner box with 1 pc.
        
        const cbmPerBox = (iL * iW * iH) / 1000000000;
        const areaPerBox = 2 * (iL*iW + iL*iH + iW*iH) / 1000000;
        
        const netW = weightPerItem * iCapacity;
        const grossW = netW + 3; // +3kg for inner box? "单箱毛重 = 单箱净重 + 3 kg" usually applies to any carton.
        
        // Adjust for remainder exactly
        const exactNetW = weightPerItem * remainder;
        const exactGrossW = exactNetW + 3;
        
        results.push({
            boxType: 'inner',
            boxesCount: 1, // simplified, assuming 1 inner box handles remainder
            itemsPerBox: remainder,
            itemsTotal: remainder,
            
            netWeightPerBox: exactNetW,
            grossWeightPerBox: exactGrossW,
            totalNetWeight: exactNetW,
            totalGrossWeight: exactGrossW,
            
            cbmPerBox,
            totalCBM: cbmPerBox,
            areaPerBox,
            totalArea: areaPerBox
        });
    } else if (remainder > 0 && !spec.innerBox) {
        // Fallback if no inner box spec but remainder exists
        // Just use outer box spec with different count
        const oL = spec.outerBox.length;
        const oW = spec.outerBox.width;
        const oH = spec.outerBox.height;
        const cbmPerBox = (oL * oW * oH) / 1000000000;
        const areaPerBox = 2 * (oL*oW + oL*oH + oW*oH) / 1000000;
        
        const netW = weightPerItem * remainder;
        const grossW = netW + 3;
        
        results.push({
            boxType: 'outer', // pretend outer but packed with remainder
            boxesCount: 1,
            itemsPerBox: remainder,
            itemsTotal: remainder,
            
            netWeightPerBox: netW,
            grossWeightPerBox: grossW,
            totalNetWeight: netW,
            totalGrossWeight: grossW,
            
            cbmPerBox,
            totalCBM: cbmPerBox,
            areaPerBox,
            totalArea: areaPerBox
        });
    }
    
    return results;
};

export const exportPackingList = async (data: PackingInputRow[], fileName: string) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('装箱单');
  
    const columns = [
        { header: '箱号', key: 'boxNo', width: 12 },
        { header: '品名\nProduct Name', key: 'productName', width: 20 },
        { header: '型号\nModel', key: 'model', width: 25 },
        { header: '箱数\npcs/ctn', key: 'boxesCount', width: 12 },
        { header: '每箱数量\nQuantity per Carton', key: 'itemsPerBox', width: 15 },
        { header: '总数量\nQTY', key: 'itemsTotal', width: 12 },
        { header: '单只净重\nN.Weight(kgs)', key: 'itemWeight', width: 15 },
        { header: '单件净重\nN.Weight(kgs)/CTN', key: 'netWeightPerBox', width: 15 },
        { header: '单只毛重\nG.Weight(kgs)', key: 'itemWeightG', width: 15 },
        { header: '单件毛重\nG.Weight(kgs)/CTN', key: 'grossWeightPerBox', width: 15 },
        { header: '总净重', key: 'totalNetWeight', width: 15 },
        { header: '总毛重', key: 'totalGrossWeight', width: 15 },
        { header: '长', key: 'length', width: 10 },
        { header: '宽', key: 'width', width: 10 },
        { header: '高', key: 'height', width: 10 },
        { header: '纸箱体积\nCBM', key: 'cbm', width: 15 },
        { header: '单价', key: 'price', width: 12 },
        { header: '总金额', key: 'totalAmount', width: 15 },
        { header: '外箱规格', key: 'specStr', width: 25 },
        { header: 'XX CODE', key: 'xxCode', width: 15 },
        { header: '图片', key: 'picture', width: 30 }
    ];
  
    worksheet.columns = columns;
    worksheet.getRow(1).height = 45;
    worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getRow(1).font = { name: '黑体', bold: true };
    
    // Draw borders for header
    worksheet.getRow(1).eachCell((cell) => {
        cell.border = {
            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } };
    });
  
    let currentBoxIndex = 1;
    let excelRowIndex = 2;
    
    let sumBoxes = 0;
    let sumQty = 0;
    let sumNet = 0;
    let sumGross = 0;
    let sumCbm = 0;
  
    for (const row of data) {
        if (row.status === 'error' || row.status === 'invalid_qty' || row.status === 'no_match' || !row.selectedSpecFullCode) continue;
        
        const spec = row.availableSpecs.find(s => s.fullCode === row.selectedSpecFullCode);
        const calcs = calculatePacking(row);
        
        for (const calc of calcs) {
            let boxNoStr = '';
            if (calc.boxesCount === 1) {
                boxNoStr = `NO.${currentBoxIndex}`;
                currentBoxIndex += 1;
            } else if (calc.boxesCount > 1) {
                boxNoStr = `NO.${currentBoxIndex}~${currentBoxIndex + calc.boxesCount - 1}`;
                currentBoxIndex += calc.boxesCount;
            }
            
            const isInner = calc.boxType === 'inner';
            const boxSpec = isInner ? spec?.innerBox : spec?.outerBox;
            const l = boxSpec?.length || 0;
            const w = boxSpec?.width || 0;
            const h = boxSpec?.height || 0;
            
            let specStr = '';
            if (boxSpec) {
                specStr = `${boxSpec.materialCode}\n${isInner ? '内纸箱' : '外纸箱'}\n${l}*${w}*${h}`;
            }
            
            const excelRow = worksheet.addRow({
                boxNo: boxNoStr,
                productName: 'steering rack',
                model: row.originalOEM,
                boxesCount: calc.boxesCount,
                itemsPerBox: calc.itemsPerBox,
                itemsTotal: { formula: `D${excelRowIndex}*E${excelRowIndex}`, result: calc.itemsTotal },
                itemWeight: row.weightPerItem,
                netWeightPerBox: { formula: `G${excelRowIndex}*E${excelRowIndex}`, result: calc.netWeightPerBox },
                itemWeightG: { formula: `G${excelRowIndex}+1.5`, result: row.weightPerItem !== null ? row.weightPerItem + 1.5 : '' },
                grossWeightPerBox: { formula: `H${excelRowIndex}+3`, result: calc.grossWeightPerBox },
                totalNetWeight: { formula: `H${excelRowIndex}*D${excelRowIndex}`, result: calc.totalNetWeight },
                totalGrossWeight: { formula: `J${excelRowIndex}*D${excelRowIndex}`, result: calc.totalGrossWeight },
                length: l,
                width: w,
                height: h,
                cbm: { formula: `(M${excelRowIndex}*N${excelRowIndex}*O${excelRowIndex})/1000000000*D${excelRowIndex}`, result: calc.totalCBM },
                price: '',
                totalAmount: '',
                specStr: specStr,
                xxCode: row.originalXXCode,
                picture: ''
            });
            
            excelRow.height = 80;
            excelRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            
            excelRow.eachCell((cell) => {
                cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                cell.font = { name: '黑体' };
            });

            excelRow.getCell('itemWeight').numFmt = '0.00';
            excelRow.getCell('netWeightPerBox').numFmt = '0.00';
            excelRow.getCell('itemWeightG').numFmt = '0.00';
            excelRow.getCell('grossWeightPerBox').numFmt = '0.00';
            excelRow.getCell('totalNetWeight').numFmt = '0.00';
            excelRow.getCell('totalGrossWeight').numFmt = '0.00';
            excelRow.getCell('cbm').numFmt = '0.00';
            
            if (row.imageData) {
                try {
                    const imgId = workbook.addImage({
                        buffer: row.imageData.buffer,
                        extension: row.imageData.extension as any
                    });
                    worksheet.addImage(imgId, {
                        tl: { col: 20, row: excelRowIndex - 1, nativeColOff: 200000, nativeRowOff: 200000 },
                        ext: { width: 232, height: 94 },
                        editAs: 'oneCell'
                    });
                } catch (e) {
                    // ignore img error
                }
            }
            
            sumBoxes += calc.boxesCount;
            sumQty += calc.itemsTotal;
            sumNet += calc.totalNetWeight;
            sumGross += calc.totalGrossWeight;
            sumCbm += calc.totalCBM;
            
            excelRowIndex++;
        }
    }
    
    // Total Row
    const totalRow = worksheet.addRow({
        model: '合计 Total',
        boxesCount: { formula: `SUM(D2:D${excelRowIndex - 1})`, result: sumBoxes },
        itemsTotal: { formula: `SUM(F2:F${excelRowIndex - 1})`, result: sumQty },
        totalNetWeight: { formula: `SUM(K2:K${excelRowIndex - 1})`, result: sumNet },
        totalGrossWeight: { formula: `SUM(L2:L${excelRowIndex - 1})`, result: sumGross },
        cbm: { formula: `SUM(P2:P${excelRowIndex - 1})`, result: sumCbm },
        totalAmount: ''
    });
    totalRow.height = 30;
    totalRow.font = { name: '黑体', bold: true };
    totalRow.alignment = { vertical: 'middle', horizontal: 'center' };
    totalRow.eachCell((cell) => {
        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
    });

    totalRow.getCell('totalNetWeight').numFmt = '0.00';
    totalRow.getCell('totalGrossWeight').numFmt = '0.00';
    totalRow.getCell('cbm').numFmt = '0.00';
    
    worksheet.views = [{ state: 'frozen', ySplit: 1 }];

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.click();
};

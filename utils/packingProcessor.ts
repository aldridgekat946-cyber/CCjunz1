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
  const wb = XLSX.read(buf, { type: 'array' });
  
  const boxSheetName = wb.SheetNames.find(n => n.includes('内外纸箱'));
  const weightSheetName = wb.SheetNames.find(n => n.includes('重量'));
  
  if (!boxSheetName || !weightSheetName) {
    throw new Error('参考库文件必须包含“内外纸箱”和“重量”两个Sheet页。');
  }

  // 1. 解析重量
  const weightSheet = wb.Sheets[weightSheetName];
  const weightDataRaw = XLSX.utils.sheet_to_json<any[]>(weightSheet, { header: 1 });
  const weightsMap = new Map<string, number>();
  
  let headerRow = 0;
  for (let i=0; i<Math.min(10, weightDataRaw.length); i++) {
      if (findColIndex(weightDataRaw[i], ['XX CODE', 'XX编码', 'XX 编码']) !== -1) {
          headerRow = i; break;
      }
  }
  const wHeaders = weightDataRaw[headerRow];
  const wCodeCol = findColIndex(wHeaders, ['XX CODE', 'XX编码', 'XX 编码']);
  const weightCol = findColIndex(wHeaders, ['重量']);
  
  if (wCodeCol !== -1 && weightCol !== -1) {
      for (let i = headerRow + 1; i < weightDataRaw.length; i++) {
          const row = weightDataRaw[i];
          if (!row || !row[wCodeCol]) continue;
          
          const codeNorm = normalize(row[wCodeCol]);
          let rawWeight = row[weightCol];
          if (typeof rawWeight === 'string') {
              rawWeight = parseFloat(rawWeight.replace(/[^\d.-]/g, ''));
          }
          if (typeof rawWeight === 'number' && !isNaN(rawWeight)) {
              weightsMap.set(codeNorm, rawWeight);
          }
      }
  }

  // 2. 解析纸箱
  const boxSheet = wb.Sheets[boxSheetName];
  const boxDataRaw = XLSX.utils.sheet_to_json<any[]>(boxSheet, { header: 1 });
  
  let bHeaderRow = 0;
  for (let i=0; i<Math.min(10, boxDataRaw.length); i++) {
      if (findColIndex(boxDataRaw[i], ['XX CODE', 'XX编码', '产品号']) !== -1) {
          bHeaderRow = i; break;
      }
  }
  const bHeaders = boxDataRaw[bHeaderRow];
  const bCodeCol = findColIndex(bHeaders, ['XX CODE']);
  const matNameCol = findColIndex(bHeaders, ['物料名称']);
  const matCodeCol = findColIndex(bHeaders, ['物料代码']);
  const specCol = findColIndex(bHeaders, ['规格']);
  const usageCol = findColIndex(bHeaders, ['用量']);

  const specsMap = new Map<string, PackingSpec>(); // fullCode -> Spec

  for (let i = bHeaderRow + 1; i < boxDataRaw.length; i++) {
      const row = boxDataRaw[i];
      if (!row || !row[bCodeCol]) continue;
      
      const fullCode = String(row[bCodeCol]).trim();
      const matName = String(row[matNameCol] || "");
      const matCode = String(row[matCodeCol] || "");
      const specDim = parseDimensions(String(row[specCol] || ""));
      const usageRaw = row[usageCol];
      
      if (!fullCode || !specDim) continue;
      
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
              capacity: capacity > 0 ? capacity : 1 // fallback inner = 1
          };
      } else if (matName.includes('外') || matName.includes('箱')) {
          spec.outerBox = {
              materialCode: matCode,
              length: specDim.l,
              width: specDim.w,
              height: specDim.h,
              capacity: capacity > 0 ? capacity : 2 // fallback outer = 2
          };
      }
  }

  return { specs: Array.from(specsMap.values()), weights: weightsMap };
};

export const processPackingQueries = async (
  file: File, 
  refData: { specs: PackingSpec[], weights: Map<string, number> },
  onProgress?: (msg: string) => void
): Promise<PackingInputRow[]> => {
  
  onProgress?.("正在读取待查清单...");
  const buf = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);
  const worksheet = workbook.worksheets[0];
  
  // Find header row and columns
  let headerRowIndex = 1;
  let oemColIdx = -1;
  let oeInputColIdx = -1;
  let xxColIdx = -1;
  let priceColIdx = -1;
  let prodColIdx = -1;

  for (let i = 1; i <= 20; i++) {
    const rowValues = worksheet.getRow(i).values as any[];
    oemColIdx = findColIndex(rowValues, ['OEM', '型号']);
    oeInputColIdx = findColIndex(rowValues, ['OE', '查询', '输入 OE']);
    xxColIdx = findColIndex(rowValues, ['XX CODE', 'XX编码', 'XX 编码']);
    priceColIdx = findColIndex(rowValues, ['广州价', '价格', '单价', '单价']);
    prodColIdx = findColIndex(rowValues, ['产品名', '名称', '品名']);
    
    // We need either XX CODE or Input OE
    if (xxColIdx !== -1 || oeInputColIdx !== -1) { 
        headerRowIndex = i; 
        break; 
    }
  }

  const imageMap: Record<number, { buffer: ArrayBuffer; extension: string }> = {};
  worksheet.getImages().forEach((image) => {
    const img = workbook.model.media.find((m: any, idx: number) => idx === (image as any).imageId || m.index === (image as any).imageId);
    if (img && image.range.tl.nativeRow + 1) {
      imageMap[image.range.tl.nativeRow + 1] = { buffer: img.buffer, extension: img.extension };
    }
  });

  const results: PackingInputRow[] = [];

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber <= headerRowIndex) return;

    let xxCodeRaw = xxColIdx !== -1 ? String(getCellValue(row.getCell(xxColIdx)) || "").trim() : "";
    const oeInputRaw = oeInputColIdx !== -1 ? String(getCellValue(row.getCell(oeInputColIdx)) || "").trim() : "";
    const oemRaw = oemColIdx !== -1 ? String(getCellValue(row.getCell(oemColIdx)) || "") : "";
    const priceRaw = priceColIdx !== -1 ? getCellValue(row.getCell(priceColIdx)) : null;
    const prodRaw = prodColIdx !== -1 ? String(getCellValue(row.getCell(prodColIdx)) || "") : "";

    // fallback extraction of XX CODE from input OE if missing
    if (!xxCodeRaw && oeInputRaw) {
        const tokens = oeInputRaw.split(/[\s\n,;:/|，；、]+/);
        for (const t of tokens) {
            const cleanT = t.trim();
            if (cleanT.toUpperCase().startsWith('X')) {
                xxCodeRaw = cleanT;
                break;
            }
        }
    }

    if (!xxCodeRaw && !oeInputRaw && !oemRaw) return;

    const baseCodeStr = String(xxCodeRaw).trim();
    // find all specs where fullCode includes -XXCODE- or starts with XXCODE-
    // Example: base is X062. fullCode is ZX-X062-020A -> matches
    let matchedSpecs: PackingSpec[] = [];
    let normBase = normalize(baseCodeStr);
    
    if (normBase) {
        matchedSpecs = refData.specs.filter(s => {
            const nFull = normalize(s.fullCode);
            // exact match or bounded segment match
            if (nFull === normBase) return true;
            // Bounded match by replacing non-alnum with hyphens in fullCode and checking parts
            const segments = String(s.fullCode).toUpperCase().replace(/[^A-Z0-9]/g, '-').split('-');
            if (segments.includes(baseCodeStr.toUpperCase())) return true;
            return false;
        });
    }

    let status: PackingInputRow['status'] = 'no_match';
    let statusMsg = '';
    
    if (matchedSpecs.length > 0) {
        status = 'matched';
    } else {
        statusMsg = '未找到规格';
    }
    
    let weight = null;
    if (normBase && refData.weights.has(normBase)) {
        weight = refData.weights.get(normBase) || null;
    } else if (matchedSpecs.length > 0) {
        // try to match weight by full code if base code failed
        for(const sp of matchedSpecs) {
             const nFull = normalize(sp.fullCode);
             if (refData.weights.has(nFull)) {
                 weight = refData.weights.get(nFull) || null;
                 break;
             }
        }
    }

    if (status === 'matched' && weight === null) {
        status = 'error';
        statusMsg = '缺重量';
    }

    results.push({
        originalIndex: rowNumber,
        originalXXCode: baseCodeStr || oeInputRaw,
        originalProductName: prodRaw,
        originalOEM: oemRaw,
        originalPrice: priceRaw,
        imageData: imageMap[rowNumber] || null,
        
        status,
        statusMsg,
        
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
                itemWeightG: '',
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

            excelRow.getCell('netWeightPerBox').numFmt = '0.000';
            excelRow.getCell('grossWeightPerBox').numFmt = '0.000';
            excelRow.getCell('totalNetWeight').numFmt = '0.000';
            excelRow.getCell('totalGrossWeight').numFmt = '0.000';
            excelRow.getCell('cbm').numFmt = '0.000000';
            
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

    totalRow.getCell('totalNetWeight').numFmt = '0.000';
    totalRow.getCell('totalGrossWeight').numFmt = '0.000';
    totalRow.getCell('cbm').numFmt = '0.000000';
    
    worksheet.views = [{ state: 'frozen', ySplit: 1 }];

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.click();
};


import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { GoogleGenAI, Type } from "@google/genai";
import { Box1Data, ProcessedRow } from '../types';

export const normalize = (s: any): string => {
  if (s === null || s === undefined) return "";
  const str = typeof s === 'object' ? (s.text || s.result || String(s)) : String(s);
  return str.replace(/[^A-Za-z0-9]/g, '').toUpperCase();
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

/**
 * 指数退避重试助手，处理 429 配额错误
 */
async function callWithRetry<T>(
  fn: () => Promise<T>, 
  onRetry?: (msg: string) => void,
  maxRetries: number = 3, 
  initialDelay: number = 2000
): Promise<T> {
  let lastError: any;
  for (let i = 0; i < maxRetries; i++) {
    try {
      return await fn();
    } catch (error: any) {
      lastError = error;
      const isQuotaError = error.status === 429 || JSON.stringify(error).includes('429') || JSON.stringify(error).includes('RESOURCE_EXHAUSTED');
      if (isQuotaError && i < maxRetries - 1) {
        const delay = initialDelay * Math.pow(2, i) + Math.random() * 500;
        onRetry?.(`配额限制，${Math.round(delay/1000)}秒后重试...`);
        await new Promise(resolve => setTimeout(resolve, delay));
        continue;
      }
      throw error;
    }
  }
  throw lastError;
}

/**
 * AI 检索结果缓存，避免重复查询
 */
const aiCache: Record<string, { productName: string, model: string, generalOE: string }> = {};

/**
 * 优化后的 AI 检索：使用 Google Search 快速获取零件信息
 */
export async function fetchPartInfoFromAI(oe: string): Promise<{ productName: string, model: string, generalOE: string }> {
  const normOE = normalize(oe);
  if (aiCache[normOE]) return aiCache[normOE];

  const result = await callWithRetry(async () => {
    const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
    const response = await ai.models.generateContent({
      model: 'gemini-3.1-flash-lite',
      contents: `Search for automotive part info for OE "${oe}". 
      Required (JSON only):
      - productName: Basic type (e.g. Starter)
      - model: Key vehicle models
      - generalOE: Comma-separated cross OE numbers`,
      config: {
        tools: [{ googleSearch: {} }],
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            productName: { type: Type.STRING },
            model: { type: Type.STRING },
            generalOE: { type: Type.STRING },
          },
          required: ["productName", "model", "generalOE"]
        }
      }
    });

    const text = response.text;
    return JSON.parse(text || "{}");
  });

  if (result && result.productName) {
    aiCache[normOE] = result;
  }
  return result;
}

export const processFiles = async (
  fileReference: File,
  fileOe: File,
  onProgress?: (msg: string) => void
): Promise<{ results: ProcessedRow[], knownOEs: Set<string> }> => {
  onProgress?.("正在解析库文件...");
  const refBuffer = await fileReference.arrayBuffer();
  const refWorkbook = new ExcelJS.Workbook();
  await refWorkbook.xlsx.load(refBuffer);
  const refWorksheet = refWorkbook.worksheets[0];

  let headerRowIndex = 1;
  let oemColIdx = -1;
  for (let i = 1; i <= 20; i++) {
    const rowValues = refWorksheet.getRow(i).values as any[];
    oemColIdx = findColIndex(rowValues, ['OEM', 'OE', '原厂编号', '零件号']);
    if (oemColIdx !== -1) { headerRowIndex = i; break; }
  }

  if (oemColIdx === -1) throw new Error("库文件中未找到 OEM 列");

  const headers = refWorksheet.getRow(headerRowIndex).values as any[];
  const xxColIdx = findColIndex(headers, ['XX CODE', 'XX编码']);
  const appColIdx = findColIndex(headers, ['Application', '适用车型']);
  const yearColIdx = findColIndex(headers, ['Year', '年份', '年度']);
  const driveColIdx = findColIndex(headers, ['Drive', '驱动', '左/右']);
  const priceColIdx = findColIndex(headers, ['广州', 'Price', '价格']);
  const prodColIdx = findColIndex(headers, ['Product', 'Description', '产品名']);

  const imageMap: Record<number, { buffer: ArrayBuffer; extension: string }> = {};
  refWorksheet.getImages().forEach((image) => {
    const img = refWorkbook.model.media.find((m: any, idx: number) => idx === (image as any).imageId || m.index === (image as any).imageId);
    if (img && image.range.tl.nativeRow + 1) {
      imageMap[image.range.tl.nativeRow + 1] = { buffer: img.buffer, extension: img.extension };
    }
  });

  const mapRef: Record<string, Box1Data> = {};
  const knownOEs = new Set<string>();

  refWorksheet.eachRow((row, rowNumber) => {
    if (rowNumber <= headerRowIndex) return;
    const oemRaw = getCellValue(row.getCell(oemColIdx));
    if (!oemRaw) return;
    const tokens = String(oemRaw).split(/[\s\n,;:/|，；、]+/);
    for (const token of tokens) {
      const norm = normalize(token);
      if (norm.length > 2) {
        knownOEs.add(norm);
        if (!mapRef[norm]) {
          mapRef[norm] = {
            xxCode: xxColIdx !== -1 ? String(getCellValue(row.getCell(xxColIdx)) || "") : "",
            application: appColIdx !== -1 ? String(getCellValue(row.getCell(appColIdx)) || "") : "",
            year: yearColIdx !== -1 ? String(getCellValue(row.getCell(yearColIdx)) || "") : "", 
            oem: String(oemRaw),
            drive: driveColIdx !== -1 ? String(getCellValue(row.getCell(driveColIdx)) || "") : "",
            picture: "已提取",
            productName: prodColIdx !== -1 ? String(getCellValue(row.getCell(prodColIdx)) || "") : "",
            price: priceColIdx !== -1 ? getCellValue(row.getCell(priceColIdx)) : null,
            imageData: imageMap[rowNumber] || null
          };
        }
        // Rule: 44250 and 44200 are interchangeable
        let ruleNorm = norm;
        if (ruleNorm.startsWith('44250')) {
          ruleNorm = '44200' + ruleNorm.substring(5);
        } else if (ruleNorm.startsWith('44200')) {
          ruleNorm = '44250' + ruleNorm.substring(5);
        }
        if (ruleNorm !== norm && !mapRef[ruleNorm]) {
          mapRef[ruleNorm] = mapRef[norm];
        }
      }
    }
  });

  onProgress?.("正在处理待查清单...");
  const bufOe = await fileOe.arrayBuffer();
  const wbOe = XLSX.read(bufOe, { type: 'array' });
  const wsOe = wbOe.Sheets[wbOe.SheetNames[0]];
  const dataOeRaw = XLSX.utils.sheet_to_json<any[]>(wsOe, { header: 1 });
  
  let oeInputCol = findColIndex(dataOeRaw[0] || [], ['OE', 'OEM', '输入']) - 1;
  if (oeInputCol < 0) oeInputCol = 0;

  const results: ProcessedRow[] = [];
  const startIdx = (typeof dataOeRaw[0][oeInputCol] === 'string' && dataOeRaw[0][oeInputCol].length > 0) ? 1 : 0;

  // 1. 预处理所有行，识别需要 AI 检索的项
  const tasks: { index: number; inputOE: string; normInput: string; match: Box1Data | null }[] = [];
  for (let i = startIdx; i < dataOeRaw.length; i++) {
    const row = dataOeRaw[i];
    if (!row || !row[oeInputCol]) continue;
    const inputOE = String(row[oeInputCol]).trim();
    const normInput = normalize(inputOE);
    tasks.push({ index: i, inputOE, normInput, match: mapRef[normInput] });
  }

  // 2. 并行处理逻辑 (带并发控制)
  const CONCURRENCY = 5; // 同时进行的 AI 请求数
  const totalTasks = tasks.length;
  let completedCount = 0;

  const processTask = async (task: typeof tasks[0]) => {
    const { inputOE, match } = task;
    const normInput = normalize(inputOE);
    let isSpecialMatch = false;
    if (match) {
        const oemTokens = String(match.oem).split(/[\s\n,;:/|，；、]+/).map(t => normalize(t)).filter(t => t.length > 0);
        if (!oemTokens.includes(normInput)) {
            let ruleNorm = normInput;
            if (ruleNorm.startsWith('44250')) {
              ruleNorm = '44200' + ruleNorm.substring(5);
            } else if (ruleNorm.startsWith('44200')) {
              ruleNorm = '44250' + ruleNorm.substring(5);
            }
            if (oemTokens.includes(ruleNorm)) {
                isSpecialMatch = true;
            }
        }
    }
    const newRow: ProcessedRow = {
      '输入 OE': inputOE,
      'XX 编码': match?.xxCode || null,
      '适用车型': match?.application || null,
      '年份': match?.year || null,
      'OEM': match?.oem || null,
      '驱动': match?.drive || null,
      '图片': match?.imageData ? "匹配成功" : null,
      '图片数据': match?.imageData || null,
      '广州价': match?.price || null,
      '产品名': match?.productName || null,
      '车型': null,
      '通用OE': null,
      isSpecialMatch
    };

    if (!match) {
      try {
        // 检查缓存
        const normInput = normalize(inputOE);
        if (aiCache[normInput]) {
          const aiInfo = aiCache[normInput];
          newRow['产品名'] = aiInfo.productName;
          newRow['车型'] = aiInfo.model;
          newRow['通用OE'] = aiInfo.generalOE;
        } else {
          onProgress?.(`正在 AI 检索 (${completedCount + 1}/${totalTasks}): ${inputOE}...`);
          const aiInfo = await fetchPartInfoFromAI(inputOE);
          newRow['产品名'] = aiInfo.productName;
          newRow['车型'] = aiInfo.model;
          newRow['通用OE'] = aiInfo.generalOE;
        }
      } catch (e) {
        newRow['产品名'] = "检索失败";
      }
    }

    completedCount++;
    return newRow;
  };

  // 使用并发控制执行所有任务
  const finalResults: ProcessedRow[] = [];
  for (let i = 0; i < tasks.length; i += CONCURRENCY) {
    const batch = tasks.slice(i, i + CONCURRENCY);
    const batchResults = await Promise.all(batch.map(task => processTask(task)));
    finalResults.push(...batchResults);
    onProgress?.(`已完成: ${Math.round((finalResults.length / totalTasks) * 100)}%`);
  }

  return { results: finalResults, knownOEs };
};

export const exportToExcel = async (data: ProcessedRow[], fileName: string, knownOEs: Set<string>) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('匹配结果');

  const columns = ['输入 OE', 'XX 编码', '适用车型', '年份', 'OEM', '驱动', '图片', '广州价', '产品名', '车型', '通用OE'];
  worksheet.columns = columns.map(c => ({ 
    header: c, 
    key: c, 
    width: (c === 'OEM' || c === '适用车型' || c === '车型' || c === '通用OE' || c === '产品名') ? 35 : (c === '图片' ? 34 : 15)
  }));

  worksheet.getRow(1).font = { bold: true };
  worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };

  data.forEach((rowData, i) => {
    const row = worksheet.addRow({});
    row.height = 80;
    row.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

    columns.forEach((col, colIdx) => {
      const cell = row.getCell(colIdx + 1);
      const val = rowData[col];

      if (col === '图片' && rowData['图片数据']) {
        const img = rowData['图片数据'];
        const imageId = workbook.addImage({ buffer: img.buffer, extension: img.extension as any });
        worksheet.addImage(imageId, {
          tl: { col: colIdx, row: i + 1, nativeColOff: 100000, nativeRowOff: 0 },
          ext: { width: 220, height: 80 }
        });
      } else if (col === 'OEM' || col === '通用OE') {
        const inputNorm = normalize(rowData['输入 OE']);
        const tokens = String(val || "").split(/([\s\n,;:/|，；、]+)/);
        const richText: any[] = [];
        tokens.forEach(t => {
          if (!t) return;
          const normT = normalize(t);
          
          if (col === 'OEM') {
            let ruleNorm = inputNorm;
            if (ruleNorm.startsWith('44250')) {
              ruleNorm = '44200' + ruleNorm.substring(5);
            } else if (ruleNorm.startsWith('44200')) {
              ruleNorm = '44250' + ruleNorm.substring(5);
            }

            // OEM 列只显示红色高亮 (匹配输入 OE) 或紫色高亮 (特殊规则匹配)
            if (normT === inputNorm) {
              richText.push({ text: t, font: { color: { argb: 'FFFF0000' }, bold: true } });
            } else if (rowData.isSpecialMatch && normT === ruleNorm) {
              richText.push({ text: t, font: { color: { argb: 'FF800080' }, bold: true } });
            } else {
              richText.push({ text: t });
            }
          } else if (col === '通用OE') {
            // 通用 OE 列只显示绿色高亮 (匹配库内已存在)
            if (knownOEs.has(normT)) {
              richText.push({ text: t, font: { color: { argb: 'FF00B050' }, bold: true } });
            } else {
              richText.push({ text: t });
            }
          }
        });
        cell.value = richText.length > 0 ? { richText } : val;
      } else {
        cell.value = val;
      }
    });
  });

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.click();
};

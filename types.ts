
export interface ProcessedRow {
  '输入 OE': string;
  'XX 编码': string | null;
  '适用车型': string | null;
  '年份': string | null;
  'OEM': string | null;
  '驱动': string | null;
  '图片': string | null;
  '图片数据'?: {
    buffer: ArrayBuffer;
    extension: string;
  } | null;
  '广州价': string | number | null;
  '产品名'?: string | null;
  '车型'?: string | null;
  '通用OE'?: string | null;
  isSpecialMatch?: boolean;
  [key: string]: any;
}

export type FileType = 'reference' | 'oe';

export interface FileState {
  file: File | null;
  name: string;
}

export interface Box1Data {
  xxCode: string;
  application: string;
  year: string;
  oem: string;
  drive: string;
  picture: string;
  productName: string;
  price: string | number | null;
  imageData?: {
    buffer: ArrayBuffer;
    extension: string;
  } | null;
}

export interface PackingSpec {
  fullCode: string; // e.g. ZX-X062-020A
  innerBox: {
    materialCode: string; // 物料代码
    length: number;
    width: number;
    height: number;
    capacity: number; // 用量推导的容量，默认 1
  } | null;
  outerBox: {
    materialCode: string;
    length: number;
    width: number;
    height: number;
    capacity: number; // 用量推导的容量，默认 2
  } | null;
}

export interface PackingInputRow {
  originalIndex: number;
  originalXXCode: string; // XX 编码
  originalProductName: string; // 产品名/名称
  originalOEM: string; // OEM
  originalPrice: string | number | null; // 广州价/价格/单价
  imageData?: { buffer: ArrayBuffer; extension: string } | null;
  
  // UI state
  status: 'matched' | 'no_match' | 'invalid_qty' | 'error';
  statusMsg?: string;
  
  availableSpecs: PackingSpec[];
  selectedSpecFullCode?: string; // 用户选择的完整规格
  quantity: string | number; // 用户输入的总支数
  
  weightPerItem: number | null; // 单支净重 (匹配到的)
}

export interface PackingCalculationResult {
  boxType: 'inner' | 'outer';
  boxesCount: number; // 箱数
  itemsPerBox: number; // 每箱数量
  itemsTotal: number; // 这部分总支数
  
  netWeightPerBox: number; // 单箱净重 (单支净重 * 每箱数量)
  grossWeightPerBox: number; // 单箱毛重 (单箱净重 + 3)
  totalNetWeight: number; // 总净重 (单箱净重 * 箱数)
  totalGrossWeight: number; // 总毛重 (单箱毛重 * 箱数)
  
  cbmPerBox: number;
  totalCBM: number;
  
  areaPerBox: number;
  totalArea: number;
}

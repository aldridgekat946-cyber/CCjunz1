
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


import React, { useState, useEffect } from 'react';
import { ProcessedRow } from '../types';
import { ChevronLeft, ChevronRight, Download, ImageIcon, Sparkles } from 'lucide-react';

interface ResultsTableProps {
  data: ProcessedRow[];
  knownOEs: Set<string>;
  onExport: () => void;
}

const normalize = (s: any): string => {
  if (s === null || s === undefined) return "";
  return String(s).replace(/[^A-Za-z0-9]/g, '').toUpperCase();
};

const ImagePreview: React.FC<{ data?: { buffer: ArrayBuffer; extension: string } | null }> = ({ data }) => {
  const [url, setUrl] = useState<string | null>(null);
  useEffect(() => {
    if (data?.buffer) {
      const blob = new Blob([data.buffer], { type: `image/${data.extension}` });
      const objectUrl = URL.createObjectURL(blob);
      setUrl(objectUrl);
      return () => URL.revokeObjectURL(objectUrl);
    }
  }, [data]);
  if (!url) return <div className="flex items-center justify-center w-24 h-9 bg-slate-100 rounded text-slate-400"><ImageIcon size={16} /></div>;
  return <img src={url} alt="Preview" className="w-24 h-9 object-cover rounded border border-slate-200 shadow-sm" />;
};

const HighlightedCell: React.FC<{ text: string, inputOE: string, knownOEs: Set<string>, colName: string, isSpecialMatch?: boolean }> = ({ text, inputOE, knownOEs, colName, isSpecialMatch }) => {
  if (!text) return null;
  const inputNorm = normalize(inputOE);
  
  // Rule: 44250 and 44200 are interchangeable
  let ruleNorm = inputNorm;
  if (ruleNorm.startsWith('44250')) {
    ruleNorm = '44200' + ruleNorm.substring(5);
  } else if (ruleNorm.startsWith('44200')) {
    ruleNorm = '44250' + ruleNorm.substring(5);
  }

  const tokens = text.split(/([\s\n,;:/|，；、]+)/);

  return (
    <span className="whitespace-normal break-words">
      {tokens.map((token, idx) => {
        if (!token) return null;
        const normToken = normalize(token);
        
        if (colName === 'OEM') {
          // OEM 列高亮
          if (normToken === inputNorm) {
            return <span key={idx} className="text-red-600 font-bold">{token}</span>;
          } else if (isSpecialMatch && normToken === ruleNorm) {
            // 触发特殊规则匹配 (44250/44200)
            return <span key={idx} className="text-purple-600 font-bold">{token}</span>;
          }
        } else if (colName === '通用OE') {
          // 通用 OE 列只显示绿色高亮 (匹配数据库已存)
          if (knownOEs.has(normToken)) {
            return <span key={idx} className="text-emerald-600 font-bold">{token}</span>;
          }
        }
        
        return <span key={idx}>{token}</span>;
      })}
    </span>
  );
};

const ResultsTable: React.FC<ResultsTableProps> = ({ data, knownOEs, onExport }) => {
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 10;
  const totalPages = Math.ceil(data.length / rowsPerPage);
  const startIndex = (currentPage - 1) * rowsPerPage;
  const currentData = data.slice(startIndex, startIndex + rowsPerPage);

  const columns = ['输入 OE', 'XX 编码', '适用车型', '年份', 'OEM', '驱动', '图片', '广州价', '产品名', '车型', '通用OE'];

  if (data.length === 0) return null;

  return (
    <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden animate-fade-in">
      <div className="flex items-center justify-between p-4 border-b border-slate-100 bg-slate-50/50">
        <div>
          <h3 className="text-lg font-bold text-slate-800">匹配结果</h3>
          <p className="text-sm text-slate-500">共 {data.length} 条记录 (OEM列仅红高亮，通用OE列仅绿高亮)</p>
        </div>
        <button 
          onClick={onExport}
          className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium rounded-lg transition-colors"
        >
          <Download size={16} />
          导出 Excel
        </button>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-sm text-left table-fixed">
          <thead className="text-xs text-slate-500 uppercase bg-slate-50 border-b border-slate-100">
            <tr>
              <th className="w-12 px-4 py-3 font-semibold">#</th>
              {columns.map(col => (
                <th key={col} className={`px-4 py-3 font-semibold ${
                  (col === '车型' || col === '通用OE' || col === 'OEM' || col === '产品名' || col === '适用车型') ? 'w-48' : 'w-32'
                }`}>
                  {col}
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {currentData.map((row, idx) => (
              <tr key={idx} className={`hover:bg-slate-50 transition-colors ${row.isSpecialMatch ? 'bg-purple-50/30' : ''}`}>
                <td className="px-4 py-3 text-slate-400 font-mono text-xs">{startIndex + idx + 1}</td>
                {columns.map(col => (
                  <td key={col} className="px-4 py-3 align-middle">
                    {col === '图片' ? (
                      <ImagePreview data={row['图片数据']} />
                    ) : (col === 'OEM' || col === '通用OE') ? (
                      <HighlightedCell text={String(row[col] || "")} inputOE={row['输入 OE']} knownOEs={knownOEs} colName={col} isSpecialMatch={row.isSpecialMatch} />
                    ) : (
                      <span className={`whitespace-normal break-words ${col === '输入 OE' ? 'font-bold' : ''}`}>
                        {col === '产品名' && !row['XX 编码'] && <Sparkles size={12} className="inline mr-1 text-indigo-400" />}
                        {String(row[col] || "-")}
                      </span>
                    )}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="flex items-center justify-between px-6 py-4 border-t border-slate-100">
        <span className="text-xs text-slate-500">第 {currentPage} / {totalPages} 页</span>
        <div className="flex gap-2">
          <button onClick={() => setCurrentPage(p => Math.max(1, p-1))} disabled={currentPage === 1} className="p-1 rounded hover:bg-slate-100 disabled:opacity-30"><ChevronLeft/></button>
          <button onClick={() => setCurrentPage(p => Math.min(totalPages, p+1))} disabled={currentPage === totalPages} className="p-1 rounded hover:bg-slate-100 disabled:opacity-30"><ChevronRight/></button>
        </div>
      </div>
    </div>
  );
};

export default ResultsTable;

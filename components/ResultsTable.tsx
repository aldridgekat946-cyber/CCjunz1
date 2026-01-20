import React, { useState, useEffect } from 'react';
import { ProcessedRow } from '../types';
import { ChevronLeft, ChevronRight, Download, ImageIcon } from 'lucide-react';

interface ResultsTableProps {
  data: ProcessedRow[];
  onExport: () => void;
}

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

  if (!url) return <div className="flex items-center justify-center w-12 h-12 bg-slate-100 rounded text-slate-400"><ImageIcon size={16} /></div>;

  return (
    <img src={url} alt="Preview" className="w-12 h-12 object-cover rounded border border-slate-200 shadow-sm" />
  );
};

const ResultsTable: React.FC<ResultsTableProps> = ({ data, onExport }) => {
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 10;
  
  const totalPages = Math.ceil(data.length / rowsPerPage);
  const startIndex = (currentPage - 1) * rowsPerPage;
  const currentData = data.slice(startIndex, startIndex + rowsPerPage);

  const columns = [
    '输入 OE',
    'XX 编码',
    '适用车型',
    '年份',
    'OEM',
    '驱动',
    '图片',
    '广州价'
  ];

  if (data.length === 0) return null;

  return (
    <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden animate-fade-in">
      <div className="flex items-center justify-between p-4 border-b border-slate-100 bg-slate-50/50">
        <div>
          <h3 className="text-lg font-bold text-slate-800">匹配结果</h3>
          <p className="text-sm text-slate-500">共 {data.length} 条记录 (包含图片提取)</p>
        </div>
        <button 
          onClick={onExport}
          className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium rounded-lg transition-colors shadow-sm"
        >
          <Download size={16} />
          导出包含图片的 Excel
        </button>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-sm text-left">
          <thead className="text-xs text-slate-500 uppercase bg-slate-50 border-b border-slate-100">
            <tr>
              <th className="px-6 py-3 font-semibold">#</th>
              {columns.map(col => (
                <th key={col} className="px-6 py-3 font-semibold whitespace-nowrap">
                  {col}
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {currentData.map((row, idx) => (
              <tr key={startIndex + idx} className="hover:bg-slate-50 transition-colors">
                <td className="px-6 py-3 text-slate-400 font-mono text-xs">{startIndex + idx + 1}</td>
                {columns.map((col) => (
                  <td key={col} className="px-6 py-3 max-w-xs truncate" title={col !== '图片' ? String(row[col] || '') : ''}>
                    {col === '图片' ? (
                      <ImagePreview data={row['图片数据']} />
                    ) : row[col] ? (
                      <span className={
                        col === 'OEM' ? 'font-medium text-red-600' :
                        col === '输入 OE' ? 'font-medium text-slate-900' : 
                        'text-slate-600'
                      }>
                        {String(row[col])}
                      </span>
                    ) : (
                      <span className="text-slate-300 italic">-</span>
                    )}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="flex items-center justify-between px-6 py-4 border-t border-slate-100 bg-slate-50/50">
        <div className="text-xs text-slate-500">
          显示第 {startIndex + 1} 至 {Math.min(startIndex + rowsPerPage, data.length)} 条，共 {data.length} 条
        </div>
        <div className="flex gap-2">
          <button
            onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
            disabled={currentPage === 1}
            className="p-1 rounded hover:bg-slate-200 disabled:opacity-30 disabled:hover:bg-transparent transition-colors"
          >
            <ChevronLeft size={20} />
          </button>
          <span className="flex items-center text-xs font-medium text-slate-600 px-2">
            第 {currentPage} 页 / 共 {totalPages} 页
          </span>
          <button
            onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
            disabled={currentPage === totalPages}
            className="p-1 rounded hover:bg-slate-200 disabled:opacity-30 disabled:hover:bg-transparent transition-colors"
          >
            <ChevronRight size={20} />
          </button>
        </div>
      </div>
    </div>
  );
};

export default ResultsTable;
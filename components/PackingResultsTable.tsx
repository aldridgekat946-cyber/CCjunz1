import React, { useState } from 'react';
import { PackingInputRow, PackingCalculationResult } from '../types';
import { calculatePacking } from '../utils/packingProcessor';
import { Download, AlertCircle, ImageIcon } from 'lucide-react';

interface PackingResultsTableProps {
  data: PackingInputRow[];
  onDataChange: (newData: PackingInputRow[]) => void;
  onExport: () => void;
}

const ImagePreview: React.FC<{ data?: { buffer: ArrayBuffer; extension: string } | null }> = ({ data }) => {
  const [url, setUrl] = React.useState<string | null>(null);
  React.useEffect(() => {
    if (data?.buffer) {
      const blob = new Blob([data.buffer], { type: `image/${data.extension}` });
      const objectUrl = URL.createObjectURL(blob);
      setUrl(objectUrl);
      return () => URL.revokeObjectURL(objectUrl);
    }
  }, [data]);
  if (!url) return <div className="flex items-center justify-center w-16 h-10 bg-slate-100 rounded text-slate-400"><ImageIcon size={14} /></div>;
  return <img src={url} alt="Preview" className="w-16 h-10 object-contain rounded border border-slate-200" />;
};

const validateRow = (row: PackingInputRow) => {
    let statusMsgParts = [];
    if (!row.originalOEM && !row.imageData && row.statusMsg?.includes('缺产品资料')) {
        statusMsgParts.push('缺产品资料');
    } else if (!row.imageData) {
        statusMsgParts.push('缺图片');
    }

    if (row.availableSpecs.length === 0) {
        statusMsgParts.push('缺包装规格');
    } else if (!row.selectedSpecFullCode) {
        statusMsgParts.push('规格未选择');
    } else {
        const spec = row.availableSpecs.find(s => s.fullCode === row.selectedSpecFullCode);
        if (spec) {
            if (!spec.innerBox) statusMsgParts.push('缺内箱');
            if (!spec.outerBox) statusMsgParts.push('缺外箱');
        }
    }

    if (row.weightPerItem === null) {
        statusMsgParts.push('缺重量');
    }

    if (row.quantity !== '') {
        const val = parseInt(String(row.quantity), 10);
        if (isNaN(val) || val <= 0 || !/^\d+$/.test(String(row.quantity))) {
            statusMsgParts.push('数量无效');
        }
    }

    const blockingErrors = ['缺包装规格', '规格未选择', '缺外箱', '缺重量', '数量无效'];
    const hasBlockingError = statusMsgParts.some(m => blockingErrors.includes(m));
    if (hasBlockingError) {
        row.status = 'error';
    } else if (row.quantity === '') {
        row.status = 'no_match';
    } else {
        row.status = 'matched';
    }
    
    row.statusMsg = statusMsgParts.join(', ');
};

export const PackingResultsTable: React.FC<PackingResultsTableProps> = ({ data, onDataChange, onExport }) => {
  const [filter, setFilter] = useState<'all' | 'error' | 'matched'>('all');

  const handleSpecChange = (index: number, fullCode: string) => {
    const newData = [...data];
    newData[index].selectedSpecFullCode = fullCode;
    validateRow(newData[index]);
    onDataChange(newData);
  };

  const handleQuantityChange = (index: number, qty: string) => {
    const newData = [...data];
    newData[index].quantity = qty;
    validateRow(newData[index]);
    onDataChange(newData);
  };

  const displayData = data.filter(r => {
      if (filter === 'all') return true;
      if (filter === 'error') return r.status === 'error' || r.status === 'no_match' || r.status === 'invalid_qty';
      if (filter === 'matched') return r.status === 'matched';
      return true;
  });

  // Calculate totals and stats
  let totalItems = 0;
  let totalBoxes = 0;
  let totalNet = 0;
  let totalGross = 0;
  let totalCbm = 0;
  let validRowsCount = 0;
  
  let statInput = data.length;
  let statProd = 0;
  let statImg = 0;
  let statPack = 0;
  let statWeight = 0;

  data.forEach(r => {
      if (r.originalOEM || r.imageData || !r.statusMsg?.includes('缺产品资料')) statProd++;
      if (r.imageData) statImg++;
      if (r.availableSpecs.length > 0) statPack++;
      if (r.weightPerItem !== null) statWeight++;

      if (r.status === 'matched' && r.selectedSpecFullCode && r.quantity) {
          const calcs = calculatePacking(r);
          if (calcs.length > 0) {
              validRowsCount++;
              calcs.forEach(c => {
                  totalItems += c.itemsTotal;
                  totalBoxes += c.boxesCount;
                  totalNet += c.totalNetWeight;
                  totalGross += c.totalGrossWeight;
                  totalCbm += c.totalCBM;
              });
          }
      }
  });

  return (
    <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden flex flex-col">
      {/* Top summary card */}
      <div className="grid grid-cols-2 md:grid-cols-6 gap-4 p-4 bg-slate-50 border-b border-slate-200">
          <div className="flex flex-col"><span className="text-xs text-slate-500">有效参与计算行</span><span className="font-bold text-lg">{validRowsCount}</span></div>
          <div className="flex flex-col"><span className="text-xs text-slate-500">总支数</span><span className="font-bold text-lg">{totalItems}</span></div>
          <div className="flex flex-col"><span className="text-xs text-slate-500">总箱数</span><span className="font-bold text-lg text-indigo-600">{totalBoxes}</span></div>
          <div className="flex flex-col"><span className="text-xs text-slate-500">总净重(kg)</span><span className="font-bold text-lg">{totalNet.toFixed(2)}</span></div>
          <div className="flex flex-col"><span className="text-xs text-slate-500">总毛重(kg)</span><span className="font-bold text-lg text-emerald-600">{totalGross.toFixed(2)}</span></div>
          <div className="flex flex-col"><span className="text-xs text-slate-500">总体积(CBM)</span><span className="font-bold text-lg">{totalCbm.toFixed(2)}</span></div>
      </div>
      
      {/* Missing stats row */}
      <div className="flex flex-wrap gap-4 px-4 py-3 bg-slate-100 border-b border-slate-200 text-sm font-medium">
          <div className="text-slate-600">匹配统计：</div>
          <div className="text-slate-700">输入行数: {statInput}</div>
          <div className={statProd < statInput ? "text-amber-600" : "text-emerald-600"}>产品匹配: {statProd}</div>
          <div className={statImg < statInput ? "text-amber-600" : "text-emerald-600"}>图片匹配: {statImg}</div>
          <div className={statPack < statInput ? "text-red-600" : "text-emerald-600"}>包装匹配: {statPack}</div>
          <div className={statWeight < statInput ? "text-red-600" : "text-emerald-600"}>重量匹配: {statWeight}</div>
      </div>
      
      {/* Controls */}
      <div className="flex items-center justify-between p-4 border-b border-slate-100">
          <div className="flex gap-2">
              <select 
                  className="text-sm border border-slate-200 rounded px-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-indigo-500/20"
                  value={filter}
                  onChange={(e) => setFilter(e.target.value as any)}
              >
                  <option value="all">全部行</option>
                  <option value="matched">仅看有效行</option>
                  <option value="error">仅看异常行</option>
              </select>
          </div>
          <button 
              onClick={onExport}
              className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium rounded-lg transition-colors"
          >
              <Download size={16} />导出装箱单
          </button>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-sm text-left whitespace-nowrap table-auto min-w-[1200px]">
          <thead className="text-xs text-slate-500 uppercase bg-slate-50 border-b border-slate-100 sticky top-0 z-10 shadow-sm">
            <tr>
              <th className="px-3 py-3 font-semibold sticky left-0 bg-slate-50 z-20">#</th>
              <th className="px-3 py-3 font-semibold sticky left-12 bg-slate-50 z-20">XX CODE / 图片</th>
              <th className="px-3 py-3 font-semibold sticky left-48 bg-slate-50 z-20">状态</th>
              <th className="px-3 py-3 font-semibold min-w-[200px]">规格选择</th>
              <th className="px-3 py-3 font-semibold">总数量(支)</th>
              <th className="px-3 py-3 font-semibold">计算结果</th>
              <th className="px-3 py-3 font-semibold">外箱尺寸</th>
              <th className="px-3 py-3 font-semibold">内箱尺寸</th>
              <th className="px-3 py-3 font-semibold">单支净重</th>
              <th className="px-3 py-3 font-semibold">OEM / 产品名</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {displayData.map((row, idx) => {
                const calcs = calculatePacking(row);
                
                return (
              <tr key={row.originalIndex} className="hover:bg-slate-50/50">
                <td className="px-3 py-3 text-slate-400 sticky left-0 bg-white group-hover:bg-slate-50 z-10">{row.originalIndex}</td>
                <td className="px-3 py-3 font-bold text-slate-700 sticky left-12 bg-white group-hover:bg-slate-50 z-10">
                    <div className="flex flex-col gap-1">
                        <span>{row.originalXXCode}</span>
                        <ImagePreview data={row.imageData} />
                    </div>
                </td>
                <td className="px-3 py-3 sticky left-48 bg-white group-hover:bg-slate-50 z-10">
                    {row.status === 'matched' ? (
                        <span className="px-2 py-1 bg-green-100 text-green-700 rounded text-xs font-medium border border-green-200">准备就绪</span>
                    ) : (
                        <span className="flex items-center gap-1 px-2 py-1 bg-red-100 text-red-700 rounded text-xs font-medium border border-red-200">
                            <AlertCircle size={12}/>{row.statusMsg || '异常'}
                        </span>
                    )}
                </td>
                <td className="px-3 py-3">
                    <select 
                        className={`w-full max-w-[220px] text-sm border rounded p-1.5 focus:outline-none ${row.availableSpecs.length === 0 ? 'bg-slate-100 text-slate-400' : 'bg-white border-slate-300'}`}
                        value={row.selectedSpecFullCode || ''}
                        onChange={(e) => handleSpecChange(data.findIndex(r => r.originalIndex === row.originalIndex), e.target.value)}
                        disabled={row.availableSpecs.length === 0}
                    >
                        {row.availableSpecs.length !== 1 && <option value="">请选择规格...</option>}
                        {row.availableSpecs.map(s => (
                            <option key={s.fullCode} value={s.fullCode}>{s.fullCode}</option>
                        ))}
                    </select>
                </td>
                <td className="px-3 py-3">
                    <input 
                        type="text"
                        value={row.quantity}
                        onChange={(e) => handleQuantityChange(data.findIndex(r => r.originalIndex === row.originalIndex), e.target.value)}
                        placeholder="输入数量"
                        className={`w-24 px-2 py-1.5 border rounded text-sm focus:outline-none focus:ring-2 ${row.status === 'invalid_qty' ? 'border-red-400 ring-red-100' : 'border-slate-300 focus:ring-indigo-100'}`}
                    />
                </td>
                
                <td className="px-3 py-3 text-xs">
                    {calcs.length > 0 ? (
                        <div className="flex flex-col gap-1">
                            {calcs.map((c, i) => (
                                <div key={i} className="flex gap-2 items-center">
                                    <span className={`px-1.5 py-0.5 rounded ${c.boxType === 'outer' ? 'bg-blue-100 text-blue-700' : 'bg-amber-100 text-amber-700'}`}>
                                        {c.boxType === 'outer' ? '外箱' : '内箱'}
                                    </span>
                                    <span>{c.boxesCount} 箱 × {c.itemsPerBox}支 = {c.itemsTotal}支</span>
                                </div>
                            ))}
                        </div>
                    ) : (
                        <span className="text-slate-400">-</span>
                    )}
                </td>

                <td className="px-3 py-3 text-xs text-slate-600">
                    {row.selectedSpecFullCode ? (() => {
                        const s = row.availableSpecs.find(x => x.fullCode === row.selectedSpecFullCode);
                        if (!s || !s.outerBox) return '-';
                        return <div>
                            <div className="font-mono">{s.outerBox.materialCode}</div>
                            <div>{s.outerBox.length}*{s.outerBox.width}*{s.outerBox.height}</div>
                        </div>;
                    })() : '-'}
                </td>
                <td className="px-3 py-3 text-xs text-slate-600">
                    {row.selectedSpecFullCode ? (() => {
                        const s = row.availableSpecs.find(x => x.fullCode === row.selectedSpecFullCode);
                        if (!s || !s.innerBox) return '无内箱';
                        return <div>
                            <div className="font-mono">{s.innerBox.materialCode}</div>
                            <div>{s.innerBox.length}*{s.innerBox.width}*{s.innerBox.height}</div>
                        </div>;
                    })() : '-'}
                </td>
                
                <td className="px-3 py-3 font-mono">
                    {row.weightPerItem !== null ? `${row.weightPerItem} kg` : <span className="text-red-500">缺失</span>}
                </td>

                <td className="px-3 py-3 text-xs max-w-[200px] truncate" title={row.originalOEM}>
                    <div className="font-bold text-slate-700 truncate">{row.originalProductName || '-'}</div>
                    <div className="text-slate-500 truncate">{row.originalOEM || '-'}</div>
                </td>

              </tr>
            )})}
          </tbody>
        </table>
      </div>
    </div>
  );
};

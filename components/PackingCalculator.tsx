import React, { useState } from 'react';
import { Settings, Play, RefreshCw, AlertCircle, Sparkles, Info, ArrowRight, Box } from 'lucide-react';
import FileUploader from './FileUploader';
import { PackingResultsTable } from './PackingResultsTable';
import { processPackingReference, processPackingQueries, exportPackingList } from '../utils/packingProcessor';
import { PackingSpec, PackingInputRow } from '../types';

export const PackingCalculator: React.FC = () => {
  const [refFile, setRefFile] = useState<File | null>(null);
  const [oeFile, setOeFile] = useState<File | null>(null);
  
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingMsg, setProcessingMsg] = useState("");
  
  const [refData, setRefData] = useState<{ specs: PackingSpec[], weights: Map<string, number> } | null>(null);
  const [inputData, setInputData] = useState<PackingInputRow[]>([]);
  
  const [error, setError] = useState<string | null>(null);

  const handleProcess = async () => {
    if (!refFile || !oeFile) return;
    
    setIsProcessing(true);
    setError(null);
    setInputData([]);
    setProcessingMsg("初始化匹配引擎...");

    try {
        let currentRefData = refData;
        if (!currentRefData) {
            setProcessingMsg("正在解析参考库...");
            currentRefData = await processPackingReference(refFile);
            setRefData(currentRefData);
        }

        const queries = await processPackingQueries(oeFile, currentRefData, (msg) => {
            setProcessingMsg(msg);
        });
        
        setInputData(queries);
    } catch (err: any) {
        setError(err.message || "处理过程发生异常，请检查文件格式。");
    } finally {
        setIsProcessing(false);
        setProcessingMsg("");
    }
  };

  const handleClearRef = () => {
      setRefFile(null);
      setRefData(null);
  };

  const handleExport = () => {
      exportPackingList(inputData, `装箱单_${new Date().toISOString().slice(0,10).replace(/-/g, '')}_${new Date().toTimeString().slice(0,5).replace(':', '')}.xlsx`);
  };

  return (
    <div className="w-full">
        {/* Uploaders */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8 mb-12 max-w-4xl mx-auto animate-in fade-in slide-in-from-bottom-8 duration-700 delay-200">
          <div className="space-y-3">
            <div className="flex items-center gap-2 px-1">
              <Info size={14} className="text-indigo-400" />
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">Step 1</span>
            </div>
            <FileUploader 
              label="上传参考库" 
              subLabel="包含内外纸箱、重量两个Sheet"
              file={refFile} 
              onFileSelect={(f) => { setRefFile(f); setRefData(null); }} 
              onClear={handleClearRef} 
              color="blue" 
            />
          </div>
          <div className="space-y-3">
            <div className="flex items-center gap-2 px-1">
              <Info size={14} className="text-purple-400" />
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">Step 2</span>
            </div>
            <FileUploader 
              label="上传待查清单" 
              subLabel="包含 XX CODE 及数量需要的明细"
              file={oeFile} 
              onFileSelect={setOeFile} 
              onClear={() => setOeFile(null)} 
              color="purple" 
            />
          </div>
        </div>

        {/* Action Button */}
        <div className="flex flex-col items-center gap-6 animate-in fade-in slide-in-from-bottom-12 duration-700 delay-300">
          {error && (
            <div className="flex items-center gap-3 px-6 py-3 bg-red-50 border border-red-100 text-red-600 rounded-2xl text-sm font-medium shadow-sm max-w-2xl text-center">
              <AlertCircle size={18} className="shrink-0" />
              {error}
            </div>
          )}
          
          {isProcessing && (
            <div className="flex flex-col items-center gap-3">
              <div className="relative">
                <div className="w-12 h-12 border-4 border-indigo-100 border-t-indigo-600 rounded-full animate-spin"></div>
                <div className="absolute inset-0 flex items-center justify-center">
                  <Box size={16} className="text-indigo-400 animate-pulse" />
                </div>
              </div>
              <div className="text-indigo-600 font-bold text-sm tracking-tight flex items-center gap-2">
                <RefreshCw className="animate-spin-slow" size={16}/>
                {processingMsg}
              </div>
            </div>
          )}
          
          {!isProcessing && (
            <button 
              onClick={handleProcess} 
              disabled={!refFile || !oeFile}
              className={`
                relative px-12 py-4 rounded-2xl font-black text-lg transition-all duration-300
                shadow-2xl hover:shadow-indigo-300/50 hover:-translate-y-1 active:scale-95
                ${(!refFile || !oeFile) 
                  ? "bg-slate-200 text-slate-400 cursor-not-allowed shadow-none" 
                  : "bg-indigo-600 text-white hover:bg-indigo-700 active:bg-indigo-800"
                }
              `}
            >
              <span className="flex items-center gap-3">
                <Play size={20} fill="currentColor" />
                开始解析与匹配
              </span>
            </button>
          )}
        </div>

        {/* Results Section */}
        {inputData.length > 0 && (
          <div className="mt-16 animate-in fade-in zoom-in-95 duration-500">
              <PackingResultsTable 
                  data={inputData}
                  onDataChange={setInputData}
                  onExport={handleExport}
              />
          </div>
        )}
    </div>
  );
};

import React, { useState } from 'react';
import { Settings, Play, RefreshCw, AlertCircle } from 'lucide-react';
import FileUploader from './components/FileUploader';
import ResultsTable from './components/ResultsTable';
import { processFiles, exportToExcel } from './utils/excelProcessor';
import { ProcessedRow } from './types';

const App: React.FC = () => {
  const [refFile, setRefFile] = useState<File | null>(null);
  const [oeFile, setOeFile] = useState<File | null>(null);
  
  const [isProcessing, setIsProcessing] = useState(false);
  const [results, setResults] = useState<ProcessedRow[]>([]);
  const [error, setError] = useState<string | null>(null);

  const handleProcess = async () => {
    if (!refFile || !oeFile) return;
    
    setIsProcessing(true);
    setError(null);
    setResults([]);

    try {
      // Small delay to allow UI to update
      setTimeout(async () => {
        try {
          const processedData = await processFiles(refFile, oeFile);
          setResults(processedData);
        } catch (err: any) {
          setError(err.message || "处理文件时发生错误。");
          console.error(err);
        } finally {
          setIsProcessing(false);
        }
      }, 100);
    } catch (err) {
      setError("无法启动处理流程。");
      setIsProcessing(false);
    }
  };

  const handleReset = () => {
    setRefFile(null);
    setOeFile(null);
    setResults([]);
    setError(null);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 pb-20">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-indigo-600 rounded-lg flex items-center justify-center text-white">
              <Settings size={20} />
            </div>
            <div>
              <h1 className="text-xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-indigo-600 to-purple-600">
                PartMatch Pro
              </h1>
            </div>
          </div>
          <div className="text-xs text-slate-500 font-medium">
             v1.1.0
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8">
        
        <div className="mb-10 text-center max-w-2xl mx-auto">
          <h2 className="text-3xl font-bold text-slate-800 mb-3">自动配件匹配系统</h2>
          <p className="text-slate-600">
            请上传您的参考数据库（包含 OEM 数据和价格）以及需要查询的 OE 清单。
          </p>
        </div>

        {/* Upload Section */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8 max-w-4xl mx-auto">
          <FileUploader 
            label="1. 参考数据库 (Reference DB)"
            subLabel="包含 OEM, XX CODE, 详细信息及价格"
            file={refFile} 
            onFileSelect={setRefFile} 
            onClear={() => setRefFile(null)}
            color="blue"
          />
          <FileUploader 
            label="2. OE 清单 (OE List)"
            subLabel="需要查询的 OE 列表"
            file={oeFile} 
            onFileSelect={setOeFile} 
            onClear={() => setOeFile(null)}
            color="purple"
          />
        </div>

        {/* Actions */}
        <div className="flex flex-col items-center justify-center mb-10 gap-4">
          {error && (
            <div className="flex items-center gap-2 text-red-600 bg-red-50 px-4 py-2 rounded-lg border border-red-100 text-sm animate-fade-in">
              <AlertCircle size={16} />
              {error}
            </div>
          )}

          <div className="flex gap-4">
            <button
              onClick={handleProcess}
              disabled={!refFile || !oeFile || isProcessing}
              className={`
                flex items-center gap-2 px-8 py-3 rounded-full font-semibold text-white shadow-lg shadow-indigo-200
                transition-all duration-200 transform hover:scale-105 active:scale-95
                ${(!refFile || !oeFile) 
                  ? 'bg-slate-300 cursor-not-allowed shadow-none' 
                  : isProcessing 
                    ? 'bg-indigo-400 cursor-wait' 
                    : 'bg-indigo-600 hover:bg-indigo-700'
                }
              `}
            >
              {isProcessing ? (
                <>
                  <RefreshCw className="animate-spin" size={20} />
                  处理中...
                </>
              ) : (
                <>
                  <Play size={20} fill="currentColor" />
                  开始匹配
                </>
              )}
            </button>

            {(results.length > 0 || refFile || oeFile) && !isProcessing && (
              <button
                onClick={handleReset}
                className="px-6 py-3 rounded-full font-medium text-slate-600 hover:bg-slate-100 transition-colors"
              >
                重置
              </button>
            )}
          </div>
        </div>

        {/* Results */}
        {results.length > 0 && (
           <ResultsTable 
             data={results} 
             onExport={() => exportToExcel(results, 'matched_results.xlsx')} 
           />
        )}
      </main>
    </div>
  );
};

export default App;
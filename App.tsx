
import React, { useState } from 'react';
import { Settings, Play, RefreshCw, AlertCircle, Sparkles, Search, CheckCircle2, Info } from 'lucide-react';
import FileUploader from './components/FileUploader';
import ResultsTable from './components/ResultsTable';
import { processFiles, exportToExcel } from './utils/excelProcessor';
import { ProcessedRow } from './types';

const App: React.FC = () => {
  const [refFile, setRefFile] = useState<File | null>(null);
  const [oeFile, setOeFile] = useState<File | null>(null);
  
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingMsg, setProcessingMsg] = useState("");
  const [results, setResults] = useState<ProcessedRow[]>([]);
  const [knownOEs, setKnownOEs] = useState<Set<string>>(new Set());
  const [error, setError] = useState<string | null>(null);

  const handleProcess = async () => {
    if (!refFile || !oeFile) return;
    
    setIsProcessing(true);
    setError(null);
    setResults([]);
    setKnownOEs(new Set());
    setProcessingMsg("初始化匹配引擎...");

    try {
      setTimeout(async () => {
        try {
          const { results: processedData, knownOEs: dbOEs } = await processFiles(refFile, oeFile, (msg) => {
            setProcessingMsg(msg);
          });
          setResults(processedData);
          setKnownOEs(dbOEs);
        } catch (err: any) {
          setError(err.message || "处理过程发生异常，请检查文件格式。");
        } finally {
          setIsProcessing(false);
          setProcessingMsg("");
        }
      }, 100);
    } catch (err) {
      setError("处理流程启动失败");
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 pb-20 selection:bg-indigo-100">
      {/* Header */}
      <header className="bg-white/80 backdrop-blur-md border-b border-slate-200 sticky top-0 z-20 px-6 py-4">
        <div className="max-w-6xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-600 p-2 rounded-xl shadow-lg shadow-indigo-200">
              <Settings className="text-white" size={20} />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight text-slate-900">PartMatch <span className="text-indigo-600">Pro</span></h1>
              <p className="text-[10px] text-slate-500 font-medium uppercase tracking-widest">Automated Automotive Solutions</p>
            </div>
          </div>
          <div className="hidden sm:flex items-center gap-4 text-xs font-medium">
            <div className="flex items-center gap-1.5 px-3 py-1.5 bg-indigo-50 text-indigo-700 rounded-full border border-indigo-100">
              <Sparkles size={14} />
              <span>Google Search Integration</span>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-12">
        {/* Hero / Description Section */}
        <section className="mb-12 text-center max-w-3xl mx-auto animate-in fade-in slide-in-from-bottom-4 duration-700">
          <h2 className="text-4xl font-extrabold text-slate-900 mb-6 tracking-tight">
            自动配件匹配系统 <span className="text-indigo-600">(Pro)</span>
          </h2>
          <div className="bg-white rounded-3xl p-8 border border-slate-200 shadow-xl shadow-slate-200/50 relative overflow-hidden group">
            <div className="absolute top-0 right-0 w-32 h-32 bg-indigo-50 rounded-full -mr-16 -mt-16 transition-transform group-hover:scale-110 duration-500" />
            
            <p className="text-lg text-slate-600 leading-relaxed mb-6 relative z-10">
              上传参考库与待查 OE。库内未命中项将自动通过 <span className="text-indigo-600 font-bold decoration-indigo-300 decoration-2 underline-offset-4 underline">Google 搜索</span> 检索，
              并自动 <span className="text-emerald-600 font-bold">高亮显示</span> 命中库内的通用 OE 编号。
            </p>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 text-left relative z-10">
              <div className="flex items-start gap-3 p-3 rounded-2xl hover:bg-slate-50 transition-colors">
                <div className="bg-blue-100 p-2 rounded-lg text-blue-600 mt-0.5">
                  <Search size={18} />
                </div>
                <div>
                  <h4 className="font-bold text-sm text-slate-800">智能补全</h4>
                  <p className="text-xs text-slate-500">自动补齐缺失车型数据</p>
                </div>
              </div>
              <div className="flex items-start gap-3 p-3 rounded-2xl hover:bg-slate-50 transition-colors">
                <div className="bg-emerald-100 p-2 rounded-lg text-emerald-600 mt-0.5">
                  <CheckCircle2 size={18} />
                </div>
                <div>
                  <h4 className="font-bold text-sm text-slate-800">库存预判</h4>
                  <p className="text-xs text-slate-500">高亮显示已入库 OE</p>
                </div>
              </div>
              <div className="flex items-start gap-3 p-3 rounded-2xl hover:bg-slate-50 transition-colors">
                <div className="bg-purple-100 p-2 rounded-lg text-purple-600 mt-0.5">
                  <Sparkles size={18} />
                </div>
                <div>
                  <h4 className="font-bold text-sm text-slate-800">实时检索</h4>
                  <p className="text-xs text-slate-500">整合全球最新配件信息</p>
                </div>
              </div>
            </div>
          </div>
        </section>

        {/* Uploaders */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8 mb-12 max-w-4xl mx-auto animate-in fade-in slide-in-from-bottom-8 duration-700 delay-200">
          <div className="space-y-3">
            <div className="flex items-center gap-2 px-1">
              <Info size={14} className="text-indigo-400" />
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">Step 1</span>
            </div>
            <FileUploader 
              label="上传参考库" 
              subLabel="包含已有的 OEM 及 价格信息"
              file={refFile} 
              onFileSelect={setRefFile} 
              onClear={() => setRefFile(null)} 
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
              subLabel="需要匹配及搜索的 OE 列表"
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
            <div className="flex items-center gap-3 px-6 py-3 bg-red-50 border border-red-100 text-red-600 rounded-2xl text-sm font-medium animate-bounce shadow-sm">
              <AlertCircle size={18} />
              {error}
            </div>
          )}
          
          {isProcessing && (
            <div className="flex flex-col items-center gap-3">
              <div className="relative">
                <div className="w-12 h-12 border-4 border-indigo-100 border-t-indigo-600 rounded-full animate-spin"></div>
                <div className="absolute inset-0 flex items-center justify-center">
                  <Sparkles size={16} className="text-indigo-400 animate-pulse" />
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
                开始高效匹配
              </span>
            </button>
          )}
        </div>

        {/* Results Section */}
        {results.length > 0 && (
          <div className="mt-16 animate-in fade-in zoom-in-95 duration-500">
            <ResultsTable 
              data={results} 
              knownOEs={knownOEs} 
              onExport={() => exportToExcel(results, `匹配结果_${new Date().toISOString().slice(0,10)}.xlsx`, knownOEs)} 
            />
          </div>
        )}
      </main>

      <footer className="max-w-6xl mx-auto px-6 py-8 border-t border-slate-200 mt-12 text-center text-slate-400 text-xs">
        <p>© 2025 PartMatch Pro - 专家级配件自动化匹配工具 | 集成最新 Gemini AI 引擎</p>
      </footer>
    </div>
  );
};

export default App;

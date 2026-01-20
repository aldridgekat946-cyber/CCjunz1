import React, { useRef } from 'react';
import { Upload, FileSpreadsheet, X, CheckCircle } from 'lucide-react';

interface FileUploaderProps {
  label: string;
  subLabel?: string;
  file: File | null;
  onFileSelect: (file: File) => void;
  onClear: () => void;
  accept?: string;
  color?: "blue" | "green" | "purple";
}

const FileUploader: React.FC<FileUploaderProps> = ({ 
  label, 
  subLabel, 
  file, 
  onFileSelect, 
  onClear, 
  accept = ".csv, .xlsx, .xls",
  color = "blue"
}) => {
  const inputRef = useRef<HTMLInputElement>(null);

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      onFileSelect(e.dataTransfer.files[0]);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      onFileSelect(e.target.files[0]);
    }
  };

  const colorClasses = {
    blue: "border-blue-200 bg-blue-50/50 hover:bg-blue-50 text-blue-600",
    green: "border-green-200 bg-green-50/50 hover:bg-green-50 text-green-600",
    purple: "border-purple-200 bg-purple-50/50 hover:bg-purple-50 text-purple-600",
  };

  const iconColors = {
    blue: "text-blue-500",
    green: "text-green-500",
    purple: "text-purple-500",
  };

  return (
    <div className="flex flex-col gap-2 w-full">
      <label className="text-sm font-semibold text-slate-700">{label}</label>
      
      {!file ? (
        <div 
          onClick={() => inputRef.current?.click()}
          onDragOver={(e) => e.preventDefault()}
          onDrop={handleDrop}
          className={`
            relative cursor-pointer group
            flex flex-col items-center justify-center 
            h-32 rounded-xl border-2 border-dashed 
            transition-all duration-200 ease-in-out
            ${colorClasses[color]}
          `}
        >
          <input 
            type="file" 
            ref={inputRef} 
            className="hidden" 
            accept={accept} 
            onChange={handleChange}
          />
          <div className={`p-3 rounded-full bg-white shadow-sm mb-2 group-hover:scale-110 transition-transform ${iconColors[color]}`}>
            <Upload size={20} />
          </div>
          <p className="text-sm font-medium">点击上传或拖拽文件至此</p>
          {subLabel && <p className="text-xs text-slate-500 mt-1">{subLabel}</p>}
        </div>
      ) : (
        <div className="relative flex items-center p-4 bg-white border border-slate-200 rounded-xl shadow-sm">
          <div className={`p-2 rounded-lg mr-3 bg-opacity-10 ${iconColors[color].replace('text-', 'bg-')}`}>
            <FileSpreadsheet className={iconColors[color]} size={24} />
          </div>
          <div className="flex-1 min-w-0">
            <p className="text-sm font-medium text-slate-900 truncate">{file.name}</p>
            <p className="text-xs text-slate-500">{(file.size / 1024).toFixed(1)} KB</p>
          </div>
          <div className="flex items-center gap-2">
            <CheckCircle size={18} className="text-emerald-500" />
            <button 
              onClick={onClear}
              className="p-1.5 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-lg transition-colors"
            >
              <X size={18} />
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default FileUploader;
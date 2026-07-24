import React, { useState, useRef, useEffect } from "react";
import {
  FileSpreadsheet,
  Upload,
  Trash2,
  RefreshCw,
  Download,
  Sparkles,
  AlertCircle,
  AlertTriangle,
  FileText,
  CheckCircle2,
  Plus,
  Info,
  Wand2,
  Layers,
  ShieldCheck,
  ChevronRight,
  Settings2,
  FilePlus,
  Check,
  Star,
  ArrowUpRight,
  MoveVertical,
  GripVertical,
  Sun,
  Moon,
  LayoutGrid,
  List
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";

interface PDFItem {
  id: string;
  file: File;
  name: string;
  base64: string;
  tabName: string;
  isPrincipal: boolean;
  size: number;
}

export default function App() {
  const [theme, setTheme] = useState<'dark' | 'light'>(() => {
    const saved = localStorage.getItem("table_magic_theme");
    return saved === "light" ? "light" : "dark";
  });

  useEffect(() => {
    localStorage.setItem("table_magic_theme", theme);
    if (theme === "dark") {
      document.documentElement.classList.add("dark");
      document.documentElement.classList.remove("light");
    } else {
      document.documentElement.classList.add("light");
      document.documentElement.classList.remove("dark");
    }
  }, [theme]);

  const toggleTheme = () => {
    setTheme((prev) => (prev === "dark" ? "light" : "dark"));
  };

  const [files, setFiles] = useState<PDFItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [loadingStep, setLoadingStep] = useState("");
  const [result, setResult] = useState<{
    base64: string;
    filename: string;
    summary: Array<{
      fileName: string;
      tabName: string;
      tableCount: number;
      rowCount: number;
      extractionMethod: string;
      warning?: string | null;
    }>;
  } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [hideResponsavel, setHideResponsavel] = useState(false);
  const [hideSecondaryTabs, setHideSecondaryTabs] = useState(false);
  const [disablePrincipalTab, setDisablePrincipalTab] = useState(false);
  const [fileViewMode, setFileViewMode] = useState<'grid' | 'list'>('grid');
  const [isDragOver, setIsDragOver] = useState(false);
  const [draggedFileIndex, setDraggedFileIndex] = useState<number | null>(null);
  const [dragOverIndex, setDragOverIndex] = useState<number | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Process dropped/selected files
  const processFiles = (rawFiles: FileList | File[]) => {
    const pdfFiles = Array.from(rawFiles).filter(f => f.type === "application/pdf" || f.name.toLowerCase().endsWith('.pdf'));
    
    if (pdfFiles.length === 0) {
      setError("Por favor, selecione ou arraste apenas arquivos no formato PDF.");
      return;
    }

    setError(null);

    pdfFiles.forEach((file) => {
      const reader = new FileReader();
      reader.onload = () => {
        const base64String = (reader.result as string).split(",")[1];
        const proposedTabName = file.name
          .replace(/\.pdf$/i, "")
          .replace(/[^a-zA-Z0-9_\-\s]/g, "")
          .trim()
          .substring(0, 20);

        setFiles((prev) => {
          // Check if file with same name already added
          if (prev.some(f => f.name === file.name && f.size === file.size)) {
            return prev;
          }

          const isFirst = prev.length === 0;
          const newItem: PDFItem = {
            id: `file-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
            file,
            name: file.name,
            base64: base64String,
            tabName: proposedTabName || (isFirst ? "Principal" : `Aba ${prev.length + 1}`),
            isPrincipal: isFirst,
            size: file.size,
          };
          return [...prev, newItem];
        });
      };
      reader.onerror = () => {
        setError(`Erro ao carregar o arquivo ${file.name}.`);
      };
      reader.readAsDataURL(file);
    });
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      processFiles(e.dataTransfer.files);
    }
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      processFiles(e.target.files);
    }
  };

  // Reordering & Principal promotion via drag & drop
  const handleCardDragStart = (e: React.DragEvent, index: number) => {
    e.stopPropagation();
    setDraggedFileIndex(index);
    e.dataTransfer.effectAllowed = "move";
    e.dataTransfer.setData("text/plain", index.toString());
  };

  const handleCardDragOver = (e: React.DragEvent, index: number) => {
    e.preventDefault();
    e.stopPropagation();
    if (dragOverIndex !== index) {
      setDragOverIndex(index);
    }
  };

  const handleCardDrop = (e: React.DragEvent, targetIndex: number) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOverIndex(null);

    if (draggedFileIndex === null || draggedFileIndex === targetIndex) return;

    setFiles((prev) => {
      const updated = [...prev];
      const [movedItem] = updated.splice(draggedFileIndex, 1);
      updated.splice(targetIndex, 0, movedItem);

      // The top item at index 0 is always marked as Principal, all others false
      return updated.map((item, idx) => ({
        ...item,
        isPrincipal: idx === 0,
      }));
    });

    setDraggedFileIndex(null);
  };

  const setAsPrincipal = (id: string) => {
    setFiles((prev) => {
      const targetIdx = prev.findIndex((item) => item.id === id);
      if (targetIdx <= 0) return prev;
      const updated = [...prev];
      const [promoted] = updated.splice(targetIdx, 1);
      updated.unshift(promoted);
      return updated.map((item, idx) => ({
        ...item,
        isPrincipal: idx === 0,
      }));
    });
  };

  const removeFile = (id: string) => {
    setFiles((prev) => {
      const updated = prev.filter((item) => item.id !== id);
      if (updated.length > 0 && !updated.some((item) => item.isPrincipal)) {
        updated[0].isPrincipal = true;
      }
      return updated.map((item, idx) => ({
        ...item,
        isPrincipal: idx === 0,
      }));
    });
  };

  const updateTabName = (id: string, name: string) => {
    setFiles((prev) =>
      prev.map((item) => (item.id === id ? { ...item, tabName: name } : item))
    );
  };

  const clearAllFiles = () => {
    setFiles([]);
    setResult(null);
    setError(null);
  };


  // Convert and Join
  const handleExecute = async () => {
    setError(null);
    setResult(null);

    if (files.length === 0) {
      setError("Por favor, faça o upload de pelo menos um arquivo PDF antes de executar.");
      return;
    }

    const principal = files.find((f) => f.isPrincipal) || files[0];
    const secondaries = files.filter((f) => f.id !== principal.id);

    const filesToSubmit = [
      {
        name: principal.name,
        tabName: principal.tabName || "Principal",
        base64: principal.base64,
        isPrincipal: true,
      },
      ...secondaries.map((item) => ({
        name: item.name,
        tabName: item.tabName || item.name.replace(/\.pdf$/i, "").substring(0, 20),
        base64: item.base64,
        isPrincipal: false,
      })),
    ];

    try {
      setLoading(true);
      setLoadingStep("Conectando com o servidor de extração...");

      const steps = [
        "Lendo e analisando a estrutura do PDF Principal...",
        "Detectando tabelas, limites e cabeçalhos...",
        secondaries.length > 0 ? "Processando PDFs secundários e unificando colunas..." : "Formatando células e tipos de dados...",
        "Estruturando abas dinâmicas no Excel...",
        "Finalizando arquivo .XLSX de alta precisão..."
      ];

      let stepIndex = 0;
      const stepInterval = setInterval(() => {
        if (stepIndex < steps.length) {
          setLoadingStep(steps[stepIndex]);
          stepIndex++;
        }
      }, 2200);

      const response = await fetch("/api/convert", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ files: filesToSubmit, hideResponsavel, hideSecondaryTabs, disablePrincipalTab }),
      });

      clearInterval(stepInterval);

      if (!response.ok) {
        const errData = await response.json();
        throw new Error(errData.error || "Ocorreu um erro no servidor durante a conversão.");
      }

      const data = await response.json();
      if (data.success) {
        setResult({
          base64: data.base64,
          filename: data.filename || "Table_Magic_Combined.xlsx",
          summary: data.summary,
        });
      } else {
        throw new Error("Resposta inesperada do servidor.");
      }
    } catch (err: any) {
      console.error(err);
      setError(err.message || "Houve uma falha no processamento dos arquivos PDF.");
    } finally {
      setLoading(false);
      setLoadingStep("");
    }
  };

  const triggerDownload = () => {
    if (!result) return;
    const byteCharacters = atob(result.base64);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = result.filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const formatFileSize = (bytes: number) => {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
  };

  const principalItem = files.find((f) => f.isPrincipal) || files[0];

  return (
    <div className={`min-h-screen font-sans selection:bg-emerald-500/30 selection:text-emerald-200 relative overflow-x-hidden pb-16 transition-colors duration-300 ${
      theme === "dark" ? "bg-[#0b0f17] text-slate-100" : "bg-slate-50 text-slate-800"
    }`}>
      
      {/* Background Ambience & Soft Grids */}
      <div className={`absolute top-0 left-1/2 -translate-x-1/2 w-full max-w-7xl h-[500px] pointer-events-none ${
        theme === "dark" 
          ? "bg-[radial-gradient(ellipse_60%_50%_at_50%_0%,rgba(16,185,129,0.14),rgba(255,255,255,0))]" 
          : "bg-[radial-gradient(ellipse_60%_50%_at_50%_0%,rgba(16,185,129,0.08),rgba(255,255,255,0))]"
      }`} />
      <div className="absolute inset-0 bg-[linear-gradient(to_right,#1f293d0f_1px,transparent_1px),linear-gradient(to_bottom,#1f293d0f_1px,transparent_1px)] bg-[size:4rem_4rem] [mask-image:radial-gradient(ellipse_60%_50%_at_50%_0%,#000_70%,transparent_100%)] pointer-events-none" />

      {/* Top Navbar */}
      <header className={`sticky top-0 z-40 border-b backdrop-blur-md px-6 py-4 transition-colors duration-300 ${
        theme === "dark" ? "border-slate-800/80 bg-[#0b0f17]/80" : "border-slate-200/80 bg-white/80"
      }`}>
        <div className="max-w-5xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-emerald-500 to-amber-400 p-[1px] shadow-lg shadow-emerald-500/20">
              <div className={`w-full h-full rounded-[11px] flex items-center justify-center ${theme === 'dark' ? 'bg-[#0d131f]' : 'bg-white'}`}>
                <Wand2 className="w-5 h-5 text-emerald-500" />
              </div>
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className={`font-display font-bold text-xl tracking-tight ${theme === 'dark' ? 'text-white' : 'text-slate-900'}`}>
                  TABLE <span className="bg-gradient-to-r from-emerald-500 to-amber-400 bg-clip-text text-transparent">MAGIC</span>
                </h1>
                <span className={`text-[10px] font-mono font-semibold px-2 py-0.5 rounded-full border ${
                  theme === 'dark' ? 'bg-emerald-500/10 text-emerald-400 border-emerald-500/20' : 'bg-emerald-50 text-emerald-700 border-emerald-200'
                }`}>
                  v2.5 Studio
                </span>
              </div>
              <p className={`text-xs ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>PDF Table Extractor & Excel Sheet Joiner</p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            <div className={`hidden sm:flex items-center gap-2 px-3 py-1.5 rounded-full border text-xs ${
              theme === 'dark' ? 'bg-slate-900/80 border-slate-800 text-slate-300' : 'bg-slate-100 border-slate-200 text-slate-700'
            }`}>
              <ShieldCheck className="w-4 h-4 text-emerald-500" />
              <span>Processamento 100% Local</span>
            </div>

            {/* Theme Toggle Button */}
            <button
              type="button"
              onClick={toggleTheme}
              className={`p-2 px-3 rounded-full border flex items-center gap-2 transition-all cursor-pointer text-xs font-semibold select-none ${
                theme === "dark"
                  ? "bg-slate-900 border-slate-800 text-amber-400 hover:bg-slate-800 hover:border-slate-700"
                  : "bg-white border-slate-200 text-slate-700 hover:bg-slate-100 hover:border-slate-300 shadow-sm"
              }`}
              title={theme === "dark" ? "Mudar para Modo Claro" : "Mudar para Modo Escuro"}
            >
              {theme === "dark" ? (
                <>
                  <Sun className="w-4 h-4 text-amber-400" />
                  <span className="hidden sm:inline font-mono">Modo Claro</span>
                </>
              ) : (
                <>
                  <Moon className="w-4 h-4 text-slate-700" />
                  <span className="hidden sm:inline font-mono">Modo Escuro</span>
                </>
              )}
            </button>
          </div>
        </div>
      </header>

      {/* Main Studio Container */}
      <main className="max-w-4xl mx-auto px-4 sm:px-6 pt-8 relative z-10">

        {/* Global Error Banner */}
        <AnimatePresence>
          {error && (
            <motion.div
              initial={{ opacity: 0, y: -10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="mb-6 p-4 rounded-xl bg-red-950/40 border border-red-500/30 text-red-300 flex items-center justify-between gap-3 shadow-lg"
            >
              <div className="flex items-center gap-3">
                <AlertCircle className="w-5 h-5 text-red-400 shrink-0" />
                <p className="text-sm font-medium">{error}</p>
              </div>
              <button
                onClick={() => setError(null)}
                className="text-xs text-red-400 hover:text-white underline font-medium"
              >
                Fechar
              </button>
            </motion.div>
          )}
        </AnimatePresence>

        {loading ? (
          /* THEMATIC CONVERSION PIPELINE ANIMATION */
          <motion.div
            initial={{ opacity: 0, scale: 0.96 }}
            animate={{ opacity: 1, scale: 1 }}
            exit={{ opacity: 0, scale: 0.96 }}
            className="obsidian-card rounded-3xl p-8 sm:p-12 border text-center space-y-8 shadow-2xl my-8 relative overflow-hidden"
          >
            {/* Background Scan Grid Lighting */}
            <div className="absolute inset-0 bg-[radial-gradient(ellipse_70%_70%_at_50%_50%,rgba(16,185,129,0.08),rgba(0,0,0,0))] pointer-events-none" />

            <div className="max-w-xl mx-auto space-y-2">
              <h3 className={`font-display font-extrabold text-2xl sm:text-3xl tracking-tight flex items-center justify-center gap-2 ${
                theme === 'dark' ? 'text-white' : 'text-slate-900'
              }`}>
                <Wand2 className="w-6 h-6 text-emerald-500 animate-bounce" />
                <span>Convertendo & Organizando Tabelas</span>
              </h3>
              <p className={`text-xs sm:text-sm ${theme === 'dark' ? 'text-slate-400' : 'text-slate-600'}`}>
                Decodificando layout dos PDFs e estruturando em abas do Excel em tempo real.
              </p>
            </div>

            {/* Visual Pipeline Animation Stage */}
            <div className="grid grid-cols-1 sm:grid-cols-7 gap-4 items-center justify-center max-w-2xl mx-auto py-4">
              
              {/* Left: PDF Document with Scanning Laser */}
              <div className={`sm:col-span-3 rounded-2xl p-4 border relative overflow-hidden flex flex-col items-center space-y-2 group shadow-lg ${
                theme === 'dark' ? 'bg-amber-950/20 border-amber-500/30' : 'bg-amber-50 border-amber-200'
              }`}>
                <div className="w-10 h-10 rounded-xl bg-amber-500/20 text-amber-500 border border-amber-500/30 flex items-center justify-center">
                  <FileText className="w-5 h-5 animate-pulse" />
                </div>
                <span className={`text-xs font-mono font-bold truncate max-w-[160px] ${
                  theme === 'dark' ? 'text-amber-200' : 'text-amber-900'
                }`}>
                  {principalItem?.name || "Documentos PDF"}
                </span>
                <span className="text-[10px] font-mono text-amber-500">Leitura de Coordenadas</span>

                {/* Laser Scanning Line Animation */}
                <motion.div
                  animate={{ y: [0, 50, 0] }}
                  transition={{ repeat: Infinity, duration: 2.2, ease: "easeInOut" }}
                  className="absolute top-2 left-0 right-0 h-0.5 bg-gradient-to-r from-transparent via-amber-500 to-transparent shadow-[0_0_12px_#f59e0b] pointer-events-none"
                />
              </div>

              {/* Middle: Data Stream / Transformation Core */}
              <div className="sm:col-span-1 flex flex-col items-center justify-center py-2 relative">
                <div className="w-10 h-10 rounded-full bg-emerald-500/20 border border-emerald-400/40 text-emerald-500 flex items-center justify-center shadow-lg emerald-glow relative z-10">
                  <RefreshCw className="w-5 h-5 animate-spin" />
                </div>

                {/* Animated Floating Pulse */}
                <motion.div
                  animate={{ scale: [0.8, 1.4, 0.8], opacity: [0.2, 0.7, 0.2] }}
                  transition={{ repeat: Infinity, duration: 1.8 }}
                  className="absolute w-16 h-16 rounded-full bg-emerald-500/10 pointer-events-none"
                />
              </div>

              {/* Right: Excel Sheet Building Rows */}
              <div className={`sm:col-span-3 rounded-2xl p-4 border relative overflow-hidden flex flex-col items-center space-y-2 shadow-lg ${
                theme === 'dark' ? 'bg-emerald-950/20 border-emerald-500/30' : 'bg-emerald-50 border-emerald-200'
              }`}>
                <div className="w-10 h-10 rounded-xl bg-emerald-500/20 text-emerald-500 border border-emerald-500/30 flex items-center justify-center">
                  <FileSpreadsheet className="w-5 h-5" />
                </div>
                <span className={`text-xs font-mono font-bold truncate max-w-[160px] ${
                  theme === 'dark' ? 'text-emerald-200' : 'text-emerald-900'
                }`}>
                  {files.length} Aba{files.length > 1 ? 's' : ''} Excel (.XLSX)
                </span>

                {/* Live Table Rows Building Simulation */}
                <div className="w-full space-y-1.5 pt-1">
                  <motion.div
                    animate={{ opacity: [0.3, 1, 0.3] }}
                    transition={{ repeat: Infinity, duration: 1.2, delay: 0.1 }}
                    className="h-2 rounded bg-emerald-500/30 border border-emerald-500/40 w-full"
                  />
                  <motion.div
                    animate={{ opacity: [0.3, 1, 0.3] }}
                    transition={{ repeat: Infinity, duration: 1.2, delay: 0.3 }}
                    className="h-2 rounded bg-emerald-500/40 border border-emerald-500/50 w-4/5 mx-auto"
                  />
                  <motion.div
                    animate={{ opacity: [0.3, 1, 0.3] }}
                    transition={{ repeat: Infinity, duration: 1.2, delay: 0.5 }}
                    className="h-2 rounded bg-amber-400/30 border border-amber-400/40 w-2/3 mx-auto"
                  />
                </div>
              </div>

            </div>

            {/* Current Processing Step Badge */}
            <div className="space-y-3 max-w-md mx-auto pt-1">
              <motion.p
                key={loadingStep}
                initial={{ opacity: 0, y: 6 }}
                animate={{ opacity: 1, y: 0 }}
                className={`text-xs font-mono font-semibold px-4 py-2 rounded-xl inline-block border shadow-inner tracking-wide ${
                  theme === 'dark' ? 'text-emerald-300 bg-emerald-950/80 border-emerald-500/30' : 'text-emerald-800 bg-emerald-100/80 border-emerald-300'
                }`}
              >
                {loadingStep}
              </motion.p>
            </div>
          </motion.div>
        ) : result ? (
          /* RESULT STAGE */
          <motion.div
            initial={{ opacity: 0, scale: 0.98 }}
            animate={{ opacity: 1, scale: 1 }}
            className="space-y-6"
          >
            {/* Header Banner */}
            <div className="obsidian-card rounded-3xl p-8 border border-emerald-500/30 text-center space-y-4 shadow-2xl relative overflow-hidden">
              <div className="w-16 h-16 rounded-2xl bg-emerald-500/20 border border-emerald-400/40 mx-auto flex items-center justify-center text-emerald-500 shadow-xl emerald-glow mb-2">
                <CheckCircle2 className="w-9 h-9" />
              </div>

              <div className="space-y-1">
                <h2 className={`font-display font-extrabold text-3xl ${theme === 'dark' ? 'text-white' : 'text-slate-900'}`}>
                  Planilha Excel Gerada com Sucesso!
                </h2>
                <p className={`text-xs sm:text-sm max-w-md mx-auto ${theme === 'dark' ? 'text-slate-300' : 'text-slate-600'}`}>
                  Suas tabelas PDF foram unificadas e estruturadas em abas no seu arquivo Excel.
                </p>
              </div>

              {/* Quick Metrics */}
              <div className="grid grid-cols-3 gap-3 max-w-lg mx-auto pt-2">
                <div className={`p-3.5 rounded-2xl border text-center space-y-0.5 ${
                  theme === 'dark' ? 'bg-slate-900/80 border-slate-800' : 'bg-slate-100 border-slate-200'
                }`}>
                  <span className={`text-[10px] font-mono uppercase ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>Abas Criadas</span>
                  <p className="font-display font-extrabold text-2xl text-emerald-500">
                    {result.summary.length}
                  </p>
                </div>

                <div className={`p-3.5 rounded-2xl border text-center space-y-0.5 ${
                  theme === 'dark' ? 'bg-slate-900/80 border-slate-800' : 'bg-slate-100 border-slate-200'
                }`}>
                  <span className={`text-[10px] font-mono uppercase ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>Tabelas</span>
                  <p className="font-display font-extrabold text-2xl text-amber-500">
                    {result.summary.reduce((acc, curr) => acc + curr.tableCount, 0)}
                  </p>
                </div>

                <div className={`p-3.5 rounded-2xl border text-center space-y-0.5 ${
                  theme === 'dark' ? 'bg-slate-900/80 border-slate-800' : 'bg-slate-100 border-slate-200'
                }`}>
                  <span className={`text-[10px] font-mono uppercase ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>Linhas</span>
                  <p className="font-display font-extrabold text-2xl text-emerald-500">
                    {result.summary.reduce((acc, curr) => acc + curr.rowCount, 0)}
                  </p>
                </div>
              </div>
            </div>

            {/* Detailed Table Summary List */}
            <div className="obsidian-card rounded-3xl p-6 border border-slate-800 space-y-4 shadow-xl">
              <h3 className={`text-xs font-mono font-bold uppercase tracking-wider flex items-center gap-2 border-b pb-3 ${
                theme === 'dark' ? 'text-slate-400 border-slate-800' : 'text-slate-600 border-slate-200'
              }`}>
                <FileSpreadsheet className="w-4 h-4 text-emerald-500" />
                Detalhamento das Abas Geradas no Excel
              </h3>

              <div className="space-y-2.5 max-h-[300px] overflow-y-auto pr-1">
                {result.summary.map((item, idx) => (
                  <div
                    key={idx}
                    className={`p-3.5 rounded-2xl border transition-all flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${
                      theme === 'dark' ? 'bg-slate-900/60 border-slate-800 hover:border-slate-700' : 'bg-slate-50 border-slate-200 hover:border-slate-300'
                    }`}
                  >
                    <div className="space-y-0.5">
                      <div className="flex items-center gap-2">
                        <span className="text-xs font-bold font-mono text-emerald-500">
                          {idx + 1}. Aba "{item.tabName}"
                        </span>
                      </div>
                      <p className={`text-[11px] font-mono ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>
                        Arquivo: {item.fileName}
                      </p>
                    </div>

                    <div className="flex items-center gap-2 shrink-0">
                      <span className="px-2.5 py-1 rounded-lg bg-emerald-500/10 border border-emerald-500/20 text-emerald-500 font-mono text-xs font-bold">
                        {item.tableCount} {item.tableCount === 1 ? 'tabela' : 'tabelas'} • {item.rowCount} linhas
                      </span>
                    </div>

                    {item.warning && (
                      <div className="w-full text-xs text-amber-300 bg-amber-950/40 border border-amber-500/30 p-2 rounded-xl flex items-center gap-2">
                        <AlertTriangle className="w-4 h-4 text-amber-400 shrink-0" />
                        <span>{item.warning}</span>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>

            {/* Action Download Buttons */}
            <div className="flex flex-col sm:flex-row gap-4 justify-center pt-2">
              <button
                type="button"
                onClick={triggerDownload}
                className="emerald-gradient-btn emerald-glow text-white py-4 px-8 rounded-2xl font-display font-bold text-base flex items-center justify-center gap-3 shadow-2xl hover:scale-[1.02] active:scale-[0.98] transition-all cursor-pointer"
              >
                <Download className="w-5 h-5 text-white" />
                <span>BAIXAR PLANILHA EXCEL (.XLSX)</span>
              </button>

              <button
                type="button"
                onClick={clearAllFiles}
                className={`py-4 px-6 rounded-2xl border font-semibold text-sm flex items-center justify-center gap-2 transition-colors cursor-pointer ${
                  theme === 'dark' ? 'bg-slate-900 border-slate-800 hover:border-slate-700 text-slate-300' : 'bg-white border-slate-300 hover:bg-slate-100 text-slate-700 shadow-sm'
                }`}
              >
                <RefreshCw className="w-4 h-4 text-slate-400" />
                <span>Converter Novos PDFs</span>
              </button>
            </div>
          </motion.div>
        ) : (
          /* DRAG & DROP FOCUSED DESIGN LAYOUT */
          <div className="space-y-6">

            {/* Hero Heading */}
            <div className="text-center space-y-2 py-2">
              <h2 className={`font-display font-extrabold text-2xl sm:text-3xl tracking-tight ${
                theme === 'dark' ? 'text-white' : 'text-slate-900'
              }`}>
                Arraste & Solte seus Arquivos PDF
              </h2>
              <p className={`text-xs sm:text-sm max-w-lg mx-auto leading-relaxed ${
                theme === 'dark' ? 'text-slate-400' : 'text-slate-600'
              }`}>
                Carregue um ou vários documentos PDF. Extraia tabelas automaticamente e unifique-os em abas no Excel.
              </p>
            </div>

            {/* Main Interactive Studio Dropzone */}
            <div
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
              className={`group relative rounded-3xl p-8 sm:p-12 border-2 border-dashed transition-all duration-300 cursor-pointer flex flex-col items-center justify-center text-center overflow-hidden ${
                isDragOver
                  ? theme === 'dark' ? "border-emerald-400 bg-emerald-950/20 shadow-2xl shadow-emerald-500/20 scale-[1.01]" : "border-emerald-500 bg-emerald-50/90 shadow-xl scale-[1.01]"
                  : files.length > 0
                  ? theme === "dark" ? "border-slate-700 hover:border-emerald-500/50 bg-slate-900/40 hover:bg-slate-900/60" : "border-slate-300 hover:border-emerald-500 bg-white hover:bg-slate-50/80 shadow-md"
                  : theme === "dark" ? "border-slate-700 hover:border-emerald-400/80 bg-slate-900/40 hover:bg-slate-900/70 shadow-2xl" : "border-slate-300 hover:border-emerald-500 bg-white hover:bg-slate-50/80 shadow-xl"
              }`}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept="application/pdf"
                multiple
                className="hidden"
                onChange={handleFileInputChange}
              />

              {/* Ambient Glow in Dropzone */}
              <div className="absolute -top-24 -left-24 w-48 h-48 bg-emerald-500/10 rounded-full blur-3xl group-hover:bg-emerald-500/20 transition-all pointer-events-none" />
              <div className="absolute -bottom-24 -right-24 w-48 h-48 bg-amber-500/10 rounded-full blur-3xl group-hover:bg-amber-500/20 transition-all pointer-events-none" />

              {/* Icon & Pulse Animation */}
              <div className="relative mb-5">
                <div className={`w-20 h-20 rounded-2xl flex items-center justify-center transition-all duration-300 shadow-xl ${
                  isDragOver
                    ? "bg-emerald-500 text-slate-950 scale-110 shadow-emerald-500/40"
                    : theme === "dark" ? "bg-slate-800/90 text-emerald-400 border border-slate-700/80 group-hover:border-emerald-500/50 group-hover:scale-105" : "bg-emerald-50 text-emerald-600 border border-emerald-200 group-hover:border-emerald-400 group-hover:scale-105"
                }`}>
                  <Upload className={`w-9 h-9 transition-transform duration-300 ${isDragOver ? "animate-bounce" : "group-hover:-translate-y-1"}`} />
                </div>
                <div className="absolute -bottom-1 -right-1 w-7 h-7 rounded-lg bg-emerald-400 text-slate-950 flex items-center justify-center shadow-md">
                  <FileSpreadsheet className="w-4 h-4 text-slate-950" />
                </div>
              </div>

              {/* Dropzone Labels */}
              <div className="space-y-1.5 max-w-md relative z-10">
                <h3 className={`font-display font-bold text-lg sm:text-xl transition-colors ${
                  theme === 'dark' ? 'text-white group-hover:text-emerald-300' : 'text-slate-900 group-hover:text-emerald-600'
                }`}>
                  {isDragOver ? "Solte os PDFs aqui para carregar" : "Solte seus PDFs aqui ou clique para buscar"}
                </h3>
                <p className={`text-xs ${theme === 'dark' ? 'text-slate-400' : 'text-slate-600'}`}>
                  Suporta seleção múltipla. O 1º PDF se torna a <span className="text-amber-500 font-semibold">Aba Principal</span> e os seguintes viram <span className="text-emerald-500 font-semibold">Abas Secundárias</span>.
                </p>
              </div>

              {/* Badges */}
              <div className="flex flex-wrap items-center justify-center gap-2 mt-6 relative z-10">
                <span className={`px-3 py-1 rounded-full border text-[11px] font-mono flex items-center gap-1.5 ${
                  theme === 'dark' ? 'bg-slate-800/80 border-slate-700/80 text-slate-300' : 'bg-slate-100 border-slate-200 text-slate-700'
                }`}>
                  <FileText className="w-3.5 h-3.5 text-amber-500" />
                  Formatos PDF
                </span>
                <span className={`px-3 py-1 rounded-full border text-[11px] font-mono flex items-center gap-1.5 ${
                  theme === 'dark' ? 'bg-slate-800/80 border-slate-700/80 text-slate-300' : 'bg-slate-100 border-slate-200 text-slate-700'
                }`}>
                  <Layers className="w-3.5 h-3.5 text-emerald-500" />
                  Leitura de Tabelas
                </span>
                <span className={`px-3 py-1 rounded-full border text-[11px] font-mono flex items-center gap-1.5 ${
                  theme === 'dark' ? 'bg-slate-800/80 border-slate-700/80 text-slate-300' : 'bg-slate-100 border-slate-200 text-slate-700'
                }`}>
                  <ShieldCheck className="w-3.5 h-3.5 text-emerald-500" />
                  Processamento Seguro
                </span>
              </div>
            </div>

            {/* STAGED FILES MANAGEMENT CONTAINER */}
            <AnimatePresence>
              {files.length > 0 && (
                <motion.div
                  initial={{ opacity: 0, y: 15 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -15 }}
                  className="obsidian-card rounded-3xl p-6 sm:p-8 border shadow-2xl space-y-6"
                >
                  {/* Staged Header Bar */}
                  <div className={`flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b pb-4 ${
                    theme === 'dark' ? 'border-slate-800' : 'border-slate-200'
                  }`}>
                    <div className="flex items-center gap-3">
                      <div className="w-8 h-8 rounded-lg bg-emerald-500/10 border border-emerald-500/20 flex items-center justify-center text-emerald-500">
                        <Layers className="w-4 h-4" />
                      </div>
                      <div>
                        <h3 className={`font-display font-bold text-base flex items-center gap-2 ${
                          theme === 'dark' ? 'text-white' : 'text-slate-900'
                        }`}>
                          <span>Arquivos Carregados ({files.length})</span>
                        </h3>
                        <p className={`text-xs ${theme === 'dark' ? 'text-slate-400' : 'text-slate-600'}`}>
                          Arraste qualquer PDF para o topo para torná-lo a <span className="text-amber-500 font-semibold">Aba Principal</span> ou reordenar as abas.
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={() => fileInputRef.current?.click()}
                        className={`px-3 py-1.5 rounded-xl border text-xs font-semibold text-emerald-500 flex items-center gap-1.5 transition-colors cursor-pointer ${
                          theme === 'dark' ? 'bg-slate-900 border-slate-700 hover:border-emerald-500/50' : 'bg-white border-slate-300 hover:bg-slate-50 shadow-sm'
                        }`}
                      >
                        <Plus className="w-3.5 h-3.5" />
                        <span>Adicionar Mais</span>
                      </button>
                      <button
                        type="button"
                        onClick={clearAllFiles}
                        className="px-3 py-1.5 rounded-xl bg-red-950/30 border border-red-500/30 hover:bg-red-950/60 text-xs font-semibold text-red-400 flex items-center gap-1.5 transition-colors cursor-pointer"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                        <span>Limpar Todos</span>
                      </button>
                    </div>
                  </div>

                  {/* Staged Files Display: Hero Principal Card + Grid/List Secondary Cards */}
                  <div className="space-y-4">
                    {/* ABA PRINCIPAL (HERO CARD) */}
                    {principalItem && (
                      <div className="space-y-1.5">
                        <span className="text-[11px] font-mono font-bold uppercase tracking-wider text-amber-500 flex items-center gap-1.5 px-1">
                          <Star className="w-3.5 h-3.5 fill-amber-500 text-amber-500" />
                          Aba Principal (PDF Base)
                        </span>
                        <motion.div
                          key={principalItem.id}
                          layout
                          draggable
                          onDragStart={(e) => handleCardDragStart(e, 0)}
                          onDragOver={(e) => handleCardDragOver(e, 0)}
                          onDrop={(e) => handleCardDrop(e, 0)}
                          onDragEnd={() => {
                            setDraggedFileIndex(null);
                            setDragOverIndex(null);
                          }}
                          initial={{ opacity: 0, y: 10 }}
                          animate={{ opacity: draggedFileIndex === 0 ? 0.4 : 1, y: 0 }}
                          className={`p-4 sm:p-4.5 rounded-2xl border transition-all space-y-3 select-none cursor-grab active:cursor-grabbing ${
                            dragOverIndex === 0
                              ? "border-amber-400 bg-amber-950/40 ring-2 ring-amber-500/40 scale-[1.01]"
                              : theme === 'dark'
                              ? "bg-amber-950/20 border-amber-500/40 shadow-xl shadow-amber-500/5 hover:border-amber-400"
                              : "bg-amber-50/90 border-amber-300 shadow-md hover:border-amber-400"
                          }`}
                        >
                          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                            <div className="flex items-center gap-3 overflow-hidden">
                              <div className="p-1 text-amber-500/70 hover:text-amber-500 transition-colors shrink-0">
                                <GripVertical className="w-5 h-5" />
                              </div>
                              <div className="w-9 h-9 rounded-xl bg-amber-500/20 text-amber-500 border border-amber-500/30 flex items-center justify-center shrink-0 shadow-sm">
                                <Star className="w-5 h-5 fill-amber-500 text-amber-500" />
                              </div>
                              <div className="overflow-hidden">
                                <div className="flex items-center gap-2">
                                  <span className={`text-xs sm:text-sm font-bold truncate ${theme === 'dark' ? 'text-white' : 'text-slate-900'}`}>
                                    {principalItem.name}
                                  </span>
                                  <span className="text-[10px] font-mono font-bold px-2 py-0.5 rounded-full bg-amber-400/20 text-amber-500 border border-amber-500/40 shrink-0">
                                    1º PDF (Base)
                                  </span>
                                </div>
                                <span className={`text-[11px] font-mono ${theme === 'dark' ? 'text-amber-300/80' : 'text-amber-800/80'}`}>
                                  Tamanho: {formatFileSize(principalItem.size)}
                                </span>
                              </div>
                            </div>

                            <div className="flex items-center gap-2 shrink-0 self-end sm:self-center">
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  removeFile(principalItem.id);
                                }}
                                className="p-1.5 rounded-lg text-slate-400 hover:text-red-400 hover:bg-red-500/10 transition-colors cursor-pointer"
                                title="Remover arquivo"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            </div>
                          </div>

                          <div className={`pt-2.5 border-t flex items-center gap-2 ${
                            theme === 'dark' ? 'border-amber-500/20' : 'border-amber-200'
                          }`}>
                            <label className={`text-[11px] font-mono shrink-0 ${theme === 'dark' ? 'text-amber-200/80' : 'text-amber-900/80'}`}>
                              Nome da Aba Principal:
                            </label>
                            <input
                              type="text"
                              value={principalItem.tabName}
                              onChange={(e) => updateTabName(principalItem.id, e.target.value)}
                              onMouseDown={(e) => e.stopPropagation()}
                              placeholder="Principal"
                              className={`w-full max-w-xs border focus:border-amber-500 rounded-lg px-2.5 py-1 text-xs font-mono focus:outline-none transition-colors ${
                                theme === 'dark' ? 'bg-slate-950/80 border-amber-500/30 text-amber-100' : 'bg-white border-amber-300 text-slate-900 shadow-sm'
                              }`}
                            />
                          </div>
                        </motion.div>
                      </div>
                    )}

                    {/* ABAS SECUNDÁRIAS (GRID OU LISTA) */}
                    {files.length > 1 && (
                      <div className="space-y-2 pt-2">
                        <div className="flex items-center justify-between px-1">
                          <span className={`text-[11px] font-mono font-bold uppercase tracking-wider flex items-center gap-1.5 ${
                            theme === 'dark' ? 'text-slate-400' : 'text-slate-600'
                          }`}>
                            <FileText className="w-3.5 h-3.5 text-emerald-500" />
                            Abas Secundárias ({files.length - 1})
                          </span>

                          {/* Seletor de Modo de Visualização: Grade / Lista */}
                          <div className={`flex items-center p-0.5 rounded-lg border text-xs ${
                            theme === 'dark' ? 'bg-slate-900 border-slate-800' : 'bg-slate-100 border-slate-200'
                          }`}>
                            <button
                              type="button"
                              onClick={() => setFileViewMode('grid')}
                              className={`p-1 px-2 rounded-md flex items-center gap-1 font-mono transition-colors cursor-pointer text-[11px] ${
                                fileViewMode === 'grid'
                                  ? 'bg-emerald-500 text-slate-950 font-bold shadow-sm'
                                  : theme === 'dark' ? 'text-slate-400 hover:text-slate-200' : 'text-slate-600 hover:text-slate-900'
                              }`}
                              title="Visualização em Grade"
                            >
                              <LayoutGrid className="w-3.5 h-3.5" />
                              <span className="hidden sm:inline">Grade</span>
                            </button>
                            <button
                              type="button"
                              onClick={() => setFileViewMode('list')}
                              className={`p-1 px-2 rounded-md flex items-center gap-1 font-mono transition-colors cursor-pointer text-[11px] ${
                                fileViewMode === 'list'
                                  ? 'bg-emerald-500 text-slate-950 font-bold shadow-sm'
                                  : theme === 'dark' ? 'text-slate-400 hover:text-slate-200' : 'text-slate-600 hover:text-slate-900'
                              }`}
                              title="Visualização em Lista"
                            >
                              <List className="w-3.5 h-3.5" />
                              <span className="hidden sm:inline">Lista</span>
                            </button>
                          </div>
                        </div>

                        <div className={fileViewMode === 'grid' ? "grid grid-cols-1 sm:grid-cols-2 gap-3" : "space-y-3"}>
                          {files.slice(1).map((item, secondaryIdx) => {
                            const index = secondaryIdx + 1;
                            const isDragged = draggedFileIndex === index;
                            const isOver = dragOverIndex === index;

                            return (
                              <motion.div
                                key={item.id}
                                layout
                                draggable
                                onDragStart={(e) => handleCardDragStart(e, index)}
                                onDragOver={(e) => handleCardDragOver(e, index)}
                                onDrop={(e) => handleCardDrop(e, index)}
                                onDragEnd={() => {
                                  setDraggedFileIndex(null);
                                  setDragOverIndex(null);
                                }}
                                initial={{ opacity: 0, scale: 0.96 }}
                                animate={{ opacity: isDragged ? 0.4 : 1, scale: 1 }}
                                exit={{ opacity: 0, scale: 0.96 }}
                                className={`p-3.5 rounded-2xl border transition-all flex flex-col justify-between gap-3 select-none cursor-grab active:cursor-grabbing ${
                                  isOver
                                    ? "border-emerald-400 bg-emerald-950/40 ring-2 ring-emerald-500/40 scale-[1.01]"
                                    : theme === 'dark'
                                    ? "bg-slate-900/60 border-slate-800 hover:border-slate-700 hover:bg-slate-900/80 shadow-md"
                                    : "bg-slate-50 border-slate-200 hover:border-slate-300 hover:bg-white shadow-sm"
                                }`}
                              >
                                <div className="flex items-start justify-between gap-2">
                                  <div className="flex items-center gap-2.5 overflow-hidden">
                                    <div className={`transition-colors p-0.5 rounded shrink-0 ${
                                      theme === 'dark' ? 'text-slate-500 hover:text-slate-300' : 'text-slate-400 hover:text-slate-600'
                                    }`}>
                                      <GripVertical className="w-4 h-4" />
                                    </div>
                                    <div className={`w-8 h-8 rounded-lg flex items-center justify-center shrink-0 ${
                                      theme === 'dark' ? 'bg-slate-800 text-emerald-400 border border-slate-700' : 'bg-emerald-100 text-emerald-700 border border-emerald-200'
                                    }`}>
                                      <FileText className="w-4 h-4" />
                                    </div>
                                    <div className="overflow-hidden">
                                      <div className="flex items-center gap-1.5">
                                        <span className={`text-xs font-bold truncate max-w-[130px] sm:max-w-[160px] ${
                                          theme === 'dark' ? 'text-slate-100' : 'text-slate-900'
                                        }`} title={item.name}>
                                          {item.name}
                                        </span>
                                      </div>
                                      <span className={`text-[10px] font-mono ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>
                                        Aba {index + 1} • {formatFileSize(item.size)}
                                      </span>
                                    </div>
                                  </div>

                                  <div className="flex items-center gap-1 shrink-0">
                                    <button
                                      type="button"
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        setAsPrincipal(item.id);
                                      }}
                                      className="p-1 px-2 rounded-lg bg-amber-500/10 hover:bg-amber-500/20 text-amber-500 border border-amber-500/30 text-[10px] font-semibold transition-colors flex items-center gap-1 cursor-pointer"
                                      title="Tornar PDF Principal"
                                    >
                                      <Star className="w-3 h-3 text-amber-500" />
                                      <span className="hidden sm:inline">Principal</span>
                                    </button>
                                    <button
                                      type="button"
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        removeFile(item.id);
                                      }}
                                      className="p-1.5 rounded-lg text-slate-400 hover:text-red-400 hover:bg-red-500/10 transition-colors cursor-pointer"
                                      title="Remover arquivo"
                                    >
                                      <Trash2 className="w-3.5 h-3.5" />
                                    </button>
                                  </div>
                                </div>

                                <div className={`pt-2 border-t flex items-center gap-2 ${
                                  theme === 'dark' ? 'border-slate-800/80' : 'border-slate-200'
                                }`}>
                                  <label className={`text-[10px] font-mono shrink-0 ${
                                    theme === 'dark' ? 'text-slate-400' : 'text-slate-600'
                                  }`}>
                                    Aba Excel:
                                  </label>
                                  <input
                                    type="text"
                                    value={item.tabName}
                                    onChange={(e) => updateTabName(item.id, e.target.value)}
                                    onMouseDown={(e) => e.stopPropagation()}
                                    placeholder={`Aba ${index + 1}`}
                                    className={`w-full border focus:border-emerald-500 rounded-lg px-2 py-1 text-xs font-mono focus:outline-none transition-colors ${
                                      theme === 'dark' ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-white border-slate-300 text-slate-900'
                                    }`}
                                  />
                                </div>
                              </motion.div>
                            );
                          })}
                        </div>
                      </div>
                    )}
                  </div>

                  {/* Formatting Options Toggles */}
                  <div className={`pt-4 border-t space-y-3 ${theme === 'dark' ? 'border-slate-800' : 'border-slate-200'}`}>
                    <div className={`flex items-center gap-2 text-xs font-semibold ${theme === 'dark' ? 'text-slate-300' : 'text-slate-700'}`}>
                      <Settings2 className="w-4 h-4 text-emerald-500" />
                      <span>Opções de Formatação</span>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                      <label className={`flex items-start gap-3 p-3 rounded-2xl border transition-all cursor-pointer select-none group hover:border-slate-400 ${
                        theme === 'dark' ? 'bg-slate-900/50 border-slate-800/80' : 'bg-slate-100/70 border-slate-200'
                      }`}>
                        <input
                          type="checkbox"
                          checked={disablePrincipalTab}
                          onChange={(e) => setDisablePrincipalTab(e.target.checked)}
                          className="mt-0.5 w-4 h-4 rounded border-slate-700 bg-slate-950 text-emerald-500 focus:ring-emerald-500/20 accent-emerald-500 cursor-pointer"
                        />
                        <div className="text-xs">
                          <span className={`font-semibold group-hover:text-emerald-500 transition-colors block ${
                            theme === 'dark' ? 'text-slate-200' : 'text-slate-800'
                          }`}>
                            Desabilitar aba principal
                          </span>
                          <span className={`text-[10px] mt-0.5 block ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>
                            Torna a aba principal secundária como as demais.
                          </span>
                        </div>
                      </label>

                      <label className={`flex items-start gap-3 p-3 rounded-2xl border transition-all select-none ${
                        disablePrincipalTab
                          ? 'opacity-50 cursor-not-allowed border-slate-800/40 bg-slate-900/20'
                          : 'cursor-pointer group hover:border-slate-400 ' + (theme === 'dark' ? 'bg-slate-900/50 border-slate-800/80' : 'bg-slate-100/70 border-slate-200')
                      }`}>
                        <input
                          type="checkbox"
                          checked={hideResponsavel}
                          disabled={disablePrincipalTab}
                          onChange={(e) => setHideResponsavel(e.target.checked)}
                          className="mt-0.5 w-4 h-4 rounded border-slate-700 bg-slate-950 text-emerald-500 focus:ring-emerald-500/20 accent-emerald-500 cursor-pointer disabled:cursor-not-allowed"
                        />
                        <div className="text-xs">
                          <span className={`font-semibold transition-colors block ${
                            disablePrincipalTab
                              ? (theme === 'dark' ? 'text-slate-500' : 'text-slate-400')
                              : 'group-hover:text-emerald-500 ' + (theme === 'dark' ? 'text-slate-200' : 'text-slate-800')
                          }`}>
                            Ocultar coluna "Responsável"
                          </span>
                          <span className={`text-[10px] mt-0.5 block ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>
                            Oculta a coluna no Excel final.
                          </span>
                        </div>
                      </label>

                      <label className={`flex items-start gap-3 p-3 rounded-2xl border transition-all select-none ${
                        disablePrincipalTab
                          ? 'opacity-50 cursor-not-allowed border-slate-800/40 bg-slate-900/20'
                          : 'cursor-pointer group hover:border-slate-400 ' + (theme === 'dark' ? 'bg-slate-900/50 border-slate-800/80' : 'bg-slate-100/70 border-slate-200')
                      }`}>
                        <input
                          type="checkbox"
                          checked={hideSecondaryTabs}
                          disabled={disablePrincipalTab}
                          onChange={(e) => setHideSecondaryTabs(e.target.checked)}
                          className="mt-0.5 w-4 h-4 rounded border-slate-700 bg-slate-950 text-emerald-500 focus:ring-emerald-500/20 accent-emerald-500 cursor-pointer disabled:cursor-not-allowed"
                        />
                        <div className="text-xs">
                          <span className={`font-semibold transition-colors block ${
                            disablePrincipalTab
                              ? (theme === 'dark' ? 'text-slate-500' : 'text-slate-400')
                              : 'group-hover:text-emerald-500 ' + (theme === 'dark' ? 'text-slate-200' : 'text-slate-800')
                          }`}>
                            Ocultar abas secundárias
                          </span>
                          <span className={`text-[10px] mt-0.5 block ${theme === 'dark' ? 'text-slate-400' : 'text-slate-500'}`}>
                            Mantém dados mas oculta as abas secundárias.
                          </span>
                        </div>
                      </label>
                    </div>
                  </div>

                  {/* Primary CTA Button */}
                  <div className="pt-2">
                    <button
                      type="button"
                      onClick={handleExecute}
                      className="w-full py-4 px-6 rounded-2xl font-display font-bold text-sm tracking-wide flex items-center justify-center gap-2.5 emerald-gradient-btn text-white emerald-glow cursor-pointer active:scale-[0.99] transition-all shadow-xl"
                    >
                      <FileSpreadsheet className="w-5 h-5 text-emerald-200" />
                      <span>GERAR PLANILHA EXCEL INTEGRADA</span>
                    </button>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

          </div>
        )}

      </main>
    </div>
  );
}



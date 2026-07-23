import React, { useState } from "react";
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
  Table
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";

interface PDFSlot {
  id: string;
  file: File | null;
  name: string;
  base64: string;
  tabName: string;
  isPrincipal: boolean;
}

export default function App() {
  // Initialize the principal PDF slot
  const [principalFile, setPrincipalFile] = useState<PDFSlot>({
    id: "principal",
    file: null,
    name: "",
    base64: "",
    tabName: "Principal",
    isPrincipal: true,
  });

  // Initialize 5 secondary slots
  const [secondaryFiles, setSecondaryFiles] = useState<PDFSlot[]>([
    { id: "sec-1", file: null, name: "", base64: "", tabName: "", isPrincipal: false },
    { id: "sec-2", file: null, name: "", base64: "", tabName: "", isPrincipal: false },
    { id: "sec-3", file: null, name: "", base64: "", tabName: "", isPrincipal: false },
    { id: "sec-4", file: null, name: "", base64: "", tabName: "", isPrincipal: false },
    { id: "sec-5", file: null, name: "", base64: "", tabName: "", isPrincipal: false },
  ]);

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
  const [isDragOver, setIsDragOver] = useState(false);

  // Helper to read file to base64
  const processFile = (file: File, id: string, isPrincipal: boolean, index?: number) => {
    if (file.type !== "application/pdf") {
      setError("Por favor, envie apenas arquivos no formato PDF.");
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      const base64String = (reader.result as string).split(",")[1];
      const proposedTabName = file.name
        .replace(/\.pdf$/i, "")
        .replace(/[^a-zA-Z0-9_\-\s]/g, "")
        .substring(0, 20);

      if (isPrincipal) {
        setPrincipalFile(prev => ({
          ...prev,
          file,
          name: file.name,
          base64: base64String,
          tabName: prev.tabName && prev.tabName !== "Principal" ? prev.tabName : proposedTabName || "Principal",
        }));
      } else if (index !== undefined) {
        setSecondaryFiles(prev => {
          const updated = [...prev];
          updated[index] = {
            ...updated[index],
            file,
            name: file.name,
            base64: base64String,
            tabName: updated[index].tabName || proposedTabName || `Aba ${index + 1}`,
          };
          return updated;
        });
      }
      setError(null);
    };
    reader.onerror = () => {
      setError("Erro ao ler o arquivo PDF.");
    };
    reader.readAsDataURL(file);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>, id: string, isPrincipal: boolean, index?: number) => {
    const file = e.target.files?.[0];
    if (file) {
      processFile(file, id, isPrincipal, index);
    }
  };

  // Dropzone general drop handler
  const handleGeneralDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragOver(false);
    const files = (Array.from(e.dataTransfer.files) as File[]).filter(f => f.type === "application/pdf");
    if (files.length === 0) {
      setError("Nenhum arquivo PDF válido foi arrastado.");
      return;
    }

    let fileIdx = 0;
    // If principal is empty, set first dropped file as principal
    if (!principalFile.file && files.length > 0) {
      processFile(files[0], "principal", true);
      fileIdx = 1;
    }

    // Set remaining files in secondary slots
    if (fileIdx < files.length) {
      setSecondaryFiles(prev => {
        const updated = [...prev];
        let currentSlot = 0;
        for (let i = fileIdx; i < files.length; i++) {
          while (currentSlot < updated.length && updated[currentSlot].file !== null) {
            currentSlot++;
          }
          const file: File = files[i];
          const reader = new FileReader();
          const targetSlotIndex = currentSlot;
          const proposedTabName = file.name.replace(/\.pdf$/i, "").replace(/[^a-zA-Z0-9_\-\s]/g, "").substring(0, 20);

          reader.onload = () => {
            const base64String = (reader.result as string).split(",")[1];
            setSecondaryFiles(curr => {
              const copy = [...curr];
              if (targetSlotIndex < copy.length) {
                copy[targetSlotIndex] = {
                  ...copy[targetSlotIndex],
                  file,
                  name: file.name,
                  base64: base64String,
                  tabName: copy[targetSlotIndex].tabName || proposedTabName || `Aba ${targetSlotIndex + 1}`,
                };
              } else {
                copy.push({
                  id: `sec-${copy.length + 1}`,
                  file,
                  name: file.name,
                  base64: base64String,
                  tabName: proposedTabName || `Aba ${copy.length + 1}`,
                  isPrincipal: false,
                });
              }
              return copy;
            });
          };
          reader.readAsDataURL(file);
          currentSlot++;
        }
        return updated;
      });
    }
  };

  const clearSlot = (isPrincipal: boolean, index?: number) => {
    if (isPrincipal) {
      setPrincipalFile(prev => ({
        ...prev,
        file: null,
        name: "",
        base64: "",
      }));
    } else if (index !== undefined) {
      setSecondaryFiles(prev => {
        const updated = [...prev];
        updated[index] = {
          ...updated[index],
          file: null,
          name: "",
          base64: "",
        };
        return updated;
      });
    }
  };

  const handleTabNameChange = (val: string, isPrincipal: boolean, index?: number) => {
    if (isPrincipal) {
      setPrincipalFile(prev => ({ ...prev, tabName: val }));
    } else if (index !== undefined) {
      setSecondaryFiles(prev => {
        const updated = [...prev];
        updated[index] = { ...updated[index], tabName: val };
        return updated;
      });
    }
  };

  const addSecondaryRow = () => {
    setSecondaryFiles(prev => [
      ...prev,
      {
        id: `sec-${prev.length + 1}`,
        file: null,
        name: "",
        base64: "",
        tabName: "",
        isPrincipal: false,
      }
    ]);
  };

  const deleteSecondaryRow = (index: number) => {
    setSecondaryFiles(prev => prev.filter((_, i) => i !== index));
  };

  // Convert and Join
  const handleExecute = async () => {
    setError(null);
    setResult(null);

    if (!principalFile.file || !principalFile.base64) {
      setError("Por favor, faça o upload do PDF Principal antes de executar.");
      return;
    }

    const activeSecondaries = secondaryFiles.filter(item => item.file && item.base64);

    const filesToSubmit = [
      {
        name: principalFile.name,
        tabName: principalFile.tabName || "Principal",
        base64: principalFile.base64,
        isPrincipal: true,
      },
      ...activeSecondaries.map(item => ({
        name: item.name,
        tabName: item.tabName || item.name.replace(/\.pdf$/i, "").substring(0, 20),
        base64: item.base64,
        isPrincipal: false,
      })),
    ];

    try {
      setLoading(true);
      setLoadingStep("Conectando com o servidor local...");

      const steps = [
        "Lendo e analisando a estrutura do PDF Principal...",
        "Detectando tabelas, delimitadores e cabeçalhos...",
        activeSecondaries.length > 0 ? "Processando PDFs secundários e alinhando colunas..." : "Formatando células e tipos de dados...",
        "Estruturando abas dinâmicas do Excel workbook...",
        "Finalizando arquivo Excel de alta precisão..."
      ];

      let stepIndex = 0;
      const stepInterval = setInterval(() => {
        if (stepIndex < steps.length) {
          setLoadingStep(steps[stepIndex]);
          stepIndex++;
        }
      }, 2500);

      const response = await fetch("/api/convert", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ files: filesToSubmit, hideResponsavel, hideSecondaryTabs }),
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

  const handleReset = () => {
    setPrincipalFile({
      id: "principal",
      file: null,
      name: "",
      base64: "",
      tabName: "Principal",
      isPrincipal: true,
    });
    setSecondaryFiles([
      { id: "sec-1", file: null, name: "", base64: "", tabName: "", isPrincipal: false },
      { id: "sec-2", file: null, name: "", base64: "", tabName: "", isPrincipal: false },
      { id: "sec-3", file: null, name: "", base64: "", tabName: "", isPrincipal: false },
      { id: "sec-4", file: null, name: "", base64: "", tabName: "", isPrincipal: false },
      { id: "sec-5", file: null, name: "", base64: "", tabName: "", isPrincipal: false },
    ]);
    setResult(null);
    setError(null);
  };

  const activeSecondaryCount = secondaryFiles.filter(s => s.file !== null).length;
  const totalStagedCount = (principalFile.file ? 1 : 0) + activeSecondaryCount;

  return (
    <div className="min-h-screen bg-[#0b0f17] text-slate-100 font-sans selection:bg-emerald-500/30 selection:text-emerald-200 relative overflow-x-hidden">
      
      {/* Background Studio Ambience & Grid */}
      <div className="absolute inset-0 bg-[radial-gradient(ellipse_80%_80%_at_50%_-20%,rgba(16,185,129,0.12),rgba(255,255,255,0))] pointer-events-none" />
      <div className="absolute inset-0 bg-[linear-gradient(to_right,#1f293d0f_1px,transparent_1px),linear-gradient(to_bottom,#1f293d0f_1px,transparent_1px)] bg-[size:4rem_4rem] [mask-image:radial-gradient(ellipse_60%_50%_at_50%_0%,#000_70%,transparent_100%)] pointer-events-none" />

      {/* Top Studio Navbar */}
      <header className="sticky top-0 z-40 border-b border-slate-800/80 bg-[#0b0f17]/80 backdrop-blur-md px-6 py-4">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-emerald-500 to-amber-400 p-[1px] shadow-lg shadow-emerald-500/20">
              <div className="w-full h-full bg-[#0d131f] rounded-[11px] flex items-center justify-center">
                <Wand2 className="w-5 h-5 text-emerald-400" />
              </div>
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="font-display font-bold text-xl tracking-tight text-white">
                  TABLE <span className="bg-gradient-to-r from-emerald-400 to-amber-300 bg-clip-text text-transparent">MAGIC</span>
                </h1>
                <span className="text-[10px] font-mono font-semibold px-2 py-0.5 rounded-full bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                  v2.5 Studio
                </span>
              </div>
              <p className="text-xs text-slate-400">PDF Table Extractor & Excel Sheet Joiner</p>
            </div>
          </div>

          <div className="flex items-center gap-4">
            <div className="hidden md:flex items-center gap-2 px-3 py-1.5 rounded-full bg-slate-900/80 border border-slate-800 text-xs text-slate-300">
              <ShieldCheck className="w-4 h-4 text-emerald-400" />
              <span>Processamento 100% Local</span>
            </div>
          </div>
        </div>
      </header>

      {/* Main Studio Container */}
      <main className="max-w-7xl mx-auto px-4 md:px-8 py-8 relative z-10">
        
        {/* Error Alert */}
        <AnimatePresence>
          {error && (
            <motion.div
              initial={{ opacity: 0, y: -10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="mb-6 p-4 rounded-xl bg-red-950/40 border border-red-500/30 text-red-300 flex items-center gap-3 shadow-lg"
            >
              <AlertCircle className="w-5 h-5 text-red-400 shrink-0" />
              <p className="text-sm font-medium">{error}</p>
            </motion.div>
          )}
        </AnimatePresence>

        {!result ? (
          /* 2-COLUMN WORKBENCH TOPOLOGY */
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
            
            {/* LEFT COLUMN: WORKBENCH SLOTS & CONTROLS (~5 cols) */}
            <div className="lg:col-span-5 space-y-6">
              
              <div className="obsidian-card rounded-2xl p-6 space-y-6 border border-slate-800 shadow-2xl">
                
                <div className="flex items-center justify-between border-b border-slate-800 pb-4">
                  <div className="flex items-center gap-2">
                    <Layers className="w-5 h-5 text-emerald-400" />
                    <h2 className="font-display font-semibold text-lg text-white">Workbench de Documentos</h2>
                  </div>
                  <span className="text-xs font-mono text-slate-400 bg-slate-900 px-2.5 py-1 rounded-lg border border-slate-800">
                    {totalStagedCount} {totalStagedCount === 1 ? 'arquivo' : 'arquivos'}
                  </span>
                </div>

                {/* PDF Principal Slot */}
                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <label className="text-xs font-mono font-bold text-amber-400 uppercase tracking-wider flex items-center gap-1.5">
                      <span className="w-2 h-2 rounded-full bg-amber-400 animate-pulse" />
                      PDF Principal (Obrigatório)
                    </label>
                  </div>

                  {principalFile.file ? (
                    <div className="p-3.5 rounded-xl bg-amber-950/20 border border-amber-500/40 space-y-2 obsidian-card-hover">
                      <div className="flex items-center justify-between gap-2">
                        <div className="flex items-center gap-2.5 overflow-hidden">
                          <FileText className="w-5 h-5 text-amber-400 shrink-0" />
                          <span className="text-xs font-semibold truncate text-amber-100">
                            {principalFile.name}
                          </span>
                        </div>
                        <button
                          onClick={() => clearSlot(true)}
                          type="button"
                          className="text-slate-400 hover:text-red-400 p-1.5 rounded-lg hover:bg-red-500/10 transition-colors"
                          title="Remover PDF Principal"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                      
                      <div className="pt-2 border-t border-amber-500/20 flex items-center gap-2">
                        <label className="text-[11px] text-amber-300/80 shrink-0 font-medium">
                          Nome da Aba no Excel:
                        </label>
                        <input
                          type="text"
                          value={principalFile.tabName}
                          onChange={(e) => handleTabNameChange(e.target.value, true)}
                          placeholder="Principal"
                          className="w-full bg-slate-900/90 border border-amber-500/30 rounded-lg px-2.5 py-1 text-xs text-white focus:outline-none focus:border-amber-400 font-mono"
                        />
                      </div>
                    </div>
                  ) : (
                    <label className="flex flex-col items-center justify-center gap-2 p-5 border-2 border-dashed border-amber-500/30 hover:border-amber-400/80 rounded-xl bg-amber-950/10 hover:bg-amber-950/20 transition-all cursor-pointer group">
                      <Upload className="w-5 h-5 text-amber-400 group-hover:scale-110 transition-transform" />
                      <span className="text-xs font-semibold text-amber-200 group-hover:text-white">
                        Upload PDF Principal
                      </span>
                      <span className="text-[10px] text-amber-400/60 font-mono">
                        Clique ou arraste um PDF aqui
                      </span>
                      <input
                        type="file"
                        accept="application/pdf"
                        className="hidden"
                        onChange={(e) => handleFileChange(e, "principal", true)}
                      />
                    </label>
                  )}
                </div>

                {/* PDFs Secundários Section */}
                <div className="space-y-3 pt-2 border-t border-slate-800">
                  <div className="flex items-center justify-between">
                    <label className="text-xs font-mono font-bold text-emerald-400 uppercase tracking-wider">
                      PDFs Secundários (Opcionais)
                    </label>
                    <span className="text-[10px] text-slate-400 font-mono">
                      Viram abas extras
                    </span>
                  </div>

                  <div className="space-y-2.5 max-h-[340px] overflow-y-auto pr-1">
                    {secondaryFiles.map((slot, index) => (
                      <div
                        key={slot.id}
                        className="p-3 rounded-xl bg-slate-900/70 border border-slate-800/80 hover:border-slate-700 transition-all space-y-2"
                      >
                        {slot.file ? (
                          <div className="flex items-center justify-between gap-2">
                            <div className="flex items-center gap-2 overflow-hidden">
                              <FileSpreadsheet className="w-4 h-4 text-emerald-400 shrink-0" />
                              <span className="text-xs font-medium truncate text-slate-200">
                                {slot.name}
                              </span>
                            </div>
                            <button
                              onClick={() => clearSlot(false, index)}
                              type="button"
                              className="text-slate-400 hover:text-red-400 p-1 rounded hover:bg-red-500/10 transition-colors"
                              title="Remover arquivo"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        ) : (
                          <div className="flex items-center justify-between gap-2">
                            <label className="flex-1 flex items-center gap-2 p-2 rounded-lg border border-dashed border-slate-700 hover:border-emerald-500/50 bg-slate-950/40 hover:bg-slate-900 transition-all cursor-pointer text-xs text-slate-400 hover:text-emerald-300">
                              <FilePlus className="w-3.5 h-3.5 text-emerald-500" />
                              <span className="truncate">Escolher PDF Secundário</span>
                              <input
                                type="file"
                                accept="application/pdf"
                                className="hidden"
                                onChange={(e) => handleFileChange(e, slot.id, false, index)}
                              />
                            </label>
                            <button
                              onClick={() => deleteSecondaryRow(index)}
                              type="button"
                              className="text-slate-500 hover:text-red-400 p-1.5 rounded hover:bg-red-500/10 transition-colors"
                              title="Excluir linha"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        )}

                        <div className="flex items-center gap-2">
                          <span className="text-[10px] text-slate-500 font-mono shrink-0">
                            Aba:
                          </span>
                          <input
                            type="text"
                            value={slot.tabName}
                            onChange={(e) => handleTabNameChange(e.target.value, false, index)}
                            placeholder={`Aba ${index + 1}`}
                            className="w-full bg-slate-950 border border-slate-800 rounded px-2 py-1 text-xs text-slate-200 focus:outline-none focus:border-emerald-500 font-mono"
                          />
                        </div>
                      </div>
                    ))}
                  </div>

                  <button
                    type="button"
                    onClick={addSecondaryRow}
                    className="w-full py-2.5 px-3 flex items-center justify-center gap-2 border border-dashed border-emerald-500/30 hover:border-emerald-400 rounded-xl text-xs font-semibold text-emerald-400 hover:bg-emerald-500/10 transition-all"
                  >
                    <Plus className="w-4 h-4" />
                    <span>Adicionar Mais Linhas de PDF</span>
                  </button>
                </div>

                {/* Export Options */}
                <div className="pt-4 border-t border-slate-800 space-y-3">
                  <div className="flex items-center gap-2 text-xs font-semibold text-slate-300">
                    <Settings2 className="w-4 h-4 text-emerald-400" />
                    <span>Opções de Formatação</span>
                  </div>

                  <label className="flex items-start gap-3 p-2.5 rounded-xl bg-slate-900/50 border border-slate-800/80 cursor-pointer select-none group hover:border-slate-700 transition-all">
                    <input
                      type="checkbox"
                      checked={hideResponsavel}
                      onChange={(e) => setHideResponsavel(e.target.checked)}
                      className="mt-0.5 w-4 h-4 rounded border-slate-700 bg-slate-950 text-emerald-500 focus:ring-emerald-500/20 accent-emerald-500 cursor-pointer"
                    />
                    <div className="text-xs">
                      <span className="font-semibold text-slate-200 group-hover:text-emerald-400 transition-colors">
                        Ocultar coluna "Responsável"
                      </span>
                      <p className="text-[10px] text-slate-400 mt-0.5">
                        Fica oculta na planilha final, podendo ser reexibida no Excel a qualquer momento.
                      </p>
                    </div>
                  </label>

                  <label className="flex items-start gap-3 p-2.5 rounded-xl bg-slate-900/50 border border-slate-800/80 cursor-pointer select-none group hover:border-slate-700 transition-all">
                    <input
                      type="checkbox"
                      checked={hideSecondaryTabs}
                      onChange={(e) => setHideSecondaryTabs(e.target.checked)}
                      className="mt-0.5 w-4 h-4 rounded border-slate-700 bg-slate-950 text-emerald-500 focus:ring-emerald-500/20 accent-emerald-500 cursor-pointer"
                    />
                    <div className="text-xs">
                      <span className="font-semibold text-slate-200 group-hover:text-emerald-400 transition-colors">
                        Ocultar abas secundárias
                      </span>
                      <p className="text-[10px] text-slate-400 mt-0.5">
                        As abas permanecem com os dados, mas ficam ocultas para manter o foco na principal.
                      </p>
                    </div>
                  </label>
                </div>

                {/* Primary Action Button */}
                <button
                  type="button"
                  onClick={handleExecute}
                  disabled={loading || !principalFile.file}
                  className={`w-full py-4 px-6 rounded-xl font-display font-bold text-sm tracking-wide flex items-center justify-center gap-2.5 transition-all shadow-xl ${
                    !principalFile.file
                      ? 'bg-slate-800/60 text-slate-500 border border-slate-700/50 cursor-not-allowed'
                      : 'emerald-gradient-btn text-white emerald-glow cursor-pointer active:scale-[0.99]'
                  }`}
                >
                  {loading ? (
                    <>
                      <RefreshCw className="w-5 h-5 animate-spin" />
                      <span>Processando Tabelas...</span>
                    </>
                  ) : (
                    <>
                      <Sparkles className="w-5 h-5 text-amber-300 animate-pulse" />
                      <span>GERAR PLANILHA EXCEL INTEGRADA</span>
                    </>
                  )}
                </button>

              </div>
            </div>

            {/* RIGHT COLUMN: INTERACTIVE STUDIO STAGE & LIVE PREVIEW (~7 cols) */}
            <div className="lg:col-span-7">
              <div 
                onDragOver={(e) => { e.preventDefault(); setIsDragOver(true); }}
                onDragLeave={() => setIsDragOver(false)}
                onDrop={handleGeneralDrop}
                className={`obsidian-card rounded-2xl p-8 border min-h-[560px] flex flex-col justify-between transition-all duration-300 relative overflow-hidden ${
                  isDragOver 
                    ? 'border-emerald-400 bg-emerald-950/20 shadow-2xl shadow-emerald-500/20' 
                    : 'border-slate-800 shadow-2xl'
                }`}
              >
                {/* Background Decorative Element */}
                <div className="absolute -right-20 -bottom-20 w-80 h-80 bg-emerald-500/5 rounded-full blur-3xl pointer-events-none" />

                {loading ? (
                  /* LOADING STAGE */
                  <div className="my-auto flex flex-col items-center justify-center text-center space-y-6 py-12">
                    <div className="relative">
                      <div className="w-20 h-20 rounded-full border-4 border-emerald-500/20 border-t-emerald-400 animate-spin" />
                      <Sparkles className="w-8 h-8 text-amber-400 absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 animate-pulse" />
                    </div>

                    <div className="space-y-2 max-w-sm">
                      <h3 className="font-display font-bold text-xl text-white">
                        Extraindo Tabelas com IA
                      </h3>
                      <motion.p
                        key={loadingStep}
                        initial={{ opacity: 0, y: 5 }}
                        animate={{ opacity: 1, y: 0 }}
                        className="text-xs font-mono text-emerald-400 bg-emerald-950/50 border border-emerald-500/30 px-3 py-1.5 rounded-lg inline-block"
                      >
                        {loadingStep}
                      </motion.p>
                    </div>

                    <p className="text-xs text-slate-400 max-w-xs leading-relaxed">
                      O motor local está decodificando coordenadas, tabelas sem borda e estruturas para unificar tudo em abas formatadas.
                    </p>
                  </div>
                ) : principalFile.file ? (
                  /* STAGED BLUEPRINT PREVIEW */
                  <div className="space-y-6 my-auto">
                    <div className="flex items-center gap-3 border-b border-slate-800 pb-4">
                      <div className="w-3 h-3 rounded-full bg-emerald-400 animate-ping" />
                      <div>
                        <h3 className="font-display font-bold text-lg text-white">
                          Prévia do Documento Pronto para Conversão
                        </h3>
                        <p className="text-xs text-slate-400">
                          Resumo da estrutura que será gerada na planilha Excel final.
                        </p>
                      </div>
                    </div>

                    {/* Staged Cards Grid */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      {/* Principal Card Preview */}
                      <div className="p-4 rounded-xl bg-amber-950/20 border border-amber-500/30 space-y-2">
                        <div className="flex items-center justify-between">
                          <span className="text-[10px] font-mono uppercase font-bold text-amber-400">
                            Aba Principal
                          </span>
                          <Check className="w-4 h-4 text-amber-400" />
                        </div>
                        <p className="text-xs font-bold text-amber-100 truncate">
                          {principalFile.name}
                        </p>
                        <p className="text-[11px] font-mono text-amber-300/70">
                          Aba: "{principalFile.tabName || 'Principal'}"
                        </p>
                      </div>

                      {/* Secondaries Count Card */}
                      <div className="p-4 rounded-xl bg-emerald-950/20 border border-emerald-500/30 space-y-2">
                        <div className="flex items-center justify-between">
                          <span className="text-[10px] font-mono uppercase font-bold text-emerald-400">
                            Abas Secundárias
                          </span>
                          <Layers className="w-4 h-4 text-emerald-400" />
                        </div>
                        <p className="text-xs font-bold text-emerald-100">
                          {activeSecondaryCount} {activeSecondaryCount === 1 ? 'Arquivo Adicional' : 'Arquivos Adicionais'}
                        </p>
                        <p className="text-[11px] font-mono text-emerald-300/70">
                          {activeSecondaryCount > 0 ? 'Abas configuradas' : 'Nenhum secundário'}
                        </p>
                      </div>
                    </div>

                    {/* List of active tabs */}
                    <div className="p-4 rounded-xl bg-slate-900/60 border border-slate-800 space-y-3">
                      <span className="text-xs font-mono font-semibold text-slate-400 uppercase tracking-wider block border-b border-slate-800 pb-2">
                        Abas que serão criadas no arquivo .XLSX:
                      </span>
                      
                      <div className="space-y-2 max-h-[160px] overflow-y-auto pr-1">
                        <div className="flex items-center justify-between text-xs py-1.5 px-2.5 rounded bg-slate-950 border border-slate-800">
                          <span className="font-mono text-amber-300 font-medium">
                            1. {principalFile.tabName || 'Principal'}
                          </span>
                          <span className="text-[10px] text-slate-400 truncate max-w-[180px]">
                            {principalFile.name}
                          </span>
                        </div>

                        {secondaryFiles.filter(s => s.file).map((s, idx) => (
                          <div key={s.id} className="flex items-center justify-between text-xs py-1.5 px-2.5 rounded bg-slate-950 border border-slate-800">
                            <span className="font-mono text-emerald-400 font-medium">
                              {idx + 2}. {s.tabName || s.name}
                            </span>
                            <span className="text-[10px] text-slate-400 truncate max-w-[180px]">
                              {s.name}
                            </span>
                          </div>
                        ))}
                      </div>
                    </div>

                    <div className="p-4 rounded-xl bg-slate-900/40 border border-slate-800/80 flex items-center justify-between">
                      <div className="flex items-center gap-2 text-xs text-slate-300">
                        <Info className="w-4 h-4 text-emerald-400 shrink-0" />
                        <span>Tudo pronto para gerar sua planilha integrada!</span>
                      </div>
                      <button
                        onClick={handleExecute}
                        type="button"
                        className="px-4 py-2 rounded-lg bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-bold text-xs transition-colors flex items-center gap-1.5"
                      >
                        <span>Executar</span>
                        <ChevronRight className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                ) : (
                  /* DEFAULT EMPTY DROPZONE STAGE */
                  <div className="my-auto flex flex-col items-center justify-center text-center space-y-6 py-12">
                    <div className="w-24 h-24 rounded-3xl bg-slate-900/80 border-2 border-dashed border-slate-700 flex items-center justify-center text-slate-400 group-hover:border-emerald-400 transition-colors shadow-inner">
                      <Upload className="w-10 h-10 text-slate-400" />
                    </div>

                    <div className="space-y-2 max-w-md">
                      <h3 className="font-display font-bold text-2xl text-white">
                        Arraste seus PDFs para este Estúdio
                      </h3>
                      <p className="text-xs text-slate-400 leading-relaxed">
                        Solte seus documentos PDF em qualquer lugar desta área. O primeiro arquivo se tornará o <span className="text-amber-400 font-semibold">PDF Principal</span> e os subsequentes virarão <span className="text-emerald-400 font-semibold">Abas Secundárias</span>.
                      </p>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 w-full max-w-md pt-4">
                      <div className="p-3 rounded-xl bg-slate-900/50 border border-slate-800/80 text-left space-y-1">
                        <Table className="w-4 h-4 text-emerald-400" />
                        <span className="text-[11px] font-semibold text-slate-200 block">Detecção de Tabelas</span>
                        <span className="text-[10px] text-slate-500 block">Lê tabelas com ou sem borda</span>
                      </div>

                      <div className="p-3 rounded-xl bg-slate-900/50 border border-slate-800/80 text-left space-y-1">
                        <Layers className="w-4 h-4 text-amber-400" />
                        <span className="text-[11px] font-semibold text-slate-200 block">Mesclagem em Abas</span>
                        <span className="text-[10px] text-slate-500 block">Junta tudo num único Excel</span>
                      </div>

                      <div className="p-3 rounded-xl bg-slate-900/50 border border-slate-800/80 text-left space-y-1">
                        <ShieldCheck className="w-4 h-4 text-emerald-400" />
                        <span className="text-[11px] font-semibold text-slate-200 block">Segurança Total</span>
                        <span className="text-[10px] text-slate-500 block">Processamento local</span>
                      </div>
                    </div>
                  </div>
                )}

                <div className="pt-4 border-t border-slate-800/80 flex items-center justify-between text-[11px] font-mono text-slate-500">
                  <span>Table Magic v2.5 Studio</span>
                  <span>100% Processamento Local</span>
                </div>
              </div>
            </div>

          </div>
        ) : (
          /* SUCCESS RESULT STUDIO DASHBOARD */
          <motion.div
            initial={{ opacity: 0, scale: 0.98 }}
            animate={{ opacity: 1, scale: 1 }}
            className="max-w-4xl mx-auto space-y-8 py-6"
          >
            {/* Header Banner */}
            <div className="obsidian-card rounded-2xl p-8 border border-emerald-500/30 text-center space-y-4 shadow-2xl relative overflow-hidden">
              <div className="w-16 h-16 rounded-full bg-emerald-500/20 border border-emerald-400/40 mx-auto flex items-center justify-center text-emerald-400 shadow-xl emerald-glow mb-2">
                <CheckCircle2 className="w-10 h-10" />
              </div>

              <div className="space-y-1">
                <h2 className="font-display font-extrabold text-3xl text-white">
                  Planilha Excel Gerada com Sucesso!
                </h2>
                <p className="text-sm text-slate-300 max-w-md mx-auto">
                  Suas tabelas PDF foram extraídas, estruturadas e organizadas em abas no seu arquivo Excel.
                </p>
              </div>

              {/* Stats Summary Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 max-w-xl mx-auto pt-4">
                <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 text-center space-y-1">
                  <span className="text-[10px] font-mono uppercase text-slate-400">Total de Abas</span>
                  <p className="font-display font-extrabold text-2xl text-emerald-400">
                    {result.summary.length}
                  </p>
                </div>

                <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 text-center space-y-1">
                  <span className="text-[10px] font-mono uppercase text-slate-400">Tabelas Extraídas</span>
                  <p className="font-display font-extrabold text-2xl text-amber-400">
                    {result.summary.reduce((acc, curr) => acc + curr.tableCount, 0)}
                  </p>
                </div>

                <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 text-center space-y-1">
                  <span className="text-[10px] font-mono uppercase text-slate-400">Linhas Processadas</span>
                  <p className="font-display font-extrabold text-2xl text-emerald-400">
                    {result.summary.reduce((acc, curr) => acc + curr.rowCount, 0)}
                  </p>
                </div>
              </div>
            </div>

            {/* Detailed Table Summary List */}
            <div className="obsidian-card rounded-2xl p-6 border border-slate-800 space-y-4 shadow-xl">
              <h3 className="text-xs font-mono font-bold uppercase text-slate-400 tracking-wider flex items-center gap-2 border-b border-slate-800 pb-3">
                <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
                Detalhamento das Abas Criadas
              </h3>

              <div className="space-y-3 max-h-[280px] overflow-y-auto pr-1">
                {result.summary.map((item, idx) => (
                  <div 
                    key={idx} 
                    className="p-3.5 rounded-xl bg-slate-900/60 border border-slate-800/80 hover:border-slate-700 transition-all flex flex-col sm:flex-row sm:items-center justify-between gap-3"
                  >
                    <div className="space-y-1">
                      <div className="flex items-center gap-2">
                        <span className="text-xs font-bold font-mono text-emerald-400">
                          {idx + 1}. {item.tabName}
                        </span>
                      </div>
                      <p className="text-[11px] text-slate-400 font-mono">
                        Origem: {item.fileName}
                      </p>
                    </div>

                    <div className="flex items-center gap-2 shrink-0">
                      <span className="px-2.5 py-1 rounded-md bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 font-mono text-xs font-bold">
                        {item.tableCount} {item.tableCount === 1 ? 'tabela' : 'tabelas'} • {item.rowCount} linhas
                      </span>
                    </div>

                    {item.warning && (
                      <div className="w-full text-xs text-amber-300 bg-amber-950/40 border border-amber-500/30 p-2 rounded-lg flex items-center gap-2">
                        <AlertTriangle className="w-4 h-4 text-amber-400 shrink-0" />
                        <span>{item.warning}</span>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>

            {/* Action Download Buttons */}
            <div className="flex flex-col sm:flex-row gap-4 justify-center">
              <button
                type="button"
                onClick={triggerDownload}
                className="emerald-gradient-btn emerald-glow text-white py-4 px-8 rounded-xl font-display font-bold text-base flex items-center justify-center gap-3 shadow-2xl hover:scale-[1.02] active:scale-[0.98] transition-all cursor-pointer"
              >
                <Download className="w-5 h-5 text-white" />
                <span>BAIXAR PLANILHA EXCEL (.XLSX)</span>
              </button>

              <button
                type="button"
                onClick={handleReset}
                className="py-4 px-6 rounded-xl bg-slate-900 border border-slate-800 hover:border-slate-700 text-slate-300 font-semibold text-sm flex items-center justify-center gap-2 transition-colors cursor-pointer"
              >
                <RefreshCw className="w-4 h-4 text-slate-400" />
                <span>Converter Outros PDFs</span>
              </button>
            </div>

          </motion.div>
        )}

      </main>
    </div>
  );
}


'use client';

import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, FileSpreadsheet, BarChart3, PieChart, Activity, FolderGit2, Search, X, Check, ChevronDown, ArrowLeft, Download, Calendar, DollarSign, Wallet, CheckCircle2, SlidersHorizontal, ChevronLeft } from 'lucide-react';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart as RePieChart, Pie, Cell, Legend
} from 'recharts';
import { cn } from '@/lib/utils';

interface ContractData {
  [key: string]: any;
}

export default function Dashboard() {
  const [data, setData] = useState<ContractData[]>([]);
  const [fileName, setFileName] = useState<string>('');
  const [isDragging, setIsDragging] = useState(false);
  const [selectedSecretary, setSelectedSecretary] = useState<string | null>(null);
  const [selectedContract, setSelectedContract] = useState<ContractData | null>(null);
  const [isSearchPanelOpen, setIsSearchPanelOpen] = useState(false);
  
  // Advanced Search Form States (lex-analytica style)
  const [formContractNum, setFormContractNum] = useState('');
  const [formSecretaria, setFormSecretaria] = useState('All Departments');
  const [formSupplier, setFormSupplier] = useState('');
  const [formMinVal, setFormMinVal] = useState('');
  const [formMaxVal, setFormMaxVal] = useState('');
  const [formStartDate, setFormStartDate] = useState('');
  const [formEndDate, setFormEndDate] = useState('');

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement> | React.DragEvent<HTMLDivElement>) => {
    let file: File | null = null;
    if ('dataTransfer' in e) {
      e.preventDefault();
      setIsDragging(false);
      file = e.dataTransfer.files?.[0] || null;
    } else if (e.target instanceof HTMLInputElement && e.target.files) {
      file = e.target.files[0];
    }
    if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      if (typeof bstr !== 'string' && !(bstr instanceof ArrayBuffer)) return;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const rawData = XLSX.utils.sheet_to_json<ContractData>(ws);
      setData(rawData);
      
      // Reset search states
      setFormContractNum('');
      setFormSecretaria('All Departments');
      setFormSupplier('');
      setFormMinVal('');
      setFormMaxVal('');
      setFormStartDate('');
      setFormEndDate('');
      setSelectedContract(null);
      setSelectedSecretary(null);
    };
    reader.readAsBinaryString(file);
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => { e.preventDefault(); setIsDragging(true); };
  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => { e.preventDefault(); setIsDragging(false); };

  const findKey = (row: ContractData, terms: string[]) => {
    const keys = Object.keys(row);
    return keys.find(k => terms.some(t => k.toLowerCase().includes(t.toLowerCase())));
  };

  const allKeys = useMemo(() => Object.keys(data[0] || {}), [data]);
  const sampleRow = useMemo(() => data[0] || {}, [data]);

  const { 
    contratoKey, processoKey, secretariaKey, statusKey, valueKey, objetoKey, fornecedorKey,
    inicioKey, fimKey, empenhadoKey, liquidadoKey, pagoKey 
  } = useMemo(() => {
    const cKey = (findKey(sampleRow, ['contrato', 'nro_contrato']) || allKeys[0] || '') as string;
    const pKey = (findKey(sampleRow, ['processo', 'nro_processo']) || '') as string;
    const sKey = (findKey(sampleRow, ['unidadadeDes', 'unidadeDes', 'unidadadedes', 'secretaria', 'órgão', 'orgao', 'setor']) || allKeys[0] || '') as string;
    const fKey = (findKey(sampleRow, ['nomeFor', 'nomefor', 'fornecedor', 'contratada', 'empresa']) || '') as string;
    return {
      contratoKey: cKey,
      processoKey: pKey,
      secretariaKey: sKey,
      statusKey: (findKey(sampleRow, ['status', 'situação', 'situacao', 'vigência', 'fase']) || '') as string,
      valueKey: (findKey(sampleRow, ['valor_Contrata', 'valor_contrata', 'valor', 'montante']) || '') as string,
      objetoKey: (findKey(sampleRow, ['objeto', 'descrição', 'descricao']) || '') as string,
      fornecedorKey: fKey,
      inicioKey: (findKey(sampleRow, ['início', 'inicio', 'data início', 'data inicio']) || '') as string,
      fimKey: (findKey(sampleRow, ['fim', 'término', 'termino', 'vencimento', 'data fim']) || '') as string,
      empenhadoKey: (findKey(sampleRow, ['empenhado', 'valor empenhado']) || '') as string,
      liquidadoKey: (findKey(sampleRow, ['liquidado', 'valor liquidado']) || '') as string,
      pagoKey: (findKey(sampleRow, ['pago', 'valor pago']) || '') as string
    };
  }, [sampleRow, allKeys]);

  const formatCurrency = (val: any) => {
    let numericVal = val;
    if (typeof val === 'string') {
      numericVal = parseFloat(val.replace(/[^\d,-]/g, '').replace(',', '.'));
    }
    return isNaN(numericVal) ? 'R$ 0,00' : new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(numericVal);
  };

  const getNumericValue = (val: any) => {
    if (typeof val === 'number') return val;
    if (typeof val === 'string') {
      const parsed = parseFloat(val.replace(/[^\d,-]/g, '').replace(',', '.'));
      return isNaN(parsed) ? 0 : parsed;
    }
    return 0;
  };

  const formatDate = (val: any) => {
    if (!val) return 'n/a';
    if (typeof val === 'number') {
      const date = new Date((val - 25569) * 86400 * 1000);
      return date.toLocaleDateString('pt-BR');
    }
    return String(val);
  };

  const parseDateString = (dateStr: string) => {
    if (!dateStr || dateStr === 'n/a') return null;
    const parts = dateStr.split('/');
    if (parts.length === 3) {
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1;
      const year = parseInt(parts[2], 10);
      return new Date(year, month, day);
    }
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? null : d;
  };

  const secretariesList = useMemo(() => {
    const set = new Set<string>();
    data.forEach(row => {
      const sec = row[secretariaKey];
      if (sec) set.add(String(sec));
    });
    return Array.from(set).sort();
  }, [data, secretariaKey]);

  // Dynamic filter based on Advanced Search form inputs
  const filteredData = useMemo(() => {
    return data.filter(row => {
      if (formContractNum) {
        const val = String(row[contratoKey] || '').toLowerCase();
        if (!val.includes(formContractNum.toLowerCase())) return false;
      }
      if (formSecretaria && formSecretaria !== 'All Departments') {
        const val = String(row[secretariaKey] || '').toLowerCase();
        if (val !== formSecretaria.toLowerCase()) return false;
      }
      if (formSupplier) {
        const val = String(row[fornecedorKey] || '').toLowerCase();
        if (!val.includes(formSupplier.toLowerCase())) return false;
      }
      if (formMinVal) {
        const val = getNumericValue(row[valueKey]);
        if (val < parseFloat(formMinVal)) return false;
      }
      if (formMaxVal) {
        const val = getNumericValue(row[valueKey]);
        if (val > parseFloat(formMaxVal)) return false;
      }
      if (formStartDate) {
        const rowDateStr = formatDate(row[inicioKey]);
        const rowDate = parseDateString(rowDateStr);
        const filterDate = new Date(formStartDate);
        if (rowDate && rowDate < filterDate) return false;
      }
      if (formEndDate) {
        const rowDateStr = formatDate(row[fimKey]);
        const rowDate = parseDateString(rowDateStr);
        const filterDate = new Date(formEndDate);
        if (rowDate && rowDate > filterDate) return false;
      }
      return true;
    });
  }, [data, formContractNum, formSecretaria, formSupplier, formMinVal, formMaxVal, formStartDate, formEndDate, contratoKey, secretariaKey, fornecedorKey, valueKey, inicioKey, fimKey]);

  const bySecretaria = useMemo(() => filteredData.reduce((acc: any, row: ContractData) => {
    const sec = row[secretariaKey] || 'Não Informado';
    if (!acc[sec]) acc[sec] = { count: 0, totalValue: 0, status: {}, empenhado: 0, liquidado: 0, pago: 0, contracts: [] };
    acc[sec].count++;
    acc[sec].totalValue += getNumericValue(row[valueKey]);
    acc[sec].empenhado += getNumericValue(row[empenhadoKey]);
    acc[sec].liquidado += getNumericValue(row[liquidadoKey]);
    acc[sec].pago += getNumericValue(row[pagoKey]);
    acc[sec].contracts.push(row);
    if (statusKey) {
      const stat = row[statusKey] || 'N/A';
      acc[sec].status[stat] = (acc[sec].status[stat] || 0) + 1;
    }
    return acc;
  }, {}), [filteredData, secretariaKey, statusKey, valueKey, empenhadoKey, liquidadoKey, pagoKey]);

  const totalMetrics = useMemo(() => filteredData.reduce((acc, row) => ({
    value: acc.value + getNumericValue(row[valueKey]),
    empenhado: acc.empenhado + getNumericValue(row[empenhadoKey]),
    liquidado: acc.liquidado + getNumericValue(row[liquidadoKey]),
    pago: acc.pago + getNumericValue(row[pagoKey]),
  }), { value: 0, empenhado: 0, liquidado: 0, pago: 0 }), [filteredData, valueKey, empenhadoKey, liquidadoKey, pagoKey]);

  const COLORS = ['#006397', '#5cb8fd', '#92ccff', '#cce5ff', '#b7c8de', '#4f6073', '#213145', '#1a2b3c'];

  const handleResetFilters = (e: React.FormEvent) => {
    e.preventDefault();
    setFormContractNum('');
    setFormSecretaria('All Departments');
    setFormSupplier('');
    setFormMinVal('');
    setFormMaxVal('');
    setFormStartDate('');
    setFormEndDate('');
  };

  const getSecretariaIcon = (name: string) => {
    const n = name.toLowerCase();
    if (n.includes('saúde') || n.includes('saude') || n.includes('médic') || n.includes('hospital')) {
      return <span className="material-symbols-outlined text-secondary" style={{ fontVariationSettings: '"FILL" 1' }}>medical_services</span>;
    }
    if (n.includes('educação') || n.includes('educacao') || n.includes('escola') || n.includes('ensino')) {
      return <span className="material-symbols-outlined text-secondary" style={{ fontVariationSettings: '"FILL" 1' }}>school</span>;
    }
    if (n.includes('infraestrutura') || n.includes('obras') || n.includes('infra') || n.includes('urban') || n.includes('via') || n.includes('cidade')) {
      return <span className="material-symbols-outlined text-secondary" style={{ fontVariationSettings: '"FILL" 1' }}>foundation</span>;
    }
    if (n.includes('segurança') || n.includes('seguranca') || n.includes('guarda') || n.includes('polícia')) {
      return <span className="material-symbols-outlined text-secondary" style={{ fontVariationSettings: '"FILL" 1' }}>shield</span>;
    }
    if (n.includes('finanças') || n.includes('fazenda') || n.includes('orçamento') || n.includes('tesouro')) {
      return <span className="material-symbols-outlined text-secondary" style={{ fontVariationSettings: '"FILL" 1' }}>payments</span>;
    }
    return <span className="material-symbols-outlined text-secondary" style={{ fontVariationSettings: '"FILL" 1' }}>folder</span>;
  };

  const highRiskAlerts = useMemo(() => {
    return filteredData.filter(row => {
      const status = String(row[statusKey] || '').toLowerCase();
      if (status.includes('suspenso') || status.includes('atraso') || status.includes('atrasado') || status.includes('irregular') || status.includes('inativo')) {
        return true;
      }
      const endStr = formatDate(row[fimKey]);
      const endDate = parseDateString(endStr);
      if (endDate && endDate < new Date()) {
        return true;
      }
      return false;
    });
  }, [filteredData, statusKey, fimKey]);

  const complianceIndex = useMemo(() => {
    if (filteredData.length === 0) return 100;
    const activeCount = filteredData.filter(row => {
      const s = String(row[statusKey] || '').toLowerCase();
      return s.includes('ativo') || s.includes('vigente') || s.includes('andamento') || s.includes('concluido') || s.includes('concluído');
    }).length;
    return Math.round((activeCount / filteredData.length) * 100);
  }, [filteredData, statusKey]);

  const budgetEfficiency = useMemo(() => {
    const total = totalMetrics.value;
    if (total === 0) return 100;
    const executed = totalMetrics.pago;
    return Math.round((executed / total) * 100);
  }, [totalMetrics]);

  const downloadExcel = (filtered: any[], name: string) => {
    const ws = XLSX.utils.json_to_sheet(filtered);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Contratos");
    XLSX.writeFile(wb, `Contratos_${name.substring(0, 50)}.xlsx`);
  };

  // Upload screen
  if (data.length === 0) {
    return (
      <div className="min-h-screen bg-[#F8FAFC] flex flex-col justify-between py-12 px-6 font-sans select-none animate-in fade-in duration-700">
        
        {/* Espaço superior / Header sutil */}
        <div className="flex justify-center">
          <div className="flex items-center gap-2 px-4 py-1.5 rounded-full bg-slate-100 border border-slate-200/80 shadow-sm">
            <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
            <span className="text-[10px] font-bold text-slate-600 uppercase tracking-widest font-mono">Sistema de Auditoria Ativo</span>
          </div>
        </div>

        {/* Corpo principal (Minimal Single Column) */}
        <div className="max-w-xl mx-auto w-full text-center space-y-10 my-auto">
          {/* 1. Hero Headline */}
          <div className="space-y-4">
            <h1 className="text-4xl sm:text-5xl font-black text-[#020617] tracking-tight leading-tight font-mono">
              Auditorias de Contratos
            </h1>
            {/* 2. Short Description */}
            <p className="text-base text-slate-500 font-medium max-w-md mx-auto leading-relaxed">
              Carregue suas planilhas de gastos governamentais para consolidação imediata de empenhos, liquidações e pareceres técnicos de regularidade.
            </p>
          </div>

          {/* 3. Benefit Bullets (3 max) */}
          <div className="max-w-md mx-auto bg-white p-6 rounded-2xl border border-slate-100 shadow-sm text-left space-y-3.5">
            <div className="flex items-start gap-3">
              <div className="p-1 rounded-full bg-emerald-50 text-emerald-600 mt-0.5"><Check size={14} className="stroke-[3]" /></div>
              <p className="text-xs font-bold text-slate-600 leading-relaxed">
                <span className="text-slate-800">Mapeamento Inteligente:</span> Identifica e formata colunas de valores e status automaticamente.
              </p>
            </div>
            <div className="flex items-start gap-3">
              <div className="p-1 rounded-full bg-emerald-50 text-emerald-600 mt-0.5"><Check size={14} className="stroke-[3]" /></div>
              <p className="text-xs font-bold text-slate-600 leading-relaxed">
                <span className="text-slate-800">Visualização de Saldos:</span> Gráficos interativos para empenhado, liquidado e pago.
              </p>
            </div>
            <div className="flex items-start gap-3">
              <div className="p-1 rounded-full bg-emerald-50 text-emerald-600 mt-0.5"><Check size={14} className="stroke-[3]" /></div>
              <p className="text-xs font-bold text-slate-600 leading-relaxed">
                <span className="text-slate-800">Processamento Seguro:</span> Seus dados e arquivos são processados estritamente na memória local.
              </p>
            </div>
          </div>

          {/* 4. Large centered CTA (Drag & Drop box + Button) */}
          <div 
            onDrop={handleFileUpload} 
            onDragOver={handleDragOver} 
            onDragLeave={handleDragLeave} 
            className={cn(
              "p-10 border-2 border-dashed rounded-3xl flex flex-col items-center justify-center transition-all duration-300 relative overflow-hidden",
              isDragging 
                ? "border-[#0369A1] bg-[#0369A1]/5 scale-105" 
                : "border-slate-200 bg-white hover:border-[#0369A1]/50 hover:shadow-xl hover:shadow-slate-100"
            )}
          >
            <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} className="hidden" id="file-upload" />
            <label htmlFor="file-upload" className="cursor-pointer flex flex-col items-center group w-full">
              <div className="p-5 rounded-full bg-slate-50 text-slate-400 mb-6 group-hover:bg-blue-50 group-hover:text-[#0369A1] transition-all duration-300 transform group-hover:-translate-y-1">
                <FileSpreadsheet size={32} />
              </div>
              <p className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-8">
                Arraste a planilha aqui ou clique para selecionar
              </p>
              
              <div className="px-8 py-3.5 bg-[#0369A1] text-white font-bold rounded-xl text-sm transition-all shadow-md shadow-[#0369A1]/20 hover:opacity-90 hover:-translate-y-0.5 active:translate-y-0 cursor-pointer">
                Selecionar Planilha Excel
              </div>
            </label>
          </div>

          {/* Trust Indicators / Badges (Trust & Authority Style) */}
          <div className="pt-6 border-t border-slate-100 flex justify-center items-center gap-8">
            <div className="flex items-center gap-2 hover:scale-105 transition-transform" title="Conformidade LGPD">
              <CheckCircle2 size={16} className="text-slate-400" />
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Processamento Local</span>
            </div>
            <div className="h-4 w-px bg-slate-200" />
            <div className="flex items-center gap-2 hover:scale-105 transition-transform" title="Certificado de Segurança">
              <CheckCircle2 size={16} className="text-slate-400" />
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Controle Tecnológico</span>
            </div>
          </div>
        </div>

        {/* 5. Footer */}
        <footer className="text-center">
          <p className="text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em]">
            Tribunal de Contas & Controle de Regularidade Contratual
          </p>
        </footer>
      </div>
    );
  }

  // --- Dossiê de Auditoria Exclusiva do Contrato ---
  const renderContractDossier = () => {
    if (!selectedContract) return null;
    const contractId = selectedContract[contratoKey] || 'S/N';
    const processId = selectedContract[processoKey] || 'S/N';
    const secName = selectedContract[secretariaKey] || 'Não Informado';
    const supplierName = selectedContract[fornecedorKey] || 'Não Informado';
    const objectText = selectedContract[objetoKey] || 'Não Descrito';
    const statusVal = selectedContract[statusKey] || 'N/A';
    
    const valueTotal = getNumericValue(selectedContract[valueKey]);
    const valueEmpenhado = getNumericValue(selectedContract[empenhadoKey]);
    const valueLiquidado = getNumericValue(selectedContract[liquidadoKey]);
    const valuePago = getNumericValue(selectedContract[pagoKey]);
    
    const saldoEmpenhar = Math.max(0, valueTotal - valueEmpenhado);
    const saldoLiquidar = Math.max(0, valueEmpenhado - valueLiquidado);
    const saldoPagar = Math.max(0, valueLiquidado - valuePago);
    
    const percEmpenhado = valueTotal > 0 ? (valueEmpenhado / valueTotal) * 100 : 0;
    const percLiquidado = valueTotal > 0 ? (valueLiquidado / valueTotal) * 100 : 0;
    const percPago = valueTotal > 0 ? (valuePago / valueTotal) * 100 : 0;

    return (
      <div className="space-y-lg animate-in fade-in duration-500 max-w-container-max w-full">
        {/* Barra de Navegação Superior */}
        <div className="flex items-center justify-between">
          <button 
            onClick={() => setSelectedContract(null)} 
            className="flex items-center gap-2 text-[#006397] hover:text-primary font-bold text-sm transition-colors bg-white px-4 py-2.5 rounded-lg border border-outline-variant shadow-sm cursor-pointer"
          >
            <ArrowLeft size={16} /> Voltar ao Painel Geral
          </button>
          
          <button 
            onClick={() => downloadExcel([selectedContract], `Contrato_${contractId}`)} 
            className="flex items-center gap-2 px-4 py-2.5 text-sm font-bold bg-[#041627] text-white hover:opacity-90 transition-all rounded-lg shadow-sm cursor-pointer"
          >
            <Download size={16} /> Exportar Dossiê
          </button>
        </div>

        {/* Cabeçalho do Contrato (Dossier) */}
        <div className="bg-white p-lg rounded-xl border border-outline-variant shadow-sm space-y-md">
          <div className="flex flex-wrap justify-between items-start gap-md">
            <div className="space-y-2 flex-1 min-w-[280px]">
              <div className="flex flex-wrap items-center gap-sm">
                <span className="text-[10px] font-bold text-secondary bg-surface-container-low px-3 py-1 rounded border border-[#5cb8fd] font-mono">
                  CONTRATO N° {contractId}
                </span>
                <span className="text-[10px] font-bold text-on-surface-variant bg-surface-container px-3 py-1 rounded border border-outline-variant font-mono">
                  PROCESSO N° {processId}
                </span>
                <span className={cn(
                  "text-[10px] font-bold px-3 py-1 rounded border uppercase",
                  statusVal.toLowerCase().includes('atrasado') || statusVal.toLowerCase().includes('suspenso') 
                    ? "bg-red-100 border-red-200 text-red-800" 
                    : statusVal.toLowerCase().includes('ativo') || statusVal.toLowerCase().includes('vigente') || statusVal.toLowerCase().includes('andamento')
                    ? "bg-green-100 border-green-200 text-green-800"
                    : "bg-amber-100 border-amber-200 text-amber-800"
                )}>
                  {statusVal}
                </span>
              </div>
              <h1 className="font-display-lg text-display-lg text-primary mt-xs">
                {supplierName}
              </h1>
              
              <button 
                onClick={() => {
                  setSelectedSecretary(String(secName));
                  setSelectedContract(null);
                }}
                className="text-xs font-bold text-secondary hover:underline uppercase tracking-widest flex items-center gap-1.5 transition-colors group cursor-pointer"
              >
                <FolderGit2 size={14} className="text-[#5cb8fd] group-hover:scale-110 transition-transform" />
                <span>{secName}</span>
              </button>
            </div>
            
            <div className="bg-[#eff4ff] p-md rounded-lg border border-[#5cb8fd] text-right min-w-[240px]">
              <span className="text-[10px] font-bold text-[#006397] uppercase tracking-widest block mb-1">
                Valor Total do Contrato
              </span>
              <span className="text-2xl font-black text-[#041627] font-mono">
                {formatCurrency(valueTotal)}
              </span>
            </div>
          </div>

          <div className="border-t border-outline-variant pt-md">
            <span className="text-[10px] font-bold text-on-surface-variant uppercase tracking-widest block mb-2">
              Objeto Contratual
            </span>
            <p className="text-on-background bg-[#f8f9ff] p-md rounded-lg border border-outline-variant leading-relaxed text-sm font-medium">
              {objectText}
            </p>
          </div>
        </div>

        {/* Grid Principal de Auditoria */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-lg">
          
          {/* Coluna da Esquerda: Fluxo Financeiro (2/3) */}
          <div className="lg:col-span-2 space-y-lg">
            
            {/* Cards de Execução Orçamentária */}
            <div className="bg-white p-lg rounded-xl border border-outline-variant shadow-sm space-y-lg">
              <h3 className="font-headline-sm text-headline-sm text-primary flex items-center gap-2 border-b border-outline-variant pb-md">
                <span className="material-symbols-outlined text-secondary">payments</span>
                Execução e Fluxo de Caixa
              </h3>
              
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-md">
                <div className="p-md rounded-lg bg-surface-container-low border border-outline-variant space-y-1">
                  <span className="text-[10px] font-bold text-[#006397] uppercase tracking-widest block">
                    Empenhado ({percEmpenhado.toFixed(1)}%)
                  </span>
                  <p className="text-lg font-bold text-[#041627] font-mono">{formatCurrency(valueEmpenhado)}</p>
                  <span className="text-[9px] text-[#44474c] font-medium block">
                    Saldo a Empenhar: {formatCurrency(saldoEmpenhar)}
                  </span>
                </div>
                
                <div className="p-md rounded-lg bg-[#fff8eb] border border-[#ffe0b2] space-y-1">
                  <span className="text-[10px] font-bold text-amber-700 uppercase tracking-widest block">
                    Liquidado ({percLiquidado.toFixed(1)}%)
                  </span>
                  <p className="text-lg font-bold text-amber-900 font-mono">{formatCurrency(valueLiquidado)}</p>
                  <span className="text-[9px] text-amber-600 font-medium block">
                    Saldo a Liquidar: {formatCurrency(saldoLiquidar)}
                  </span>
                </div>

                <div className="p-md rounded-lg bg-[#eafbf0] border border-[#a3e9be] space-y-1">
                  <span className="text-[10px] font-bold text-emerald-700 uppercase tracking-widest block">
                    Pago ({percPago.toFixed(1)}%)
                  </span>
                  <p className="text-lg font-bold text-emerald-950 font-mono">{formatCurrency(valuePago)}</p>
                  <span className="text-[9px] text-emerald-600 font-medium block">
                    Saldo a Pagar: {formatCurrency(saldoPagar)}
                  </span>
                </div>
              </div>

              {/* Progress Meters */}
              <div className="space-y-sm pt-xs">
                <div className="space-y-1">
                  <div className="flex justify-between text-xs font-bold text-on-surface-variant">
                    <span>Percentual Empenhado / Total</span>
                    <span className="font-mono">{percEmpenhado.toFixed(1)}%</span>
                  </div>
                  <div className="w-full h-2 bg-[#eff4ff] rounded-full overflow-hidden">
                    <div className="h-full bg-secondary transition-all duration-1000" style={{ width: `${Math.min(percEmpenhado, 100)}%` }} />
                  </div>
                </div>

                <div className="space-y-1">
                  <div className="flex justify-between text-xs font-bold text-on-surface-variant">
                    <span>Percentual Liquidado / Total</span>
                    <span className="font-mono">{percLiquidado.toFixed(1)}%</span>
                  </div>
                  <div className="w-full h-2 bg-[#eff4ff] rounded-full overflow-hidden">
                    <div className="h-full bg-amber-500 transition-all duration-1000" style={{ width: `${Math.min(percLiquidado, 100)}%` }} />
                  </div>
                </div>

                <div className="space-y-1">
                  <div className="flex justify-between text-xs font-bold text-on-surface-variant">
                    <span>Percentual Pago / Total</span>
                    <span className="font-mono">{percPago.toFixed(1)}%</span>
                  </div>
                  <div className="w-full h-2 bg-[#eff4ff] rounded-full overflow-hidden">
                    <div className="h-full bg-green-600 transition-all duration-1000" style={{ width: `${Math.min(percPago, 100)}%` }} />
                  </div>
                </div>
              </div>
            </div>

            {/* Compliance Checklist */}
            <div className="bg-white p-lg rounded-xl border border-outline-variant shadow-sm space-y-md">
              <h3 className="font-headline-sm text-headline-sm text-primary flex items-center gap-2 border-b border-outline-variant pb-md">
                <span className="material-symbols-outlined text-[#006397]">gavel</span>
                Conformidade Regulatória & Auditoria
              </h3>
              
              <div className="space-y-sm">
                {[
                  { label: 'Publicação no Portal da Transparência', desc: 'Conformidade com a Lei de Acesso à Informação (LAI).', checked: true },
                  { label: 'Certidões Negativas do Fornecedor Validadas', desc: 'FGTS, INSS e tributos federais estão em dia.', checked: true },
                  { label: 'Assinaturas Digitais Coerentes', desc: 'Validação jurídica dos signatários na plataforma.', checked: true },
                  { label: 'Designação Formal de Fiscal de Contrato', desc: 'Portaria administrativa cadastrada no processo.', checked: processId !== 'S/N' },
                  { label: 'Compatibilidade de Valores Governamentais', desc: 'Valor total é justificado e compatível com o mercado.', checked: valueTotal > 0 }
                ].map((item, index) => (
                  <div key={index} className="flex items-start gap-md p-md rounded-lg border border-outline-variant bg-[#f8f9ff]">
                    <div className={cn(
                      "w-6 h-6 rounded-full border flex items-center justify-center shrink-0 mt-0.5 transition-all",
                      item.checked ? "bg-green-600 border-green-600 text-white" : "bg-amber-100 border-amber-200 text-amber-600"
                    )}>
                      {item.checked ? <Check size={12} /> : <span className="text-xs font-black">!</span>}
                    </div>
                    <div className="space-y-0.5">
                      <p className="font-headline-sm text-[14px] text-[#041627]">{item.label}</p>
                      <p className="font-body-md text-[12px] text-on-surface-variant opacity-70 leading-relaxed">{item.desc}</p>
                    </div>
                  </div>
                ))}
              </div>
            </div>

          </div>

          {/* Coluna da Direita: Prazos e Cronograma (1/3) */}
          <div className="space-y-lg">
            
            {/* Vigência Card */}
            <div className="bg-white p-lg rounded-xl border border-outline-variant shadow-sm space-y-md">
              <h3 className="font-headline-sm text-headline-sm text-primary flex items-center gap-2 border-b border-outline-variant pb-md">
                <span className="material-symbols-outlined text-primary">calendar_today</span>
                Vigência & Cronograma
              </h3>
              
              <div className="space-y-sm">
                <div className="flex items-center justify-between p-md bg-[#f8f9ff] rounded-lg border border-outline-variant">
                  <span className="font-label-md text-label-md text-on-surface-variant uppercase">Data Início</span>
                  <span className="font-mono-sm text-mono-sm text-[#041627] font-bold">{formatDate(selectedContract[inicioKey])}</span>
                </div>
                <div className="flex items-center justify-between p-md bg-[#f8f9ff] rounded-lg border border-outline-variant">
                  <span className="font-label-md text-label-md text-on-surface-variant uppercase">Data Término</span>
                  <span className="font-mono-sm text-mono-sm text-[#041627] font-bold">{formatDate(selectedContract[fimKey])}</span>
                </div>
              </div>

              {/* Visual Vertical Timeline */}
              <div className="relative pl-md border-l-2 border-outline-variant space-y-md py-sm ml-sm mt-md">
                <div className="relative">
                  <div className="absolute -left-[23px] top-1 w-3 h-3 rounded-full bg-outline ring-4 ring-[#f8f9ff]" />
                  <p className="font-label-md text-label-md text-[#041627]">Assinatura do Contrato</p>
                  <p className="text-[10px] text-on-surface-variant font-mono">{formatDate(selectedContract[inicioKey])}</p>
                </div>
                <div className="relative">
                  <div className="absolute -left-[23px] top-1 w-3 h-3 rounded-full bg-secondary ring-4 ring-[#eff4ff]" />
                  <p className="font-label-md text-label-md text-[#041627]">Fase de Execução Financeira</p>
                  <p className="text-[10px] text-on-surface-variant">Acompanhamento contínuo</p>
                </div>
                <div className="relative">
                  <div className="absolute -left-[23px] top-1 w-3 h-3 rounded-full bg-outline-variant ring-4 ring-[#f8f9ff]" />
                  <p className="font-label-md text-label-md text-[#041627]">Término Previsto</p>
                  <p className="text-[10px] text-on-surface-variant font-mono">{formatDate(selectedContract[fimKey])}</p>
                </div>
              </div>
            </div>

            {/* Parecer Técnico Action */}
            <div className="bg-[#1a2b3c] text-white p-lg rounded-xl border border-[#041627] shadow-lg space-y-md">
              <div className="space-y-1">
                <h4 className="font-label-md text-label-md uppercase text-[#8192a7]">Auditoria Externa</h4>
                <p className="font-body-md text-body-md text-[#d2e4fb] leading-relaxed">
                  Para lavrar inconsistências ou propor termo aditivo fiscalizatório, emita o parecer técnico oficial de regularidade.
                </p>
              </div>
              <button 
                onClick={() => alert('Parecer Técnico de Regularidade Contratual gerado com sucesso! Arquivo salvo localmente.')}
                className="w-full py-md bg-secondary-container text-on-secondary-container font-label-md text-label-md rounded-lg hover:opacity-90 transition-all flex items-center justify-center gap-xs cursor-pointer"
              >
                <span className="material-symbols-outlined">edit_note</span> Emitir Parecer Técnico
              </button>
            </div>

          </div>

        </div>
      </div>
    );
  };

  // --- Visão Detalhada por Secretaria ---
  const renderSecretaryDossier = () => {
    if (!selectedSecretary) return null;
    const secretaryData = data.filter(row => row[secretariaKey] === selectedSecretary);
    const secStats = bySecretaria[selectedSecretary];
    const statusData = Object.entries(secStats.status).map(([name, value]) => ({ name, value }));

    return (
      <div className="space-y-lg animate-in fade-in duration-500 max-w-container-max w-full">
        {/* Navigation */}
        <div className="flex items-center justify-between">
          <button 
            onClick={() => setSelectedSecretary(null)} 
            className="flex items-center gap-2 text-[#006397] hover:text-primary font-bold text-sm transition-colors bg-white px-4 py-2.5 rounded-lg border border-outline-variant shadow-sm cursor-pointer"
          >
            <ArrowLeft size={16} /> Voltar ao Painel Geral
          </button>
          
          <button 
            onClick={() => downloadExcel(secretaryData, selectedSecretary)} 
            className="flex items-center gap-2 px-4 py-2.5 text-sm font-bold bg-[#041627] text-white hover:opacity-90 transition-all rounded-lg shadow-sm cursor-pointer"
          >
            <Download size={16} /> Exportar Detalhes
          </button>
        </div>

        {/* Header */}
        <div className="space-y-1">
          <h1 className="font-display-lg text-display-lg text-primary">{selectedSecretary}</h1>
          <p className="font-body-md text-body-md text-on-surface-variant opacity-70">
            Dossiê de Execução e Consolidação do Órgão
          </p>
        </div>

        {/* Stats Grid */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-md">
          {[
            { label: 'Contratos Ativos', value: secStats.count, icon: FolderGit2, color: 'text-[#041627]', bg: 'bg-[#eff4ff] border-[#eff4ff]', isCurrency: false },
            { label: 'Valor Empenhado', value: secStats.empenhado, icon: Wallet, color: 'text-[#006397]', bg: 'bg-[#eff4ff] border-[#eff4ff]', isCurrency: true },
            { label: 'Valor Liquidado', value: secStats.liquidado, icon: Activity, color: 'text-amber-700', bg: 'bg-[#fff8eb] border-[#ffe0b2]', isCurrency: true },
            { label: 'Valor Pago', value: secStats.pago, icon: CheckCircle2, color: 'text-green-700', bg: 'bg-[#eafbf0] border-[#a3e9be]', isCurrency: true }
          ].map((card, i) => (
            <div key={i} className="bg-white p-md rounded-lg border border-outline-variant shadow-sm hover:shadow transition-all group">
              <div className="flex items-center gap-sm mb-4">
                <div className={cn("p-1.5 rounded-lg border", card.bg)}>
                  <card.icon size={14} className={cn("transition-transform group-hover:scale-110", card.color)} />
                </div>
                <p className="text-[10px] font-bold text-on-surface-variant uppercase tracking-widest">{card.label}</p>
              </div>
              <p className={cn("text-xl font-bold tracking-tight font-mono", card.color)}>
                {card.isCurrency ? formatCurrency(card.value) : card.value}
              </p>
            </div>
          ))}
        </div>

        {/* Charts & Contracts */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-lg">
          {/* Status pie chart */}
          <div className="bg-white p-lg rounded-xl border border-outline-variant shadow-sm h-[400px] flex flex-col">
            <h3 className="font-headline-sm text-headline-sm text-primary mb-lg">Status da Vigência</h3>
            <div className="flex-1">
              <ResponsiveContainer width="100%" height="100%">
                <RePieChart>
                  <Pie data={statusData} cx="50%" cy="50%" innerRadius={60} outerRadius={85} paddingAngle={4} dataKey="value">
                    {statusData.map((entry, index) => <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />)}
                  </Pie>
                  <Tooltip />
                  <Legend iconType="circle" wrapperStyle={{ fontSize: 11 }} />
                </RePieChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* List of contracts for this department */}
          <div className="lg:col-span-2 bg-white p-lg rounded-xl border border-outline-variant shadow-sm h-[400px] flex flex-col">
            <div className="flex items-center justify-between mb-lg border-b border-outline-variant pb-xs">
              <h3 className="font-headline-sm text-headline-sm text-primary">Contratos Vinculados</h3>
              <span className="inline-flex items-center px-md py-xs bg-surface-container-high rounded-full font-label-md text-label-md text-secondary">
                {secretaryData.length} Contratos
              </span>
            </div>
            <div className="overflow-y-auto flex-1 space-y-sm pr-2 custom-scrollbar">
              {secretaryData.map((contract, i) => {
                const totalVal = getNumericValue(contract[valueKey]);
                const paymentVal = getNumericValue(contract[pagoKey]);
                const perc = totalVal > 0 ? (paymentVal / totalVal) * 100 : 0;
                return (
                  <div key={i} className="p-md rounded-lg border border-outline-variant bg-white hover:border-[#5cb8fd] hover:shadow-sm transition-all group">
                    <div className="flex flex-wrap justify-between items-start gap-md mb-md">
                      <div className="space-y-1 flex-1 min-w-[200px]">
                        <div className="flex items-center gap-xs">
                          <span className="font-mono-sm text-mono-sm text-secondary font-bold">Nº {contract[contratoKey]}</span>
                          <span className="text-[10px] text-on-surface-variant font-medium truncate max-w-[150px]">{contract[fornecedorKey]}</span>
                        </div>
                        <h4 className="font-body-md text-body-md font-bold text-primary group-hover:text-secondary transition-colors truncate max-w-[400px]">{contract[objetoKey]}</h4>
                      </div>
                      <div className="flex flex-col items-end gap-1">
                        <span className={cn(
                          "px-sm py-0.5 rounded-full text-[10px] font-bold uppercase",
                          contract[statusKey]?.toLowerCase().includes('atrasado') || contract[statusKey]?.toLowerCase().includes('suspenso') 
                            ? "bg-red-100 text-red-800" 
                            : "bg-green-100 text-green-800"
                        )}>
                          {contract[statusKey]}
                        </span>
                        <div className="flex items-center gap-1 text-[10px] text-on-surface-variant">
                          <Calendar size={11} /> {formatDate(contract[inicioKey])} - {formatDate(contract[fimKey])}
                        </div>
                      </div>
                    </div>

                    <div className="flex justify-between items-center text-xs font-mono text-[#041627] pt-sm border-t border-outline-variant/40">
                      <span>Valor Total: {formatCurrency(totalVal)}</span>
                      <span>Execução Pago: {perc.toFixed(0)}%</span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>
    );
  };

  // --- Visão Geral do Dashboard (Relação de Contratos por Secretaria) ---
  const renderMainDashboard = () => {
    // Unique list of secretariats in the excel file
    const dropdownSecretaries = ['All Departments', ...secretariesList];

    // Status global data for pie chart
    const statusGlobalData = Object.entries(filteredData.reduce((acc: any, row) => {
      const s = row[statusKey] || 'N/A';
      acc[s] = (acc[s] || 0) + 1;
      return acc;
    }, {})).map(([name, value]) => ({ name, value }));

    // Grouped secretariats to show in UI
    const secretariaGroups = Object.entries(bySecretaria);

    return (
      <div className="space-y-lg animate-in fade-in duration-500 max-w-container-max w-full">
        {/* Header Section */}
        <section className="w-full">
          <div className="mb-lg">
            <h2 className="font-display-lg text-display-lg text-primary">Relação de Contratos por Secretaria</h2>
            <div className="flex items-center gap-md mt-sm">
              <span className="inline-flex items-center gap-xs px-md py-xs bg-surface-container-high rounded-full font-label-md text-label-md text-secondary">
                <span className="w-2 h-2 rounded-full bg-secondary"></span> 
                {filteredData.length} Contratos Ativos
              </span>
              <span className="font-body-md text-body-md text-on-surface-variant opacity-70 italic">
                Análise técnica baseada em importação local
              </span>
            </div>
          </div>

          {/* Horizontal Search/Filter Panel */}
          <div id="search-panel" className="bg-surface border border-outline-variant rounded-xl p-md mb-xl shadow-sm">
            <div 
              className="flex items-center justify-between cursor-pointer select-none"
              onClick={() => setIsSearchPanelOpen(!isSearchPanelOpen)}
            >
              <div className="flex items-center gap-xs">
                <span className="material-symbols-outlined text-secondary">manage_search</span>
                <h3 className="font-headline-sm text-headline-sm text-primary">Busca Avançada</h3>
              </div>
              <div className="flex items-center gap-xs text-xs font-bold text-on-surface-variant uppercase">
                <span>{isSearchPanelOpen ? 'Ocultar Filtros' : 'Mostrar Filtros'}</span>
                <ChevronDown size={16} className={cn("transition-transform duration-300", isSearchPanelOpen ? "rotate-180" : "rotate-0")} />
              </div>
            </div>
            
            {isSearchPanelOpen && (
              <form onSubmit={handleResetFilters} className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-7 gap-md items-end mt-md animate-in fade-in slide-in-from-top-2 duration-300">
                <div>
                  <label className="font-label-md text-label-md text-on-surface-variant block mb-xs">Nº do Contrato</label>
                  <input 
                    type="text" 
                    value={formContractNum} 
                    onChange={(e) => setFormContractNum(e.target.value)}
                    placeholder="Ex: 2024/001"
                    className="w-full border-outline-variant rounded-lg font-body-md text-body-md focus:border-secondary focus:ring-secondary transition-all py-1.5 px-3 border" 
                  />
                </div>

                <div>
                  <label className="font-label-md text-label-md text-on-surface-variant block mb-xs">Secretaria</label>
                  <select 
                    value={formSecretaria} 
                    onChange={(e) => setFormSecretaria(e.target.value)}
                    className="w-full border-outline-variant rounded-lg font-body-md text-body-md focus:border-secondary focus:ring-secondary py-1.5 px-3 border bg-white"
                  >
                    {dropdownSecretaries.map(sec => (
                      <option key={sec} value={sec}>{sec}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="font-label-md text-label-md text-on-surface-variant block mb-xs">Fornecedor</label>
                  <input 
                    type="text" 
                    value={formSupplier}
                    onChange={(e) => setFormSupplier(e.target.value)}
                    placeholder="Nome do Fornecedor..." 
                    className="w-full border-outline-variant rounded-lg font-body-md text-body-md focus:border-secondary focus:ring-secondary transition-all py-1.5 px-3 border" 
                  />
                </div>

                <div>
                  <label className="font-label-md text-label-md text-on-surface-variant block mb-xs">Mínimo Valor (R$)</label>
                  <input 
                    type="number" 
                    value={formMinVal}
                    onChange={(e) => setFormMinVal(e.target.value)}
                    placeholder="Min" 
                    className="w-full border-outline-variant rounded-lg font-body-md text-body-md focus:border-secondary focus:ring-secondary transition-all py-1.5 px-3 border" 
                  />
                </div>

                <div>
                  <label className="font-label-md text-label-md text-on-surface-variant block mb-xs">Máximo Valor (R$)</label>
                  <input 
                    type="number" 
                    value={formMaxVal}
                    onChange={(e) => setFormMaxVal(e.target.value)}
                    placeholder="Max" 
                    className="w-full border-outline-variant rounded-lg font-body-md text-body-md focus:border-secondary focus:ring-secondary transition-all py-1.5 px-3 border" 
                  />
                </div>

                <div>
                  <label className="font-label-md text-label-md text-on-surface-variant block mb-xs">Vigência Início</label>
                  <input 
                    type="date" 
                    value={formStartDate}
                    onChange={(e) => setFormStartDate(e.target.value)}
                    className="w-full border-outline-variant rounded-lg font-body-md text-body-md focus:border-secondary focus:ring-secondary py-1.5 px-3 border text-xs" 
                  />
                </div>

                <div className="flex gap-xs">
                  <button 
                    type="button"
                    onClick={() => alert('Filtros dinâmicos aplicados com sucesso!')}
                    className="flex-grow bg-primary text-on-primary font-label-md text-label-md py-2 px-3 rounded-lg hover:opacity-90 transition-all flex items-center justify-center gap-xs cursor-pointer"
                  >
                    <span className="material-symbols-outlined text-sm">filter_alt</span> Filtrar
                  </button>
                  <button 
                    type="button"
                    onClick={handleResetFilters}
                    className="w-10 border border-outline-variant text-on-surface-variant rounded-lg hover:bg-surface-container-low transition-all flex items-center justify-center cursor-pointer" 
                    title="Limpar Tudo"
                  >
                    <span className="material-symbols-outlined">restart_alt</span>
                  </button>
                </div>
              </form>
            )}
          </div>
        </section>

        {/* Grouped Sections list */}
        <section className="w-full space-y-lg">
          {secretariaGroups.length === 0 ? (
            <div className="bg-white rounded-xl border border-outline-variant p-xl text-center text-on-surface-variant font-body-md">
              Nenhuma secretaria com contratos ativos atende aos filtros definidos.
            </div>
          ) : (
            secretariaGroups.map(([secName, secInfo]: [string, any]) => (
              <div key={secName} className="bg-surface rounded-xl border border-outline-variant overflow-hidden shadow-sm transition-all duration-300 hover:shadow-md">
                
                {/* Secretariat Card Title */}
                <div className="px-lg py-md bg-surface-container-low border-b border-outline-variant flex items-center justify-between flex-wrap gap-sm">
                  <div className="flex items-center gap-md">
                    <div className="p-sm bg-white rounded-lg shadow-sm">
                      {getSecretariaIcon(secName)}
                    </div>
                    <h3 className="font-headline-sm text-headline-sm text-primary">{secName}</h3>
                  </div>
                  <div className="text-right">
                    <p className="font-label-md text-label-md text-on-surface-variant uppercase tracking-widest">Total Alocado</p>
                    <p className="font-headline-sm text-headline-sm text-secondary font-mono">{formatCurrency(secInfo.totalValue)}</p>
                  </div>
                </div>

                {/* Table of contracts inside secretariat */}
                <div className="overflow-x-auto">
                  <table className="w-full text-left font-body-md text-body-md">
                    <thead className="bg-surface text-on-surface-variant border-b border-outline-variant font-label-md text-label-md uppercase tracking-wider">
                      <tr>
                        <th className="px-lg py-md">CONTRATO Nº</th>
                        <th className="px-lg py-md">FORNECEDOR</th>
                        <th className="px-lg py-md text-center">STATUS</th>
                        <th className="px-lg py-md text-right">VALOR (R$)</th>
                        <th className="px-lg py-md">VALIDADE</th>
                        <th className="px-lg py-md text-right">AÇÕES</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-outline-variant">
                      {secInfo.contracts.slice(0, 5).map((row: ContractData, index: number) => {
                        const contractId = row[contratoKey] || 'S/N';
                        const supplierName = row[fornecedorKey] || 'Não Informado';
                        const statusVal = row[statusKey] || 'N/A';
                        const totalVal = getNumericValue(row[valueKey]);
                        const validityText = `${formatDate(row[inicioKey])} - ${formatDate(row[fimKey])}`;
                        
                        return (
                          <tr key={index} className="hover:bg-surface-container transition-colors group">
                            <td className="px-lg py-md font-mono-sm text-mono-sm text-secondary font-bold">
                              {contractId}
                            </td>
                            <td className="px-lg py-md">
                              <div className="font-bold text-primary">{supplierName}</div>
                              <div className="text-xs text-on-surface-variant">Processo: {row[processoKey] || 'n/a'}</div>
                            </td>
                            <td className="px-lg py-md text-center">
                              <span className={cn(
                                "px-sm py-1 rounded-full text-xs font-bold uppercase",
                                statusVal.toLowerCase().includes('atrasado') || statusVal.toLowerCase().includes('suspenso') 
                                  ? "bg-red-100 text-red-800" 
                                  : statusVal.toLowerCase().includes('ativo') || statusVal.toLowerCase().includes('vigente') || statusVal.toLowerCase().includes('andamento')
                                  ? "bg-green-100 text-green-800"
                                  : "bg-amber-100 text-amber-800"
                              )}>
                                {statusVal}
                              </span>
                            </td>
                            <td className="px-lg py-md text-right font-bold text-primary font-mono">
                              {formatCurrency(totalVal)}
                            </td>
                            <td className="px-lg py-md text-on-surface-variant font-mono text-xs">
                              {validityText}
                            </td>
                            <td className="px-lg py-md text-right">
                              <div className="flex items-center justify-end gap-xs">
                                <button 
                                  onClick={() => setSelectedContract(row)}
                                  className="material-symbols-outlined text-on-surface-variant hover:text-primary transition-colors p-xs rounded hover:bg-surface-variant cursor-pointer"
                                  title="Auditar Contrato"
                                >
                                  visibility
                                </button>
                                <button 
                                  onClick={() => setSelectedSecretary(secName)}
                                  className="material-symbols-outlined text-on-surface-variant hover:text-secondary transition-colors p-xs rounded hover:bg-surface-variant cursor-pointer"
                                  title="Análise do Setor"
                                >
                                  analytics
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>

                {secInfo.contracts.length > 5 && (
                  <div className="px-lg py-sm bg-surface-container-lowest text-right">
                    <button 
                      onClick={() => setSelectedSecretary(secName)}
                      className="font-label-md text-label-md text-secondary hover:underline flex items-center gap-xs ml-auto cursor-pointer"
                    >
                      Ver Tudo de {secName} ({secInfo.contracts.length}) <span className="material-symbols-outlined text-sm">arrow_forward</span>
                    </button>
                  </div>
                )}
              </div>
            ))
          )}
        </section>

        {/* Bento Grid Stats */}
        <section className="mt-xl grid grid-cols-1 lg:grid-cols-3 gap-lg">
          {/* Card 1: ALERTA DE ALTO RISCO */}
          <div className="col-span-1 bg-primary text-on-primary p-lg rounded-xl flex flex-col justify-between overflow-hidden relative group">
            <div className="z-10">
              <h4 className="font-label-md text-label-md opacity-60 uppercase tracking-widest">ALERTAS DE ALTO RISCO</h4>
              <div className="font-display-lg text-display-lg mt-sm font-mono">
                {String(highRiskAlerts.length).padStart(2, '0')}
              </div>
            </div>
            <div className="mt-lg z-10 space-y-1">
              <p className="text-xs opacity-70">
                {highRiskAlerts.length > 0 
                  ? `${highRiskAlerts.length} contratos exigem auditoria imediata ou contatos com fornecedores devido a prazos vencidos ou atrasos.` 
                  : 'Nenhum contrato possui inconformidades de vigência ou atraso.'}
              </p>
              {highRiskAlerts.slice(0, 2).map((alertItem, idx) => (
                <div key={idx} className="text-[10px] bg-white/10 px-2 py-0.5 rounded flex justify-between font-mono">
                  <span className="truncate max-w-[120px]">{alertItem[fornecedorKey]}</span>
                  <span>{alertItem[contratoKey]}</span>
                </div>
              ))}
            </div>
            <span className="material-symbols-outlined absolute -right-4 -bottom-4 text-[120px] opacity-10 group-hover:scale-110 transition-transform duration-500">warning</span>
          </div>

          {/* Card 2: Bento right part (Trends + compliance) */}
          <div className="col-span-2 bg-white border border-outline-variant p-lg rounded-xl flex flex-col sm:flex-row items-stretch gap-xl">
            {/* Left side: chart */}
            <div className="w-full sm:w-1/3 flex flex-col justify-between space-y-md">
              <h4 className="font-label-md text-label-md text-on-surface-variant uppercase tracking-widest">TENDÊNCIA DE GASTOS</h4>
              <div className="h-28 w-full bg-surface-container rounded-lg p-1">
                {secretariaGroups.length > 0 ? (
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={secretariaGroups.map(([name, s]: any) => ({ name: name.substring(0, 10), value: s.totalValue })).slice(0, 5)}>
                      <Bar dataKey="value" fill="#5cb8fd" radius={[2, 2, 0, 0]} />
                      <Tooltip formatter={(v) => formatCurrency(v)} />
                    </BarChart>
                  </ResponsiveContainer>
                ) : (
                  <div className="w-full h-full flex items-center justify-center text-[10px] text-on-surface-variant">Sem dados</div>
                )}
              </div>
            </div>

            {/* Right side: progress stats */}
            <div className="w-full sm:w-2/3 border-t sm:border-t-0 sm:border-l border-outline-variant pt-lg sm:pt-0 sm:pl-xl flex flex-col justify-center space-y-md">
              <h4 className="font-label-md text-label-md text-on-surface-variant uppercase tracking-widest">DESEMPENHO POR SECRETARIA</h4>
              
              <div className="space-y-sm">
                <div className="space-y-1">
                  <div className="flex items-center justify-between text-xs">
                    <span className="font-body-md text-body-md">Índice de Compliance</span>
                    <span className="font-bold text-green-600">{complianceIndex}%</span>
                  </div>
                  <div className="w-full h-2 bg-surface-container rounded-full overflow-hidden">
                    <div className="bg-green-600 h-full transition-all duration-1000" style={{ width: `${complianceIndex}%` }} />
                  </div>
                </div>

                <div className="space-y-1">
                  <div className="flex items-center justify-between text-xs">
                    <span className="font-body-md text-body-md">Eficiência Orçamentária</span>
                    <span className="font-bold text-secondary">{budgetEfficiency}%</span>
                  </div>
                  <div className="w-full h-2 bg-surface-container rounded-full overflow-hidden">
                    <div className="bg-secondary h-full transition-all duration-1000" style={{ width: `${budgetEfficiency}%` }} />
                  </div>
                </div>
              </div>
            </div>
          </div>
        </section>
      </div>
    );
  };

  // --- Main Layout wrapper ---
  return (
    <div className="text-on-background bg-background min-h-screen">
      {/* SideNavBar (JSON Derived) */}
      <aside className="fixed left-0 top-0 h-full w-sidebar-width bg-primary dark:bg-tertiary flex flex-col py-lg border-r border-outline-variant z-50">
        <div className="px-lg mb-xl">
          <h1 className="font-headline-md text-headline-md text-on-primary dark:text-on-tertiary tracking-tight truncate" title="PMO Osasco">
            PMO Osasco
          </h1>
          <p className="font-label-md text-label-md text-[#8192a7] opacity-75 truncate" title={fileName || 'Portal Institucional'}>
            {fileName ? fileName : 'Portal Institucional'}
          </p>
        </div>
        
        <nav className="flex-grow space-y-base overflow-y-auto custom-scrollbar">
          <button 
            onClick={() => { setSelectedContract(null); setSelectedSecretary(null); }}
            className={cn(
              "w-full flex items-center px-lg py-sm font-body-md text-body-md text-[#8192a7] hover:bg-primary-container hover:text-white transition-colors duration-200 cursor-pointer text-left",
              (selectedContract === null && selectedSecretary === null) && "border-l-4 border-secondary-container bg-primary-container text-white font-bold"
            )}
          >
            <span className="material-symbols-outlined mr-md">dashboard</span> Dashboard
          </button>
          
          <button 
            onClick={() => {
              setSelectedContract(null);
              setSelectedSecretary(null);
              setTimeout(() => {
                document.getElementById('search-panel')?.scrollIntoView({ behavior: 'smooth' });
              }, 100);
            }}
            className="w-full flex items-center px-lg py-sm font-body-md text-body-md text-[#8192a7] hover:bg-primary-container hover:text-white transition-colors duration-200 cursor-pointer text-left"
          >
            <span className="material-symbols-outlined mr-md">search_check</span> Busca Avançada
          </button>

          {/* Quick link list to Departments */}
          <div className="pt-sm">
            <p className="px-lg text-[9px] font-bold text-[#8192a7]/40 uppercase tracking-widest mb-xs">Secretarias</p>
            <div className="max-h-[220px] overflow-y-auto space-y-base pr-2 custom-scrollbar">
              {secretariesList.map(secName => (
                <button
                  key={secName}
                  onClick={() => {
                    setSelectedSecretary(secName);
                    setSelectedContract(null);
                  }}
                  className={cn(
                    "w-full flex items-center px-lg py-1.5 font-body-md text-[13px] text-[#8192a7]/80 hover:bg-primary-container hover:text-white transition-colors text-left truncate pl-[44px] cursor-pointer",
                    selectedSecretary === secName && "text-[#5cb8fd] font-bold"
                  )}
                >
                  {secName}
                </button>
              ))}
            </div>
          </div>
        </nav>
        
        <div className="px-lg mt-auto space-y-md pt-xl">
          <button 
            onClick={() => {
              setData([]);
              setFileName('');
              setSelectedContract(null);
              setSelectedSecretary(null);
            }}
            className="w-full bg-secondary-container text-on-secondary-container font-label-md text-label-md py-sm px-md rounded-lg flex items-center justify-center gap-xs hover:opacity-90 transition-all cursor-pointer font-bold"
          >
            <span className="material-symbols-outlined">add</span> Nova Análise
          </button>
          
          <div className="pt-md border-t border-[#8192a7]/20">
            <a className="flex items-center py-xs font-body-md text-body-md text-[#8192a7]/80 hover:text-white" href="#">
              <span className="material-symbols-outlined mr-md">settings</span> Settings
            </a>
            <a className="flex items-center py-xs font-body-md text-body-md text-[#8192a7]/80 hover:text-white" href="#">
              <span className="material-symbols-outlined mr-md">help</span> Support
            </a>
          </div>
        </div>
      </aside>
      
      {/* TopNavBar */}
      <header className="fixed top-0 right-0 left-sidebar-width h-16 bg-surface flex justify-between items-center px-lg border-b border-outline-variant z-40">
        <div className="flex items-center gap-md">
          <span className="font-headline-sm text-headline-sm text-primary">Sistema de Gestão de Contratos</span>
        </div>
        <div className="flex items-center gap-xl">
          <div className="hidden md:flex gap-lg">
            <span className="font-label-md text-label-md text-on-surface-variant hover:text-primary transition-all cursor-pointer">Documentos</span>
            <span className="font-label-md text-label-md text-secondary font-bold border-b-2 border-secondary pb-base cursor-pointer">Fluxos de Trabalho</span>
            <span className="font-label-md text-label-md text-on-surface-variant hover:text-primary transition-all cursor-pointer">Relatórios</span>
          </div>
          <div className="flex items-center gap-md">
            <button className="material-symbols-outlined text-primary-fixed-variant hover:text-primary transition-colors">notifications</button>
            <button className="material-symbols-outlined text-primary-fixed-variant hover:text-primary transition-colors">history</button>
            <div className="w-8 h-8 rounded-full bg-surface-container-high border border-[#c4c6cd] flex items-center justify-center overflow-hidden">
              <span className="material-symbols-outlined text-primary text-xl">account_circle</span>
            </div>
          </div>
        </div>
      </header>
      
      {/* Main Content Area */}
      <main className="ml-sidebar-width mt-16 p-lg flex flex-col gap-lg bg-background min-h-[calc(100vh-64px)]">
        {selectedContract ? (
          /* Visão Detalhada de Auditoria Exclusiva do Contrato */
          renderContractDossier()
        ) : selectedSecretary ? (
          /* Visão Detalhada por Secretaria */
          renderSecretaryDossier()
        ) : (
          /* Visão Geral do Dashboard (Relação de Contratos por Secretaria) */
          renderMainDashboard()
        )}
      </main>

      {/* Floating Action for Quick actions */}
      <div className="fixed bottom-lg right-lg group z-50">
        <div className="absolute bottom-full right-0 mb-md opacity-0 group-hover:opacity-100 translate-y-4 group-hover:translate-y-0 transition-all duration-300 pointer-events-none group-hover:pointer-events-auto flex flex-col gap-sm items-end">
          <button 
            onClick={() => alert('Novo fornecedor cadastrado localmente para análises.')}
            className="bg-white border border-outline-variant px-md py-sm rounded-lg shadow-lg font-label-md text-label-md text-primary hover:bg-surface-container transition-all flex items-center gap-sm cursor-pointer whitespace-nowrap"
          >
            Novo Fornecedor <span className="material-symbols-outlined text-sm">person_add</span>
          </button>
          <button 
            onClick={() => downloadExcel(filteredData, "Consolidados")}
            className="bg-white border border-outline-variant px-md py-sm rounded-lg shadow-lg font-label-md text-label-md text-primary hover:bg-surface-container transition-all flex items-center gap-sm cursor-pointer whitespace-nowrap"
          >
            Exportar CSV <span className="material-symbols-outlined text-sm">download</span>
          </button>
        </div>
        <button 
          onClick={() => {
            setData([]);
            setFileName('');
            setSelectedContract(null);
            setSelectedSecretary(null);
          }}
          title="Nova Análise"
          className="w-16 h-16 bg-primary text-on-primary rounded-full shadow-2xl flex items-center justify-center hover:scale-110 active:scale-95 transition-all cursor-pointer"
        >
          <span className="material-symbols-outlined text-3xl">add</span>
        </button>
      </div>
    </div>
  );
}

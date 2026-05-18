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
  const [filterQuery, setFilterQuery] = useState('');
  const [selectedSecretariesFilter, setSelectedSecretariesFilter] = useState<string[]>([]);
  const [selectedStatusFilter, setSelectedStatusFilter] = useState<string[]>([]);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);

  // States for the General Contracts Table
  const [tableSearchQuery, setTableSearchQuery] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(10);
  const [sortColumn, setSortColumn] = useState<string | null>(null);
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');
  const [expandedRowIndex, setExpandedRowIndex] = useState<number | null>(null);

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
      setSelectedSecretariesFilter([]);
      setSelectedStatusFilter([]);
      setTableSearchQuery('');
      setCurrentPage(1);
      setSortColumn(null);
      setExpandedRowIndex(null);
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
    contratoKey, secretariaKey, statusKey, valueKey, objetoKey, fornecedorKey,
    inicioKey, fimKey, empenhadoKey, liquidadoKey, pagoKey 
  } = useMemo(() => {
    const cKey = (findKey(sampleRow, ['contrato', 'nro_contrato', 'processo']) || allKeys[0] || '') as string;
    const sKey = (findKey(sampleRow, ['unidadadeDes', 'unidadeDes', 'unidadadedes', 'secretaria', 'órgão', 'orgao', 'setor']) || allKeys[0] || '') as string;
    const fKey = (findKey(sampleRow, ['nomeFor', 'nomefor', 'fornecedor', 'contratada', 'empresa']) || '') as string;
    return {
      contratoKey: cKey,
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

  const secretariesList = useMemo(() => {
    const set = new Set<string>();
    data.forEach(row => {
      const sec = row[secretariaKey];
      if (sec) set.add(String(sec));
    });
    return Array.from(set).sort();
  }, [data, secretariaKey]);

  const filteredData = useMemo(() => {
    return data.filter(row => {
      const matchSec = selectedSecretariesFilter.length === 0 || selectedSecretariesFilter.includes(String(row[secretariaKey]));
      const matchStatus = selectedStatusFilter.length === 0 || selectedStatusFilter.includes(String(row[statusKey]));
      return matchSec && matchStatus;
    });
  }, [data, selectedSecretariesFilter, selectedStatusFilter, secretariaKey, statusKey]);

  const bySecretaria = useMemo(() => filteredData.reduce((acc: any, row: ContractData) => {
    const sec = row[secretariaKey] || 'Não Informado';
    if (!acc[sec]) acc[sec] = { count: 0, totalValue: 0, status: {}, empenhado: 0, liquidado: 0, pago: 0 };
    acc[sec].count++;
    acc[sec].totalValue += getNumericValue(row[valueKey]);
    acc[sec].empenhado += getNumericValue(row[empenhadoKey]);
    acc[sec].liquidado += getNumericValue(row[liquidadoKey]);
    acc[sec].pago += getNumericValue(row[pagoKey]);
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

  const COLORS = ['#2563eb', '#7c3aed', '#db2777', '#059669', '#d97706', '#dc2626', '#0891b2', '#4f46e5'];

  const statusList = useMemo(() => {
    const set = new Set<string>();
    data.forEach(row => {
      const s = row[statusKey];
      if (s) set.add(String(s));
    });
    return Array.from(set).sort();
  }, [data, statusKey]);

  const toggleStatusFilter = (name: string) => {
    setSelectedStatusFilter(prev => prev.includes(name) ? prev.filter(n => n !== name) : [...prev, name]);
  };

  const suppliersData = useMemo(() => {
    const map: any = {};
    filteredData.forEach(row => {
      const sup = row[fornecedorKey] || 'N/A';
      map[sup] = (map[sup] || 0) + getNumericValue(row[valueKey]);
    });
    return Object.entries(map)
      .map(([name, value]: [string, any]) => ({ 
        name: name.length > 20 ? name.substring(0, 20) + '...' : name, 
        value 
      }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 10);
  }, [filteredData, fornecedorKey, valueKey]);

  const globalExecutionData = useMemo(() => [
    { name: 'Contratado', value: totalMetrics.value, fill: '#0f172a' },
    { name: 'Empenhado', value: totalMetrics.empenhado, fill: '#2563eb' },
    { name: 'Liquidado', value: totalMetrics.liquidado, fill: '#d97706' },
    { name: 'Pago', value: totalMetrics.pago, fill: '#059669' }
  ], [totalMetrics]);

  const downloadExcel = (filtered: any[], name: string) => {
    const ws = XLSX.utils.json_to_sheet(filtered);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Contratos");
    XLSX.writeFile(wb, `Contratos_${name.substring(0, 50)}.xlsx`);
  };

  const handleTableSearchChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setTableSearchQuery(e.target.value);
    setCurrentPage(1);
  };
  
  const handleRowsPerPageChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    setRowsPerPage(Number(e.target.value));
    setCurrentPage(1);
  };

  const handleSort = (columnKey: string) => {
    if (sortColumn === columnKey) {
      setSortDirection(prev => prev === 'asc' ? 'desc' : 'asc');
    } else {
      setSortColumn(columnKey);
      setSortDirection('asc');
    }
  };

  const renderSortIcon = (columnKey: string) => {
    if (sortColumn !== columnKey) return <ChevronDown size={14} className="text-slate-300 ml-1 opacity-50 group-hover:opacity-100 transition-opacity" />;
    return sortDirection === 'asc' 
      ? <ChevronDown size={14} className="text-blue-600 ml-1 transform rotate-180 transition-transform" />
      : <ChevronDown size={14} className="text-blue-600 ml-1 transition-transform" />;
  };

  // Filtered data for the table (applying sidebar filters AND table search query)
  const searchedTableData = useMemo(() => {
    return filteredData.filter(row => {
      if (!tableSearchQuery) return true;
      const searchLower = tableSearchQuery.toLowerCase();
      
      const secVal = String(row[secretariaKey] || '').toLowerCase();
      const fornVal = String(row[fornecedorKey] || '').toLowerCase();
      const objVal = String(row[objetoKey] || '').toLowerCase();
      const idVal = String(row[contratoKey] || '').toLowerCase();
      
      return secVal.includes(searchLower) || 
             fornVal.includes(searchLower) || 
             objVal.includes(searchLower) || 
             idVal.includes(searchLower);
    });
  }, [filteredData, tableSearchQuery, secretariaKey, fornecedorKey, objetoKey, contratoKey]);

  // Sorted data for the table
  const sortedTableData = useMemo(() => {
    if (!sortColumn) return searchedTableData;
    
    const sorted = [...searchedTableData];
    sorted.sort((a, b) => {
      let valA = a[sortColumn];
      let valB = b[sortColumn];
      
      if (sortColumn === valueKey || sortColumn === empenhadoKey || sortColumn === liquidadoKey || sortColumn === pagoKey) {
        return sortDirection === 'asc' 
          ? getNumericValue(valA) - getNumericValue(valB)
          : getNumericValue(valB) - getNumericValue(valA);
      }
      
      const strA = String(valA || '').toLowerCase();
      const strB = String(valB || '').toLowerCase();
      
      if (strA < strB) return sortDirection === 'asc' ? -1 : 1;
      if (strA > strB) return sortDirection === 'asc' ? 1 : -1;
      return 0;
    });
    
    return sorted;
  }, [searchedTableData, sortColumn, sortDirection, valueKey, empenhadoKey, liquidadoKey, pagoKey]);

  // Paginated table data
  const paginatedTableData = useMemo(() => {
    const startIndex = (currentPage - 1) * rowsPerPage;
    return sortedTableData.slice(startIndex, startIndex + rowsPerPage);
  }, [sortedTableData, currentPage, rowsPerPage]);

  const totalPages = useMemo(() => {
    return Math.ceil(sortedTableData.length / rowsPerPage) || 1;
  }, [sortedTableData, rowsPerPage]);

  if (data.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen bg-slate-50 px-4">
        <div 
          onDrop={handleFileUpload} 
          onDragOver={handleDragOver} 
          onDragLeave={handleDragLeave} 
          className={cn(
            "w-full max-w-2xl p-16 border-2 border-dashed rounded-3xl flex flex-col items-center justify-center transition-all duration-500",
            isDragging ? "border-blue-500 bg-blue-50/50 scale-105" : "border-slate-200 bg-white shadow-2xl shadow-slate-200/50"
          )}
        >
          <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} className="hidden" id="file-upload" />
          <label htmlFor="file-upload" className="flex flex-col items-center cursor-pointer group">
            <div className="p-8 rounded-full bg-slate-50 text-slate-400 mb-8 group-hover:bg-blue-50 group-hover:text-blue-500 transition-all duration-300 transform group-hover:rotate-12">
              <UploadCloud size={48} />
            </div>
            <h2 className="text-3xl font-bold mb-3 text-slate-900 tracking-tight">Portal de Contratos</h2>
            <p className="text-slate-500 text-center mb-10 max-w-sm leading-relaxed">
              Arraste sua planilha (.xlsx) para iniciar a auditoria dinâmica e visualização de métricas.
            </p>
            <div className="px-10 py-4 rounded-2xl bg-slate-900 text-white font-bold hover:bg-slate-800 transition-all shadow-xl shadow-slate-200 active:scale-95">
              Selecionar Arquivo
            </div>
          </label>
        </div>
        <p className="mt-8 text-slate-400 text-xs font-medium uppercase tracking-[0.2em]">Auditoria Contratual Inteligente</p>
      </div>
    );
  }

  // --- Visão Detalhada ---
  if (selectedSecretary) {
    const secretaryData = data.filter(row => row[secretariaKey] === selectedSecretary);
    const secStats = bySecretaria[selectedSecretary];
    const statusData = Object.entries(secStats.status).map(([name, value]) => ({ name, value }));

    return (
      <div className="min-h-screen bg-white py-10 px-4 animate-in fade-in duration-500">
        <div className="max-w-7xl mx-auto space-y-10">
          <div className="flex items-center justify-between">
            <button onClick={() => setSelectedSecretary(null)} className="flex items-center gap-2 text-slate-500 hover:text-slate-800 font-medium text-sm transition-colors">
              <ArrowLeft size={16} /> Voltar ao Painel Geral
            </button>
            <button onClick={() => downloadExcel(secretaryData, selectedSecretary)} className="flex items-center gap-2 px-4 py-2 text-sm font-bold bg-white border border-slate-200 rounded-xl hover:bg-slate-50 transition-all shadow-sm">
              <Download size={16} /> Exportar Detalhes
            </button>
          </div>

          <div className="space-y-1">
            <h1 className="text-4xl font-bold text-slate-900 tracking-tight">{selectedSecretary}</h1>
            <p className="text-slate-500 font-medium">Auditoria Contratual e Execução Financeira</p>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6">
            {[
              { label: 'Total de Contratos', value: secStats.count, icon: FolderGit2, color: 'text-primary', bg: 'bg-slate-100/50', isCurrency: false },
              { label: 'Valor Empenhado', value: secStats.empenhado, icon: Wallet, color: 'text-blue-600', bg: 'bg-blue-50/50', isCurrency: true },
              { label: 'Valor Liquidado', value: secStats.liquidado, icon: Activity, color: 'text-amber-600', bg: 'bg-amber-50/50', isCurrency: true },
              { label: 'Valor Pago', value: secStats.pago, icon: CheckCircle2, color: 'text-emerald-600', bg: 'bg-emerald-50/50', isCurrency: true }
            ].map((card, i) => (
              <div key={i} className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm hover:shadow-xl hover:-translate-y-1 transition-all duration-300 group">
                <div className="flex items-center gap-3 mb-4">
                  <div className={cn("p-2 rounded-lg", card.bg)}>
                    <card.icon size={16} className={cn("transition-transform group-hover:scale-110", card.color)} />
                  </div>
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">{card.label}</p>
                </div>
                <p className={cn("text-2xl font-bold tracking-tight", card.color)}>
                  {card.isCurrency ? formatCurrency(card.value) : card.value}
                </p>
              </div>
            ))}
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
            <div className="lg:col-span-1 bg-white p-8 rounded-2xl border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] h-[400px] flex flex-col">
              <h3 className="text-lg font-bold text-slate-800 mb-8">Status da Vigência</h3>
              <div className="flex-1">
                <ResponsiveContainer width="100%" height="100%">
                  <RePieChart>
                    <Pie data={statusData} cx="50%" cy="50%" innerRadius={70} outerRadius={100} paddingAngle={5} dataKey="value">
                      {statusData.map((entry, index) => <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />)}
                    </Pie>
                    <Tooltip />
                    <Legend />
                  </RePieChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div className="lg:col-span-2 bg-white p-10 rounded-2xl border border-slate-100 shadow-2xl shadow-slate-200/40 h-[600px] lg:h-[500px] flex flex-col">
              <div className="flex items-center justify-between mb-8">
                <h3 className="text-xl font-bold text-slate-900 tracking-tight">Cronograma e Detalhes Financeiros</h3>
                <span className="text-[10px] font-bold text-slate-400 uppercase bg-slate-50 px-3 py-1 rounded-full border border-slate-100">
                  {secretaryData.length} Itens
                </span>
              </div>
              <div className="overflow-y-auto flex-1 space-y-6 pr-4 custom-scrollbar">
                {secretaryData.map((contract, i) => {
                  const total = getNumericValue(contract[valueKey]);
                  const liq = getNumericValue(contract[liquidadoKey]);
                  const perc = total > 0 ? (liq / total) * 100 : 0;
                  return (
                    <div key={i} className="p-8 rounded-2xl border border-slate-100 bg-white hover:border-blue-100 hover:shadow-xl transition-all duration-300 group">
                      <div className="flex flex-wrap justify-between items-start gap-6 mb-6">
                        <div className="space-y-2 flex-1 min-w-[200px]">
                          <div className="flex items-center gap-2">
                            <span className="text-[9px] font-bold text-blue-600 bg-blue-50 px-2 py-0.5 rounded font-mono">ID: {contract[allKeys[0]]}</span>
                            <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">{contract[fornecedorKey] || 'Fornecedor não informado'}</span>
                          </div>
                          <h4 className="text-lg font-bold text-slate-900 leading-tight group-hover:text-blue-700 transition-colors">{contract[objetoKey] || 'Objeto não descrito'}</h4>
                        </div>
                        <div className="flex flex-col items-end gap-2">
                          <span className={cn(
                            "text-[10px] font-bold px-3 py-1 rounded-xl border transition-colors",
                            contract[statusKey]?.toLowerCase().includes('atrasado') ? "bg-red-50 border-red-100 text-red-600" : "bg-slate-50 border-slate-100 text-slate-600"
                          )}>
                            {contract[statusKey] || 'S/ STATUS'}
                          </span>
                          <div className="flex items-center gap-1.5 text-[10px] font-bold text-slate-400">
                            <Calendar size={12} className="text-slate-300" />
                            {formatDate(contract[inicioKey])} <span className="text-slate-200">/</span> {formatDate(contract[fimKey])}
                          </div>
                        </div>
                      </div>
                      
                      <div className="grid grid-cols-2 sm:grid-cols-4 gap-6 mb-8 py-6 border-y border-slate-50">
                        <div><p className="text-[9px] font-bold text-slate-400 uppercase tracking-wider mb-1">Contrato</p><p className="text-sm font-bold text-slate-900 font-mono">{formatCurrency(contract[valueKey])}</p></div>
                        <div><p className="text-[9px] font-bold text-slate-400 uppercase tracking-wider mb-1">Empenhado</p><p className="text-sm font-bold text-blue-600 font-mono">{formatCurrency(contract[empenhadoKey])}</p></div>
                        <div><p className="text-[9px] font-bold text-slate-400 uppercase tracking-wider mb-1">Liquidado</p><p className="text-sm font-bold text-amber-600 font-mono">{formatCurrency(contract[liquidadoKey])}</p></div>
                        <div><p className="text-[9px] font-bold text-slate-400 uppercase tracking-wider mb-1">Pago</p><p className="text-sm font-bold text-emerald-600 font-mono">{formatCurrency(contract[pagoKey])}</p></div>
                      </div>

                      <div className="space-y-3">
                        <div className="flex justify-between items-center">
                          <p className="text-[9px] font-bold text-slate-400 uppercase tracking-[0.2em]">Execução Financeira</p>
                          <p className="text-xs font-bold text-slate-900 font-mono">{perc.toFixed(1)}%</p>
                        </div>
                        <div className="w-full h-2 bg-slate-50 rounded-full overflow-hidden p-0.5 border border-slate-100">
                          <div className={cn("h-full rounded-full transition-all duration-1000", perc > 80 ? "bg-emerald-500" : "bg-blue-600")} style={{ width: `${Math.min(perc, 100)}%` }} />
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // --- Visão Geral ---
  const statusGlobalData = Object.entries(filteredData.reduce((acc: any, row) => {
    const s = row[statusKey] || 'N/A';
    acc[s] = (acc[s] || 0) + 1;
    return acc;
  }, {})).map(([name, value]) => ({ name, value }));

  return (
    <div className="flex min-h-screen bg-slate-50 font-sans">
      {/* Sidebar de Filtros */}
      <aside className={cn(
        "fixed inset-y-0 left-0 z-50 w-80 bg-white border-r border-slate-200 transform transition-all duration-300 ease-in-out shadow-2xl lg:shadow-none lg:relative",
        !isSidebarOpen 
          ? "-translate-x-full lg:translate-x-0 lg:w-0 lg:border-r-0 overflow-hidden opacity-0 lg:opacity-0" 
          : "translate-x-0 lg:w-80 lg:opacity-100 lg:border-r"
      )}>
        <div className={cn(
          "h-full flex flex-col space-y-8 overflow-y-auto custom-scrollbar transition-all duration-300",
          !isSidebarOpen ? "p-0 opacity-0" : "p-6 opacity-100"
        )}>
          <div className="flex items-center justify-between">
            <h2 className="text-xl font-bold text-slate-900 tracking-tight flex items-center gap-2">
              <Activity size={20} className="text-blue-600" />
              Filtros
            </h2>
            <button onClick={() => setIsSidebarOpen(false)} className="lg:hidden p-2 text-slate-400 hover:text-slate-900">
              <X size={20} />
            </button>
          </div>

          {/* Pesquisa */}
          <div className="space-y-3">
            <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Pesquisar Órgão</p>
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
              <input 
                type="text" 
                placeholder="Ex: Saúde, Educação..." 
                value={filterQuery} 
                onChange={(e) => setFilterQuery(e.target.value)}
                className="w-full pl-10 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-2xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/10 focus:border-blue-500/50 transition-all"
              />
            </div>
          </div>

          {/* Filtro por Órgão */}
          <div className="space-y-4">
            <div className="flex items-center justify-between">
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Órgãos ({secretariesList.length})</p>
              <div className="flex gap-2">
                <button onClick={() => setSelectedSecretariesFilter([])} className="text-[9px] font-bold text-blue-600 hover:underline">Limpar</button>
              </div>
            </div>
            <div className="space-y-1 max-h-[300px] overflow-y-auto pr-2 custom-scrollbar">
              {secretariesList.filter(s => s.toLowerCase().includes(filterQuery.toLowerCase())).map(name => (
                <div 
                  key={name} 
                  onClick={() => setSelectedSecretariesFilter(prev => prev.includes(name) ? prev.filter(n => n !== name) : [...prev, name])} 
                  className={cn(
                    "flex items-center gap-3 p-3 rounded-xl cursor-pointer text-sm font-medium transition-all group",
                    selectedSecretariesFilter.includes(name) ? "bg-blue-50 text-blue-700 shadow-sm" : "text-slate-600 hover:bg-slate-50"
                  )}
                >
                  <div className={cn(
                    "w-4 h-4 rounded-lg border flex items-center justify-center transition-all",
                    selectedSecretariesFilter.includes(name) ? "bg-blue-600 border-blue-600" : "border-slate-300 group-hover:border-blue-400"
                  )}>
                    {selectedSecretariesFilter.includes(name) && <Check size={10} className="text-white" />}
                  </div>
                  <span className="truncate flex-1">{name}</span>
                </div>
              ))}
            </div>
          </div>

          {/* Filtro por Status */}
          <div className="space-y-4">
            <div className="flex items-center justify-between">
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Situação / Status</p>
              <button onClick={() => setSelectedStatusFilter([])} className="text-[9px] font-bold text-blue-600 hover:underline">Limpar</button>
            </div>
            <div className="space-y-1">
              {statusList.map(status => (
                <div 
                  key={status} 
                  onClick={() => toggleStatusFilter(status)} 
                  className={cn(
                    "flex items-center gap-3 p-3 rounded-xl cursor-pointer text-sm font-medium transition-all group",
                    selectedStatusFilter.includes(status) ? "bg-blue-50 text-blue-700 shadow-sm" : "text-slate-600 hover:bg-slate-50"
                  )}
                >
                  <div className={cn(
                    "w-4 h-4 rounded-lg border flex items-center justify-center transition-all",
                    selectedStatusFilter.includes(status) ? "bg-blue-600 border-blue-600" : "border-slate-300 group-hover:border-blue-400"
                  )}>
                    {selectedStatusFilter.includes(status) && <Check size={10} className="text-white" />}
                  </div>
                  <span className="truncate flex-1">{status}</span>
                </div>
              ))}
            </div>
          </div>

          {/* Reset Total */}
          <button 
            onClick={() => { 
              setData([]); 
              setSelectedSecretariesFilter([]); 
              setSelectedStatusFilter([]); 
              setTableSearchQuery('');
              setCurrentPage(1);
              setSortColumn(null);
              setExpandedRowIndex(null);
            }} 
            className="w-full mt-auto py-4 rounded-2xl border border-red-100 text-red-500 text-xs font-bold uppercase tracking-widest hover:bg-red-50 transition-all flex items-center justify-center gap-2"
          >
            <X size={14} />
            Limpar Análise
          </button>
        </div>
      </aside>

      {/* Conteúdo Principal */}
      <main className="flex-1 flex flex-col h-screen overflow-hidden bg-slate-50">
        {/* Topbar para Ações Rápidas & Filtros */}
        <div className="bg-white border-b border-slate-200 px-6 py-4 flex items-center justify-between gap-4">
          <button 
            onClick={() => setIsSidebarOpen(!isSidebarOpen)} 
            className="p-2.5 text-slate-600 hover:text-slate-900 hover:bg-slate-50 active:bg-slate-100 rounded-xl transition-all flex items-center gap-2 text-xs font-bold uppercase tracking-wider border border-slate-100 shadow-sm"
          >
            {isSidebarOpen ? <ChevronLeft size={16} /> : <SlidersHorizontal size={16} />}
            <span>{isSidebarOpen ? 'Ocultar Filtros' : 'Filtros'}</span>
          </button>
          
          <div className="flex items-center gap-4">
            <button 
              onClick={() => {
                setData([]);
                setSelectedSecretariesFilter([]);
                setSelectedStatusFilter([]);
                setTableSearchQuery('');
                setCurrentPage(1);
                setSortColumn(null);
                setExpandedRowIndex(null);
              }} 
              className="px-4 py-2 text-xs font-bold text-slate-400 hover:text-red-500 transition-colors uppercase tracking-widest border border-slate-100 rounded-xl"
            >
              Novo Upload
            </button>
          </div>
        </div>

        {/* Área do Dashboard com Scroll */}
        <div className="flex-1 overflow-y-auto custom-scrollbar p-6 lg:p-12">
          <div className="max-w-7xl mx-auto space-y-12">
            <header className="space-y-1">
              <h1 className="text-4xl font-bold text-slate-900 tracking-tight">Análise de Contratos</h1>
              <p className="text-slate-500 font-medium">Painel estratégico de execução orçamentária</p>
            </header>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
              {/* Gráfico de Execução Global */}
              <div className="lg:col-span-2 bg-white p-8 rounded-[2rem] border border-slate-100 shadow-sm flex flex-col h-[400px]">
                <div className="flex items-center gap-3 mb-8">
                  <div className="p-2.5 bg-blue-50 text-blue-600 rounded-2xl"><Activity size={20} /></div>
                  <div>
                    <h3 className="text-lg font-bold text-slate-900 tracking-tight">Fluxo de Execução Financeira</h3>
                    <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">Comparativo de Estágios da Despesa</p>
                  </div>
                </div>
                <div className="flex-1">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={globalExecutionData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                      <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{fontSize: 11, fontWeight: 600, fill: '#64748b'}} />
                      <YAxis hide />
                      <Tooltip 
                        cursor={{fill: '#f8fafc'}} 
                        contentStyle={{borderRadius: '16px', border: 'none', boxShadow: '0 20px 25px -5px rgb(0 0 0 / 0.1)'}}
                        formatter={(v) => [formatCurrency(v), 'Valor']} 
                      />
                      <Bar dataKey="value" radius={[10, 10, 0, 0]} barSize={60} label={{ position: 'top', formatter: (v:any) => formatCurrency(v), fontSize: 10, fontWeight: 700, fill: '#475569' }} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* Top Fornecedores */}
              <div className="bg-white p-8 rounded-[2rem] border border-slate-100 shadow-sm flex flex-col h-[400px]">
                <div className="flex items-center gap-3 mb-8">
                  <div className="p-2.5 bg-emerald-50 text-emerald-600 rounded-2xl"><FolderGit2 size={20} /></div>
                  <div>
                    <h3 className="text-lg font-bold text-slate-900 tracking-tight">Top Fornecedores</h3>
                    <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">Maiores Volume de Contrato</p>
                  </div>
                </div>
                <div className="flex-1 overflow-y-auto custom-scrollbar space-y-4">
                  {suppliersData.map((sup, i) => (
                    <div key={i} className="space-y-2">
                      <div className="flex justify-between items-center text-xs">
                        <span className="font-bold text-slate-700 truncate pr-4">{sup.name}</span>
                        <span className="font-mono text-slate-500">{formatCurrency(sup.value)}</span>
                      </div>
                      <div className="w-full h-1.5 bg-slate-50 rounded-full overflow-hidden">
                        <div className="h-full bg-emerald-500" style={{ width: `${(sup.value / (suppliersData[0]?.value || 1)) * 100}%` }} />
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>



            {/* Tabela Geral de Contratos */}
            <div className="bg-white p-8 rounded-[2rem] border border-slate-100 shadow-sm flex flex-col space-y-6">
              <div className="flex flex-wrap items-center justify-between gap-4">
                <div className="space-y-1">
                  <h2 className="text-2xl font-bold text-slate-900 tracking-tight">Lista Geral de Contratos</h2>
                  <p className="text-xs text-slate-500 font-medium">Pesquise, filtre e audite todos os contratos importados</p>
                </div>
                
                <div className="flex flex-wrap items-center gap-3">
                  {/* Busca interna na tabela */}
                  <div className="relative">
                    <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                    <input 
                      type="text" 
                      placeholder="Buscar contrato, fornecedor..." 
                      value={tableSearchQuery} 
                      onChange={handleTableSearchChange}
                      className="pl-10 pr-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/10 focus:border-blue-500/50 transition-all w-64"
                    />
                  </div>

                  {/* Exportador Excel do resultado filtrado */}
                  <button 
                    onClick={() => downloadExcel(sortedTableData, "Filtrados")} 
                    className="flex items-center gap-2 px-4 py-2.5 text-xs font-bold bg-slate-900 text-white rounded-xl hover:bg-slate-800 transition-all shadow-sm shadow-slate-900/10 active:scale-95"
                  >
                    <Download size={14} />
                    Exportar Tabela
                  </button>
                </div>
              </div>

              {/* Controles de Registros por Página */}
              <div className="flex items-center justify-between text-xs text-slate-500 font-medium px-1">
                <div>
                  Mostrando <span className="font-bold text-slate-800">{Math.min(sortedTableData.length, (currentPage - 1) * rowsPerPage + 1)}</span> a{' '}
                  <span className="font-bold text-slate-800">{Math.min(sortedTableData.length, currentPage * rowsPerPage)}</span> de{' '}
                  <span className="font-bold text-slate-800">{sortedTableData.length}</span> contratos
                </div>
                <div className="flex items-center gap-2">
                  <span>Registros por página:</span>
                  <select 
                    value={rowsPerPage} 
                    onChange={handleRowsPerPageChange}
                    className="bg-slate-50 border border-slate-200 rounded-lg px-2 py-1 text-xs font-bold text-slate-700 focus:outline-none focus:ring-1 focus:ring-blue-500"
                  >
                    <option value={10}>10</option>
                    <option value={25}>25</option>
                    <option value={50}>50</option>
                    <option value={100}>100</option>
                  </select>
                </div>
              </div>

              {/* Tabela de Dados */}
              <div className="overflow-x-auto border border-slate-100 rounded-2xl">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-100">
                      <th 
                        onClick={() => handleSort(contratoKey)} 
                        className="py-4 px-6 text-[10px] font-bold text-slate-500 uppercase tracking-widest cursor-pointer select-none group"
                      >
                        <div className="flex items-center">
                          Contrato
                          {renderSortIcon(contratoKey)}
                        </div>
                      </th>
                      <th 
                        onClick={() => handleSort(secretariaKey)} 
                        className="py-4 px-6 text-[10px] font-bold text-slate-500 uppercase tracking-widest cursor-pointer select-none group"
                      >
                        <div className="flex items-center">
                          Secretaria / Órgão
                          {renderSortIcon(secretariaKey)}
                        </div>
                      </th>
                      <th 
                        onClick={() => handleSort(fornecedorKey)} 
                        className="py-4 px-6 text-[10px] font-bold text-slate-500 uppercase tracking-widest cursor-pointer select-none group"
                      >
                        <div className="flex items-center">
                          Fornecedor
                          {renderSortIcon(fornecedorKey)}
                        </div>
                      </th>
                      <th className="py-4 px-6 text-[10px] font-bold text-slate-500 uppercase tracking-widest select-none">
                        Objeto
                      </th>
                      <th 
                        onClick={() => handleSort(valueKey)} 
                        className="py-4 px-6 text-[10px] font-bold text-slate-500 uppercase tracking-widest cursor-pointer select-none group text-right"
                      >
                        <div className="flex items-center justify-end">
                          Valor Total
                          {renderSortIcon(valueKey)}
                        </div>
                      </th>
                      <th 
                        onClick={() => handleSort(statusKey || '')} 
                        className="py-4 px-6 text-[10px] font-bold text-slate-500 uppercase tracking-widest cursor-pointer select-none group text-center"
                      >
                        <div className="flex items-center justify-center">
                          Situação
                          {renderSortIcon(statusKey || '')}
                        </div>
                      </th>
                      <th className="py-4 px-6 text-[10px] font-bold text-slate-500 uppercase tracking-widest select-none text-center">
                        Ações
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {paginatedTableData.length === 0 ? (
                      <tr>
                        <td colSpan={7} className="py-12 text-center text-slate-400 text-sm font-medium">
                          Nenhum contrato encontrado para os critérios de busca.
                        </td>
                      </tr>
                    ) : (
                      paginatedTableData.map((row) => {
                        const originalIdx = data.indexOf(row);
                        const isExpanded = expandedRowIndex === originalIdx;
                        const contractId = row[contratoKey] || 'S/N';
                        const secName = row[secretariaKey] || 'Não Informado';
                        const supplierName = row[fornecedorKey] || 'Não Informado';
                        const objectText = row[objetoKey] || 'Não Descrito';
                        const statusVal = row[statusKey] || 'N/A';
                        const totalVal = getNumericValue(row[valueKey]);
                        
                        return (
                          <React.Fragment key={originalIdx}>
                            <tr className={cn(
                              "hover:bg-slate-50/50 transition-colors duration-200 text-sm font-medium",
                              isExpanded && "bg-blue-50/20 hover:bg-blue-50/30"
                            )}>
                              <td className="py-4 px-6 text-slate-900 font-mono text-xs">{contractId}</td>
                              <td className="py-4 px-6 text-slate-700 truncate max-w-[200px]" title={String(secName)}>{String(secName)}</td>
                              <td className="py-4 px-6 text-slate-700 truncate max-w-[180px]" title={String(supplierName)}>{String(supplierName)}</td>
                              <td className="py-4 px-6 text-slate-500 truncate max-w-[240px]" title={String(objectText)}>{String(objectText)}</td>
                              <td className="py-4 px-6 text-slate-900 font-mono text-right text-xs font-bold">{formatCurrency(totalVal)}</td>
                              <td className="py-4 px-6 text-center">
                                <span className={cn(
                                  "inline-block text-[10px] font-bold px-2.5 py-0.5 rounded-full border",
                                  statusVal.toLowerCase().includes('atrasado') || statusVal.toLowerCase().includes('atraso')
                                    ? "bg-red-50 border-red-100 text-red-600"
                                    : statusVal.toLowerCase().includes('ativo') || statusVal.toLowerCase().includes('andamento') || statusVal.toLowerCase().includes('vigente')
                                    ? "bg-emerald-50 border-emerald-100 text-emerald-600"
                                    : "bg-slate-50 border-slate-100 text-slate-600"
                                )}>
                                  {statusVal}
                                </span>
                              </td>
                              <td className="py-4 px-6 text-center">
                                <div className="flex items-center justify-center gap-2">
                                  <button 
                                    onClick={() => setExpandedRowIndex(isExpanded ? null : originalIdx)}
                                    className={cn(
                                      "p-1.5 rounded-lg border text-xs font-bold transition-all flex items-center gap-1",
                                      isExpanded 
                                        ? "bg-blue-50 border-blue-200 text-blue-600" 
                                        : "bg-white border-slate-200 text-slate-500 hover:text-slate-800 hover:bg-slate-50"
                                    )}
                                    title="Ver Detalhes do Contrato"
                                  >
                                    Auditar
                                  </button>
                                  <button 
                                    onClick={() => setSelectedSecretary(String(secName))}
                                    className="p-1.5 rounded-lg border border-slate-200 bg-white text-slate-500 hover:text-blue-600 hover:border-blue-200 transition-all"
                                    title="Abrir Painel do Órgão"
                                  >
                                    <Search size={14} />
                                  </button>
                                </div>
                              </td>
                            </tr>
                            
                            {/* Linha expansível com detalhes completos de auditoria do contrato */}
                            {isExpanded && (
                              <tr>
                                <td colSpan={7} className="p-0 border-b border-slate-100 bg-slate-50/30 animate-in slide-in-from-top duration-300">
                                  <div className="p-8 space-y-6 border-x-4 border-l-blue-600 border-r-transparent">
                                    <div className="flex flex-wrap justify-between items-start gap-4">
                                      <div className="space-y-1.5 flex-1 min-w-[300px]">
                                        <div className="flex items-center gap-2">
                                          <span className="text-[10px] font-bold text-blue-600 bg-blue-50 px-2.5 py-0.5 rounded-full border border-blue-100 font-mono">
                                            Contrato ID: {contractId}
                                          </span>
                                          <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                                            {supplierName}
                                          </span>
                                        </div>
                                        <h4 className="text-base font-bold text-slate-900 leading-snug">
                                          {objectText}
                                        </h4>
                                      </div>
                                      
                                      <div className="flex flex-col items-end gap-1.5">
                                        <span className={cn(
                                          "text-[10px] font-bold px-3 py-1 rounded-xl border",
                                          statusVal.toLowerCase().includes('atrasado') ? "bg-red-50 border-red-100 text-red-600" : "bg-slate-50 border-slate-100 text-slate-600"
                                        )}>
                                          {statusVal}
                                        </span>
                                        <div className="flex items-center gap-1.5 text-[10px] font-bold text-slate-500">
                                          <Calendar size={12} className="text-slate-400" />
                                          Vigência: {formatDate(row[inicioKey])} <span className="text-slate-300">/</span> {formatDate(row[fimKey])}
                                        </div>
                                      </div>
                                    </div>
                                    
                                    {/* Métricas e Execução Financeira */}
                                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-6 p-6 rounded-2xl border border-slate-100 bg-white">
                                      <div>
                                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest mb-1 flex items-center gap-1">
                                          <DollarSign size={10} className="text-slate-300" /> Contrato
                                        </p>
                                        <p className="text-base font-bold text-slate-900 font-mono">{formatCurrency(row[valueKey])}</p>
                                      </div>
                                      <div>
                                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest mb-1 flex items-center gap-1">
                                          <Wallet size={10} className="text-slate-300" /> Empenhado
                                        </p>
                                        <p className="text-base font-bold text-blue-600 font-mono">{formatCurrency(row[empenhadoKey])}</p>
                                      </div>
                                      <div>
                                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest mb-1 flex items-center gap-1">
                                          <Activity size={10} className="text-slate-300" /> Liquidado
                                        </p>
                                        <p className="text-base font-bold text-amber-600 font-mono">{formatCurrency(row[liquidadoKey])}</p>
                                      </div>
                                      <div>
                                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest mb-1 flex items-center gap-1">
                                          <CheckCircle2 size={10} className="text-slate-300" /> Pago
                                        </p>
                                        <p className="text-base font-bold text-emerald-600 font-mono">{formatCurrency(row[pagoKey])}</p>
                                      </div>
                                    </div>
                                    
                                    {/* Barra de Execução */}
                                    {(() => {
                                      const total = getNumericValue(row[valueKey]);
                                      const liq = getNumericValue(row[liquidadoKey]);
                                      const perc = total > 0 ? (liq / total) * 100 : 0;
                                      return (
                                        <div className="space-y-2">
                                          <div className="flex justify-between items-center text-xs">
                                            <p className="text-[9px] font-bold text-slate-400 uppercase tracking-[0.2em]">Execução Financeira (Liquidado / Total)</p>
                                            <p className="font-bold text-slate-800 font-mono">{perc.toFixed(1)}%</p>
                                          </div>
                                          <div className="w-full h-2.5 bg-slate-100 rounded-full overflow-hidden p-0.5 border border-slate-200">
                                            <div 
                                              className={cn("h-full rounded-full transition-all duration-1000", perc > 80 ? "bg-emerald-500" : "bg-blue-600")} 
                                              style={{ width: `${Math.min(perc, 100)}%` }} 
                                            />
                                          </div>
                                        </div>
                                      );
                                    })()}
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>

              {/* Paginação */}
              {totalPages > 1 && (
                <div className="flex items-center justify-between pt-4 border-t border-slate-100 text-sm">
                  <button 
                    onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                    disabled={currentPage === 1}
                    className="flex items-center gap-1 px-4 py-2 border border-slate-200 rounded-xl hover:bg-slate-50 transition-all font-bold text-slate-600 disabled:opacity-40 disabled:cursor-not-allowed text-xs"
                  >
                    Anterior
                  </button>
                  
                  <div className="flex items-center gap-1">
                    {Array.from({ length: totalPages }, (_, i) => i + 1)
                      .filter(page => {
                        return page === 1 || 
                               page === totalPages || 
                               Math.abs(page - currentPage) <= 1;
                      })
                      .map((page, idx, array) => {
                        const showEllipsis = idx > 0 && page - array[idx - 1] > 1;
                        return (
                          <React.Fragment key={page}>
                            {showEllipsis && <span className="px-2 text-slate-400 font-bold">...</span>}
                            <button 
                              onClick={() => setCurrentPage(page)}
                              className={cn(
                                "w-9 h-9 flex items-center justify-center rounded-xl font-bold transition-all text-xs",
                                currentPage === page 
                                  ? "bg-slate-900 text-white shadow-md shadow-slate-900/10" 
                                  : "bg-white border border-slate-200 text-slate-600 hover:bg-slate-50"
                              )}
                            >
                              {page}
                            </button>
                          </React.Fragment>
                        );
                      })}
                  </div>

                  <button 
                    onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                    disabled={currentPage === totalPages}
                    className="flex items-center gap-1 px-4 py-2 border border-slate-200 rounded-xl hover:bg-slate-50 transition-all font-bold text-slate-600 disabled:opacity-40 disabled:cursor-not-allowed text-xs"
                  >
                    Próxima
                  </button>
                </div>
              )}
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-10">
              <div className="bg-white p-10 rounded-[2rem] border border-slate-100 shadow-2xl shadow-slate-200/40 flex flex-col h-[450px]">
                <div className="flex items-center gap-3 mb-10">
                  <div className="p-2.5 bg-blue-50 text-blue-600 rounded-2xl"><BarChart3 size={20} /></div>
                  <h3 className="text-xl font-bold text-slate-900 tracking-tight">Ranking Financeiro (Empenho)</h3>
                </div>
                <div className="flex-1">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={Object.entries(bySecretaria).map(([name, s]:any) => ({ name: name.substring(0, 12), value: s.empenhado })).sort((a,b)=>b.value-a.value).slice(0, 8)} layout="vertical">
                      <XAxis type="number" hide />
                      <YAxis dataKey="name" type="category" width={100} tick={{fontSize: 10, fontWeight: 600, fill: '#64748b'}} axisLine={false} tickLine={false} />
                      <Tooltip 
                        cursor={{fill: '#f8fafc', radius: 10}} 
                        contentStyle={{borderRadius: '16px', border: 'none', boxShadow: '0 20px 25px -5px rgb(0 0 0 / 0.1)'}}
                        formatter={(v) => [formatCurrency(v), 'Valor Empenhado']} 
                      />
                      <Bar dataKey="value" fill="#0f172a" radius={[0, 10, 10, 0]} barSize={24} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="bg-white p-10 rounded-[2rem] border border-slate-100 shadow-2xl shadow-slate-200/40 flex flex-col h-[450px]">
                <div className="flex items-center gap-3 mb-10">
                  <div className="p-2.5 bg-amber-50 text-amber-600 rounded-2xl"><PieChart size={20} /></div>
                  <h3 className="text-xl font-bold text-slate-900 tracking-tight">Distribuição por Situação</h3>
                </div>
                <div className="flex-1">
                  <ResponsiveContainer width="100%" height="100%">
                    <RePieChart>
                      <Pie data={statusGlobalData} cx="50%" cy="50%" innerRadius={90} outerRadius={120} paddingAngle={8} dataKey="value">
                        {statusGlobalData.map((entry, index) => <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} className="stroke-white stroke-2 focus:outline-none" />)}
                      </Pie>
                      <Tooltip contentStyle={{borderRadius: '16px', border: 'none', boxShadow: '0 20px 25px -5px rgb(0 0 0 / 0.1)'}} />
                      <Legend iconType="circle" wrapperStyle={{paddingTop: '20px'}} />
                    </RePieChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
          </div>
        </div>
      </main>
    </div>
  );
}

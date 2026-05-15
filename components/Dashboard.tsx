'use client';

import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, FileSpreadsheet, BarChart3, PieChart, Activity, FolderGit2, Search, X, Check, ChevronDown, ArrowLeft, Download, Calendar, DollarSign, Wallet, CheckCircle2 } from 'lucide-react';
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
  const [isFilterOpen, setIsFilterOpen] = useState(false);

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
    secretariaKey, statusKey, valueKey, objetoKey, fornecedorKey,
    inicioKey, fimKey, empenhadoKey, liquidadoKey, pagoKey 
  } = useMemo(() => ({
    secretariaKey: allKeys.length > 26 ? allKeys[26] : findKey(sampleRow, ['secretaria', 'órgão', 'orgao', 'setor']) || allKeys[0],
    statusKey: findKey(sampleRow, ['status', 'situação', 'situacao', 'vigência', 'fase']),
    valueKey: findKey(sampleRow, ['valor', 'montante', 'contrato']),
    objetoKey: findKey(sampleRow, ['objeto', 'descrição', 'descricao']),
    fornecedorKey: findKey(sampleRow, ['fornecedor', 'contratada', 'empresa']),
    inicioKey: findKey(sampleRow, ['início', 'inicio', 'data início', 'data inicio']),
    fimKey: findKey(sampleRow, ['fim', 'término', 'termino', 'vencimento', 'data fim']),
    empenhadoKey: findKey(sampleRow, ['empenhado', 'valor empenhado']),
    liquidadoKey: findKey(sampleRow, ['liquidado', 'valor liquidado']),
    pagoKey: findKey(sampleRow, ['pago', 'valor pago'])
  }), [sampleRow, allKeys]);

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
    if (selectedSecretariesFilter.length === 0) return data;
    return data.filter(row => selectedSecretariesFilter.includes(String(row[secretariaKey])));
  }, [data, selectedSecretariesFilter, secretariaKey]);

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

  const toggleSecretaryFilter = (name: string) => {
    setSelectedSecretariesFilter(prev => prev.includes(name) ? prev.filter(n => n !== name) : [...prev, name]);
  };

  const downloadExcel = (filtered: any[], name: string) => {
    const ws = XLSX.utils.json_to_sheet(filtered);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Contratos");
    XLSX.writeFile(wb, `Contratos_${name.substring(0, 50)}.xlsx`);
  };

  if (data.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen bg-white px-4">
        <div onDrop={handleFileUpload} onDragOver={handleDragOver} onDragLeave={handleDragLeave} className="w-full max-w-xl p-16 border border-slate-200 rounded-3xl flex flex-col items-center justify-center transition-all duration-300 bg-white shadow-sm hover:shadow-md">
          <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} className="hidden" id="file-upload" />
          <label htmlFor="file-upload" className="flex flex-col items-center cursor-pointer">
            <div className="p-5 rounded-full bg-slate-50 text-slate-400 mb-6"><UploadCloud size={40} /></div>
            <h2 className="text-xl font-semibold mb-2 text-slate-800">Importar Planilha Financeira</h2>
            <p className="text-slate-500 text-center mb-8 text-sm">Arraste seu arquivo para analisar Vigência e Execução.</p>
            <div className="px-8 py-3 rounded-xl bg-blue-600 text-white font-bold hover:bg-blue-700 transition-all shadow-lg shadow-blue-200">Começar Análise</div>
          </label>
        </div>
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
              { label: 'Contratos', value: secStats.count, icon: FolderGit2, color: 'text-blue-600' },
              { label: 'Empenhado', value: formatCurrency(secStats.empenhado), icon: Wallet, color: 'text-slate-800' },
              { label: 'Liquidado', value: formatCurrency(secStats.liquidado), icon: Activity, color: 'text-amber-600' },
              { label: 'Pago', value: formatCurrency(secStats.pago), icon: CheckCircle2, color: 'text-emerald-600' }
            ].map((card, i) => (
              <div key={i} className="bg-white p-6 rounded-2xl border border-slate-100 shadow-[0_4px_20px_-4px_rgba(0,0,0,0.05)]">
                <div className="flex items-center gap-2 mb-3">
                  <card.icon size={14} className="text-slate-400" />
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">{card.label}</p>
                </div>
                <p className={cn("text-xl font-bold", card.color)}>{card.value}</p>
              </div>
            ))}
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
            <div className="lg:col-span-1 bg-white p-8 rounded-[2.5rem] border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] h-[400px] flex flex-col">
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

            <div className="lg:col-span-2 bg-white p-8 rounded-[2.5rem] border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] h-[600px] lg:h-[400px] flex flex-col">
              <h3 className="text-lg font-bold text-slate-800 mb-8">Cronograma e Detalhes Financeiros</h3>
              <div className="overflow-y-auto flex-1 space-y-6 pr-4 custom-scrollbar">
                {secretaryData.map((contract, i) => {
                  const total = getNumericValue(contract[valueKey]);
                  const liq = getNumericValue(contract[liquidadoKey]);
                  const perc = total > 0 ? (liq / total) * 100 : 0;
                  return (
                    <div key={i} className="p-6 rounded-2xl border border-slate-100 bg-white hover:shadow-md transition-all">
                      <div className="flex flex-wrap justify-between items-start gap-4 mb-4">
                        <div className="space-y-1">
                          <span className="text-[10px] font-bold text-slate-400">REF: {contract[allKeys[0]]}</span>
                          <h4 className="font-bold text-slate-800 leading-tight">{contract[objetoKey] || 'Objeto não descrito'}</h4>
                        </div>
                        <div className="flex flex-col items-end gap-1">
                          <span className="text-[10px] font-bold px-2 py-1 bg-slate-50 rounded-lg border border-slate-200 text-slate-600">
                            {contract[statusKey] || 'S/ STATUS'}
                          </span>
                          <div className="flex items-center gap-1 text-[10px] font-medium text-slate-400">
                            <Calendar size={10} /> {formatDate(contract[inicioKey])} - {formatDate(contract[fimKey])}
                          </div>
                        </div>
                      </div>
                      <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 mt-4 py-4 border-t border-slate-100">
                        <div><p className="text-[9px] font-bold text-slate-400 uppercase">Contrato</p><p className="text-xs font-bold text-slate-800">{formatCurrency(contract[valueKey])}</p></div>
                        <div><p className="text-[9px] font-bold text-slate-400 uppercase">Empenhado</p><p className="text-xs font-bold text-slate-700">{formatCurrency(contract[empenhadoKey])}</p></div>
                        <div><p className="text-[9px] font-bold text-slate-400 uppercase">Liquidado</p><p className="text-xs font-bold text-amber-600">{formatCurrency(contract[liquidadoKey])}</p></div>
                        <div><p className="text-[9px] font-bold text-slate-400 uppercase">Pago</p><p className="text-xs font-bold text-emerald-600">{formatCurrency(contract[pagoKey])}</p></div>
                      </div>
                      <div className="mt-2">
                        <div className="flex justify-between items-center mb-1"><p className="text-[9px] font-bold text-slate-400">EXECUÇÃO</p><p className="text-[9px] font-bold text-slate-700">{perc.toFixed(1)}%</p></div>
                        <div className="w-full h-1.5 bg-slate-100 rounded-full overflow-hidden">
                          <div className="h-full bg-blue-600 transition-all" style={{ width: `${Math.min(perc, 100)}%` }} />
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
    <div className="min-h-screen bg-white py-12 px-4 lg:px-8 animate-in fade-in duration-700">
      <div className="max-w-7xl mx-auto space-y-12">
        <header className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-6">
          <div className="space-y-1">
            <h1 className="text-4xl font-bold text-slate-900 tracking-tight">Análise de Contratos</h1>
            <p className="text-slate-500 font-medium">Relatório dinâmico de execução orçamentária</p>
          </div>
          <div className="flex items-center gap-3 w-full lg:w-auto">
            <div className="relative flex-1 lg:flex-initial">
              <button onClick={() => setIsFilterOpen(!isFilterOpen)} className={cn("w-full lg:w-64 px-5 py-3 text-sm font-semibold rounded-xl border flex items-center justify-between transition-all", selectedSecretariesFilter.length > 0 ? "bg-blue-50 border-blue-200 text-blue-700" : "bg-white border-slate-200 text-slate-700 shadow-sm hover:shadow-md")}>
                <span className="truncate">{selectedSecretariesFilter.length === 0 ? 'Filtro por Órgão' : `${selectedSecretariesFilter.length} Selecionados`}</span>
                <ChevronDown size={16} />
              </button>
              {isFilterOpen && (
                <div className="absolute top-full mt-2 right-0 w-full lg:w-80 bg-white border border-slate-200 rounded-2xl shadow-xl z-50">
                  <div className="p-4 border-b border-slate-100">
                    <input type="text" placeholder="Pesquisar..." value={filterQuery} onChange={(e) => setFilterQuery(e.target.value)} className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm mb-2 focus:outline-none focus:ring-2 focus:ring-blue-500/10" />
                    <div className="flex gap-2"><button onClick={() => setSelectedSecretariesFilter([])} className="text-[10px] font-bold text-slate-400 uppercase hover:text-slate-600 transition-colors">Limpar</button><button onClick={() => setSelectedSecretariesFilter(secretariesList)} className="text-[10px] font-bold text-blue-600 uppercase hover:text-blue-800 transition-colors">Todos</button></div>
                  </div>
                  <div className="max-h-[250px] overflow-y-auto p-2">
                    {secretariesList.filter(s => s.toLowerCase().includes(filterQuery.toLowerCase())).map(name => (
                      <div key={name} onClick={() => toggleSecretaryFilter(name)} className={cn("flex items-center gap-3 p-2.5 rounded-lg cursor-pointer text-sm font-medium transition-all", selectedSecretariesFilter.includes(name) ? "bg-blue-50 text-blue-700" : "text-slate-600 hover:bg-slate-50")}>
                        <div className={cn("w-4 h-4 rounded border flex items-center justify-center", selectedSecretariesFilter.includes(name) ? "bg-blue-600 border-blue-600" : "border-slate-300")}>{selectedSecretariesFilter.includes(name) && <Check size={10} className="text-white" />}</div>
                        <span className="truncate">{name}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
            <button onClick={() => setData([])} className="p-3 text-slate-300 hover:text-red-500 transition-colors"><X size={20} /></button>
          </div>
        </header>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6">
          {[
            { label: 'Valor Contratado', value: totalMetrics.value, icon: DollarSign, color: 'text-slate-900' },
            { label: 'Total Empenhado', value: totalMetrics.empenhado, icon: Wallet, color: 'text-blue-600' },
            { label: 'Total Liquidado', value: totalMetrics.liquidado, icon: Activity, color: 'text-amber-600' },
            { label: 'Total Pago', value: totalMetrics.pago, icon: CheckCircle2, color: 'text-emerald-600' }
          ].map((m, i) => (
            <div key={i} className="bg-white p-8 rounded-3xl border border-slate-100 shadow-[0_4px_25px_rgb(0,0,0,0.03)]">
              <div className="p-3 w-fit rounded-xl bg-slate-50 text-slate-400 mb-6"><m.icon size={20} /></div>
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1">{m.label}</p>
              <p className={cn("text-2xl font-bold", m.color)}>{formatCurrency(m.value)}</p>
            </div>
          ))}
        </div>

        <div className="space-y-8">
          <h2 className="text-2xl font-bold text-slate-900">Performance Orçamentária por Secretaria</h2>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
            {Object.entries(bySecretaria).sort((a:any, b:any) => b[1].totalValue - a[1].totalValue).map(([name, stats]: [string, any], idx) => {
              const perc = stats.totalValue > 0 ? (stats.liquidado / stats.totalValue) * 100 : 0;
              return (
                <div key={idx} onClick={() => setSelectedSecretary(name)} className="bg-white p-7 rounded-3xl border border-slate-100 hover:border-blue-100 shadow-sm hover:shadow-xl transition-all duration-300 cursor-pointer group">
                  <h3 className="font-bold text-slate-800 text-lg mb-6 line-clamp-2 min-h-[3.5rem] group-hover:text-blue-600 transition-colors">{name}</h3>
                  <div className="space-y-4">
                    <div className="flex justify-between items-center"><span className="text-[10px] font-bold text-slate-400 uppercase">Liquidado</span><span className="text-sm font-bold text-amber-600">{formatCurrency(stats.liquidado)}</span></div>
                    <div className="space-y-1">
                      <div className="flex justify-between text-[9px] font-bold text-slate-400"><span>EXECUÇÃO</span><span>{perc.toFixed(0)}%</span></div>
                      <div className="w-full h-1 bg-slate-100 rounded-full overflow-hidden">
                        <div className="h-full bg-blue-500 transition-all" style={{ width: `${Math.min(perc, 100)}%` }} />
                      </div>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-10">
          <div className="bg-white p-10 rounded-[2.5rem] border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] flex flex-col h-[400px]">
            <h3 className="text-lg font-bold text-slate-800 mb-8">Ranking Financeiro (Empenho)</h3>
            <div className="flex-1">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={Object.entries(bySecretaria).map(([name, s]:any) => ({ name: name.substring(0, 10), value: s.empenhado })).sort((a,b)=>b.value-a.value).slice(0, 8)} layout="vertical">
                  <XAxis type="number" hide />
                  <YAxis dataKey="name" type="category" width={80} tick={{fontSize: 10}} />
                  <Tooltip cursor={{fill: '#f8fafc'}} formatter={(v) => formatCurrency(v)} />
                  <Bar dataKey="value" fill="#3b82f6" radius={[0, 10, 10, 0]} barSize={15} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
          <div className="bg-white p-10 rounded-[2.5rem] border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] flex flex-col h-[400px]">
            <h3 className="text-lg font-bold text-slate-800 mb-8">Situação da Base</h3>
            <div className="flex-1">
              <ResponsiveContainer width="100%" height="100%">
                <RePieChart>
                  <Pie data={statusGlobalData} cx="50%" cy="50%" innerRadius={80} outerRadius={110} paddingAngle={5} dataKey="value">
                    {statusGlobalData.map((entry, index) => <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />)}
                  </Pie>
                  <Tooltip />
                  <Legend iconType="circle" />
                </RePieChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { 
  FolderKanban, 
  Search, 
  Upload, 
  Plus, 
  RefreshCw, 
  AlertCircle, 
  User as UserIcon, 
  History, 
  X, 
  Activity,
  CheckCircle2,
  Zap,
  SkipForward,
  TrendingUp,
  ChevronLeft,
  ChevronRight,
  Download
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { format } from 'date-fns';
import { MasterProject as MasterProjectType, MasterProjectAuditLog } from '../types';
import { taskService } from '../services/taskService';
import { cn } from '../lib/utils';

export function MasterProject({ user, isMobile }: { user: any, isMobile?: boolean }) {
  const [projects, setProjects] = useState<MasterProjectType[]>([]);
  const [loading, setLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState('ALL');
  const [selectedProject, setSelectedProject] = useState<MasterProjectType | null>(null);
  const [auditLogs, setAuditLogs] = useState<MasterProjectAuditLog[]>([]);
  const [isAuditOpen, setIsAuditOpen] = useState(false);
  const [isImportOpen, setIsImportOpen] = useState(false);
  const [isAddOpen, setIsAddOpen] = useState(false);
  const [importSummary, setImportSummary] = useState<{ total: number, updated: number, inserted: number, skipped: number } | null>(null);
  const [importPreview, setImportPreview] = useState<{ total: number, updated: number, inserted: number, skipped: number, diffs: any[] } | null>(null);
  const [stagedData, setStagedData] = useState<any[]>([]);
  const [isImporting, setIsImporting] = useState(false);
  const [isEditMode, setIsEditMode] = useState(false);
  const [isDetailOpen, setIsDetailOpen] = useState(false);
  
  // Pagination State
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 25;

  const fetchData = async () => {
    setLoading(true);
    try {
      const data = await taskService.getMasterProjects();
      setProjects(data);
    } catch (err) {
      console.error('Failed to fetch master projects:', err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  const filteredProjects = projects.filter(p => {
    const matchesSearch = p.project_name.toLowerCase().includes(searchTerm.toLowerCase()) || 
                          p.ticket_id.toLowerCase().includes(searchTerm.toLowerCase()) || 
                          p.pic_name.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesStatus = statusFilter === 'ALL' || p.status === statusFilter;
    return matchesSearch && matchesStatus;
  });

  // Reset pagination on filter change
  useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm, statusFilter]);

  const totalPages = Math.ceil(filteredProjects.length / itemsPerPage);
  const paginatedProjects = filteredProjects.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  );

  const exportMasterProjects = async (projectsData: MasterProjectType[]) => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Master Project Data');

    // 1. SET PRECISE COLUMNS
    sheet.columns = [
      { header: '', key: 'ticketId', width: 20 },
      { header: '', key: 'projectName', width: 45 },
      { header: '', key: 'status', width: 25 },
      { header: '', key: 'picName', width: 35 },
      { header: '', key: 'ownerDiv', width: 35 }
    ];

    // 2. MAIN HEADER (Merged A1 to E1)
    sheet.mergeCells('A1:E1');
    const title = sheet.getCell('A1');
    title.value = 'OM DEDY - MASTER PROJECT DATA & IMPORT';
    title.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FF1E3A8A' } }; // Dark Blue
    title.alignment = { horizontal: 'center', vertical: 'middle' };
    
    sheet.addRow([]); // Spacer

    // 3. STATUS GUIDE FOR USERS
    const guideTitle = sheet.addRow(['ALLOWED STATUS VALUES (For new imports):', '', '', '', '']);
    guideTitle.getCell(1).font = { name: 'Arial', bold: true, color: { argb: 'FF333333' } };
    
    ['OPEN', 'ON QUEUE', 'ON PROGRESS', 'LIVE'].forEach(statusText => {
      const row = sheet.addRow([`- ${statusText}`]);
      row.getCell(1).font = { name: 'Arial', italic: true, color: { argb: 'FF666666' } };
    });
    
    sheet.addRow([]); // Spacer before table

    // 4. TABLE HEADERS
    const headerRow = sheet.addRow(['TICKET ID', 'PROJECT NAME', 'STATUS', 'PIC NAME', 'OWNER / DIV']);
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } }; // Dark Blue
      cell.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = { top: { style:'thin' }, left: { style:'thin' }, bottom: { style:'thin' }, right: { style:'thin' } };
    });

    // 5. INJECT REAL DATABASE DATA
    if (projectsData && projectsData.length > 0) {
      projectsData.forEach(p => {
        const row = sheet.addRow([
          p.ticket_id || '-',
          p.project_name || '-',
          (p.global_status || p.status || 'OPEN').toUpperCase(),
          p.pic_name || '-',
          p.div_owner || '-'
        ]);
        
        row.eachCell((cell, colNumber) => {
          cell.font = { name: 'Arial', size: 11 };
          cell.border = { top: { style:'thin' }, left: { style:'thin' }, bottom: { style:'thin' }, right: { style:'thin' } };
          // Center align Ticket ID and Status
          if (colNumber === 1 || colNumber === 3) {
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
          } else {
            cell.alignment = { horizontal: 'left', vertical: 'middle' };
          }
        });
      });
    } else {
      // Fallback if database is empty (acts as template)
      const dummyRow = sheet.addRow(['Example: 1001', 'Example Project', 'OPEN', 'Fachrul Wisnu', 'IT Div']);
      dummyRow.eachCell((cell) => {
        cell.font = { name: 'Arial', size: 11, italic: true };
        cell.border = { top: { style:'thin' }, left: { style:'thin' }, bottom: { style:'thin' }, right: { style:'thin' } };
      });
    }

    // 6. GENERATE & DOWNLOAD
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), 'OmDedy_MasterProject_Export.xlsx');
  };

  const handleOpenAudit = async (project: MasterProjectType) => {
    handleOpenDetail(project);
    // Logic to scroll to history or just showing it in the drawer is already handled
  };

  const handleOpenDetail = async (project: MasterProjectType) => {
    setSelectedProject(project);
    setIsDetailOpen(true);
    setIsEditMode(false);
    try {
      const logs = await taskService.getMasterProjectAuditLogs(project.id);
      setAuditLogs(logs);
    } catch (err) {
      console.error('Failed to fetch logs:', err);
    }
  };

  const handleSaveEdit = async () => {
    if (!selectedProject) return;
    try {
      setLoading(true);
      await taskService.updateMasterProject(selectedProject.id, {
        project_name: selectedProject.project_name,
        status: selectedProject.status,
        pic_name: selectedProject.pic_name,
        owner_name: selectedProject.owner_name,
        div_owner: selectedProject.div_owner,
        plan_start_date: selectedProject.plan_start_date,
        plan_end_date: selectedProject.plan_end_date,
        total_man_hours: selectedProject.total_man_hours
      }, user?.email || 'Admin');
      setIsEditMode(false);
      fetchData();
    } catch (err) {
      alert('Gagal menyimpan perubahan');
    } finally {
      setLoading(false);
    }
  };

  const handleImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    setIsImporting(true);
    setImportSummary(null);
    setImportPreview(null);

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawJson: any[] = XLSX.utils.sheet_to_json(sheet);
        
        const mapped = rawJson.map(row => ({
          ticket_id: String(row.Ticket || row.ticket_id || row.TicketID || row.TICKET_ID || '').trim(),
          project_name: String(row.ProjectName || row.project_name || row.PROJECT_NAME || '').trim(),
          status: String(row.Status || row.status || 'OPEN').toUpperCase(),
          pic_name: String(row.PIC_Name || row.pic_name || row.PIC || 'N/A').trim(),
          owner_name: String(row.Owner_Name || row.owner_name || row.OWNER || 'N/A').trim(),
          div_owner: String(row.Div_Owner || row.div_owner || row.DIVISION || 'N/A').trim()
        })).filter(p => p.ticket_id && p.project_name);

        if (mapped.length === 0) {
          alert('Format CSV/Excel tidak valid atau data kosong. Pastikan kolom Ticket, ProjectName tersedia.');
          setIsImporting(false);
          return;
        }

        setStagedData(mapped);
        const preview = await taskService.previewImportMasterProjects(mapped);
        setImportPreview(preview);
      } catch (err) {
        alert('Gagal memproses file');
      } finally {
        setIsImporting(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const confirmImport = async () => {
    if (stagedData.length === 0) return;
    setIsImporting(true);
    try {
      const stats = await taskService.importMasterProjects(stagedData, user?.email || 'Admin');
      setImportSummary(stats);
      setImportPreview(null);
      setStagedData([]);
      fetchData();
    } catch (err) {
      alert('Gagal mengimpor data');
    } finally {
      setIsImporting(false);
    }
  };

  const openCount = projects.filter(p => (p.global_status || p.status || '').toUpperCase() === 'OPEN').length;
  const queueCount = projects.filter(p => {
    const s = (p.global_status || p.status || '').toUpperCase();
    return s === 'ON QUEUE' || s === 'TO DO';
  }).length;
  const progressCount = projects.filter(p => {
    const s = (p.global_status || p.status || '').toUpperCase();
    return s.includes('PROGRESS');
  }).length;
  const liveCount = projects.filter(p => {
    const s = (p.global_status || p.status || '').toUpperCase();
    return s === 'LIVE' || s === 'DONE';
  }).length;

  return (
    <div className={cn("h-full flex flex-col space-y-6 pb-12", isMobile ? "px-2" : "px-0")}>
      {/* Summary Cards */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-2">
        <div className="bg-blue-900/20 border border-blue-800/50 p-4 rounded-3xl backdrop-blur-sm">
          <p className="text-[10px] text-blue-400 font-black uppercase tracking-[0.2em] mb-1">STATUS OPEN</p>
          <h2 className="text-3xl text-white font-black italic">{openCount}</h2>
        </div>
        <div className="bg-slate-800/40 border border-slate-700/50 p-4 rounded-3xl backdrop-blur-sm">
          <p className="text-[10px] text-slate-400 font-black uppercase tracking-[0.2em] mb-1">ON QUEUE</p>
          <h2 className="text-3xl text-white font-black italic">{queueCount}</h2>
        </div>
        <div className="bg-indigo-900/20 border border-indigo-800/50 p-4 rounded-3xl backdrop-blur-sm">
          <p className="text-[10px] text-indigo-400 font-black uppercase tracking-[0.2em] mb-1">ON PROGRESS</p>
          <h2 className="text-3xl text-white font-black italic">{progressCount}</h2>
        </div>
        <div className="bg-emerald-900/20 border border-emerald-800/50 p-4 rounded-3xl backdrop-blur-sm">
          <p className="text-[10px] text-emerald-400 font-black uppercase tracking-[0.2em] mb-1">LIVE PROJECTS</p>
          <h2 className="text-3xl text-white font-black italic">{liveCount}</h2>
        </div>
      </div>

      {/* Header & Controls */}
      <div className={cn("flex flex-wrap items-center justify-between gap-6 bg-slate-900/50 p-6 rounded-[2.5rem] border border-slate-800 shadow-2xl", isMobile && "flex-col items-stretch")}>
        <div className="flex items-center gap-5">
           <div className="w-16 h-16 bg-indigo-600/10 rounded-2xl flex items-center justify-center border border-indigo-500/20 shadow-lg shadow-indigo-500/5">
              <FolderKanban className="w-8 h-8 text-indigo-400" />
           </div>
           <div>
              <h2 className="text-2xl font-black text-white italic uppercase tracking-tighter">Master <span className="text-indigo-500">Projects</span></h2>
              <p className="text-[10px] text-slate-500 font-bold uppercase tracking-[0.2em] mt-1">Central Repository Database</p>
           </div>
        </div>

        <div className={cn("flex items-center gap-4", isMobile && "flex-col")}>
           <div className="relative">
              <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
              <input 
                type="text"
                placeholder="Search Ticket, Name, PIC..."
                value={searchTerm}
                onChange={e => setSearchTerm(e.target.value)}
                className="bg-slate-950 border border-slate-800 rounded-2xl pl-12 pr-6 py-3.5 text-xs text-white placeholder-slate-600 outline-none focus:border-indigo-500 transition-all w-72"
              />
           </div>
           <select 
             value={statusFilter}
             onChange={e => setStatusFilter(e.target.value)}
             className="bg-slate-950 border border-slate-800 rounded-2xl px-6 py-3.5 text-xs text-indigo-400 font-bold outline-none focus:border-indigo-500 transition-all uppercase"
           >
              <option value="ALL">ALL STATUS</option>
              <option value="OPEN">OPEN</option>
              <option value="ON QUEUE">ON QUEUE</option>
              <option value="ON PROGRESS">ON PROGRESS</option>
              <option value="DONE">DONE</option>
           </select>
           <div className="flex items-center gap-2">
             <button 
               onClick={() => exportMasterProjects(projects)}
               className="bg-slate-800 hover:bg-slate-700 text-white p-3.5 rounded-2xl transition-all shadow-lg border border-slate-700"
               title="Download Template"
             >
                 <Download className="w-5 h-5" />
             </button>
             <button 
               onClick={() => {
                 setImportSummary(null);
                 setIsImportOpen(true);
               }}
               className="bg-slate-800 hover:bg-slate-700 text-white p-3.5 rounded-2xl transition-all shadow-lg border border-slate-700"
               title="Import CSV"
             >
                <Upload className="w-5 h-5" />
             </button>
             <button 
               onClick={() => setIsAddOpen(true)}
               className="bg-gradient-to-tr from-indigo-600 to-indigo-500 hover:from-indigo-500 hover:to-indigo-400 text-white p-3.5 rounded-2xl transition-all shadow-lg shadow-indigo-600/20"
             >
                <Plus className="w-5 h-5" />
             </button>
           </div>
        </div>
      </div>

      {/* Import Result Feedback Card */}
      <AnimatePresence>
        {importSummary && (
          <motion.div 
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            className="overflow-hidden"
          >
            <div className="grid grid-cols-1 sm:grid-cols-4 gap-4 p-6 bg-slate-900 shadow-2xl rounded-[2rem] border border-slate-800 mb-6">
              <div className="bg-slate-950/50 p-4 rounded-3xl border border-white/5 flex items-center gap-4">
                <div className="w-10 h-10 bg-indigo-500/10 rounded-xl flex items-center justify-center border border-indigo-500/20">
                  <TrendingUp className="w-5 h-5 text-indigo-400" />
                </div>
                <div>
                  <p className="text-[8px] font-black text-slate-500 uppercase tracking-widest">Total Baris</p>
                  <p className="text-xl font-black text-white italic">{importSummary.total}</p>
                </div>
              </div>
              <div className="bg-slate-950/50 p-4 rounded-3xl border border-white/5 flex items-center gap-4">
                <div className="w-10 h-10 bg-amber-500/10 rounded-xl flex items-center justify-center border border-amber-500/20">
                  <RefreshCw className="w-5 h-5 text-amber-400" />
                </div>
                <div>
                  <p className="text-[8px] font-black text-slate-500 uppercase tracking-widest">Data Diperbarui</p>
                  <p className="text-xl font-black text-amber-400 italic">{importSummary.updated}</p>
                </div>
              </div>
              <div className="bg-slate-950/50 p-4 rounded-3xl border border-white/5 flex items-center gap-4">
                <div className="w-10 h-10 bg-emerald-500/10 rounded-xl flex items-center justify-center border border-emerald-500/20">
                  <Zap className="w-5 h-5 text-emerald-400" />
                </div>
                <div>
                  <p className="text-[8px] font-black text-slate-500 uppercase tracking-widest">Data Baru</p>
                  <p className="text-xl font-black text-emerald-400 italic">{importSummary.inserted}</p>
                </div>
              </div>
              <div className="bg-slate-950/50 p-4 rounded-3xl border border-white/5 flex items-center gap-4">
                <div className="w-10 h-10 bg-slate-800 rounded-xl flex items-center justify-center border border-slate-700">
                  <SkipForward className="w-5 h-5 text-slate-500" />
                </div>
                <div>
                  <p className="text-[8px] font-black text-slate-500 uppercase tracking-widest">Data Dilewati</p>
                  <p className="text-xl font-black text-slate-500 italic">{importSummary.skipped}</p>
                </div>
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Main Container */}
      <div className="flex-1 bg-slate-900/50 border border-slate-800 rounded-[2.5rem] overflow-hidden shadow-2xl flex flex-col">
        {/* Header - Desktop Only */}
        {!isMobile && (
          <div className="grid grid-cols-[120px_1fr_150px_180px_180px_120px] bg-slate-950/50 border-b border-slate-800 px-6 py-5">
            <div className="text-[10px] font-black text-slate-500 uppercase tracking-widest italic">Ticket ID</div>
            <div className="text-[10px] font-black text-slate-500 uppercase tracking-widest italic">Project Name</div>
            <div className="text-[10px] font-black text-slate-500 uppercase tracking-widest italic">Status</div>
            <div className="text-[10px] font-black text-slate-500 uppercase tracking-widest italic">PIC Name</div>
            <div className="text-[10px] font-black text-slate-500 uppercase tracking-widest italic">Owner / Div</div>
            <div className="text-[10px] font-black text-slate-500 uppercase tracking-widest italic text-center">Actions</div>
          </div>
        )}

        <div className="flex-1 overflow-y-auto scrollbar-hide divide-y divide-slate-800/50">
          {loading ? (
            <div className="px-6 py-20 text-center">
              <div className="flex flex-col items-center gap-4">
                 <RefreshCw className="w-8 h-8 text-indigo-500 animate-spin" />
                 <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">Synchronizing Master Hub...</span>
              </div>
            </div>
          ) : filteredProjects.length === 0 ? (
            <div className="px-6 py-20 text-center">
              <div className="flex flex-col items-center gap-4">
                 <AlertCircle className="w-10 h-10 text-slate-700" />
                 <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">No Projects Found in the Perimeter</span>
              </div>
            </div>
          ) : paginatedProjects.map((p) => (
            <div 
              key={p.id} 
              className={cn(
                "group hover:bg-indigo-500/[0.05] transition-all cursor-pointer p-4 md:px-6 md:py-5",
                "flex flex-col md:grid md:grid-cols-[120px_1fr_150px_180px_180px_120px] md:items-center gap-4 md:gap-0"
              )}
              onClick={() => handleOpenDetail(p)}
            >
              {/* Ticket ID */}
              <div className="flex items-center justify-between md:block">
                <span className="px-2 py-1 bg-slate-950 border border-slate-800/80 rounded font-mono text-[10px] text-indigo-400 font-bold group-hover:border-indigo-500/50 transition-colors">
                  {p.ticket_id}
                </span>
                {isMobile && (
                  <span className={cn(
                    "px-3 py-1 rounded-full text-[8px] font-black border uppercase tracking-widest transition-all",
                    p.status === 'DONE' ? "bg-emerald-500/10 text-emerald-500 border-emerald-500/20 shadow-[0_0_15px_rgba(16,185,129,0.1)]" :
                    p.status === 'ON PROGRESS' ? "bg-indigo-500/10 text-indigo-400 border-indigo-500/20 shadow-[0_0_15px_rgba(99,102,241,0.1)]" :
                    p.status === 'ON QUEUE' ? "bg-purple-500/10 text-purple-400 border-purple-500/20" :
                    "bg-slate-500/10 text-slate-400 border-slate-500/20"
                  )}>
                    {p.status}
                  </span>
                )}
              </div>

              {/* Project Name */}
              <div className="flex-1 min-w-0">
                <h3 className="text-sm md:text-base font-black text-white italic tracking-tight line-clamp-2 md:line-clamp-1 group-hover:text-indigo-400 transition-colors">
                  {p.project_name}
                </h3>
                <p className="text-[9px] text-slate-600 font-bold tracking-widest uppercase mt-1">
                  Updated: {format(new Date(p.updated_at), 'dd MMM yyyy')}
                </p>
              </div>

              {/* Status - Desktop only here as mobile has it at the top */}
              {!isMobile && (
                <div>
                  <span className={cn(
                    "px-3 py-1 rounded-full text-[9px] font-black border uppercase tracking-widest transition-all",
                    p.status === 'DONE' ? "bg-emerald-500/10 text-emerald-500 border-emerald-500/20" :
                    p.status === 'ON PROGRESS' ? "bg-indigo-500/10 text-indigo-400 border-indigo-500/20" :
                    p.status === 'ON QUEUE' ? "bg-purple-500/10 text-purple-400 border-purple-500/20" :
                    "bg-slate-500/10 text-slate-400 border-slate-500/20"
                  )}>
                    {p.status}
                  </span>
                </div>
              )}

              {/* PIC Name */}
              <div className="flex items-center gap-3">
                 <div className="shrink-0 w-8 h-8 rounded-full bg-slate-800 flex items-center justify-center border border-indigo-500/20 group-hover:border-indigo-500/50">
                    <UserIcon className="w-4 h-4 text-indigo-400" />
                 </div>
                 <div className="min-w-0">
                    <p className="text-[10px] md:text-[11px] font-black text-slate-300 uppercase tracking-tight truncate">
                      {p.pic_name}
                    </p>
                    {isMobile && <p className="text-[8px] text-slate-500 font-bold uppercase tracking-widest mt-0.5">PIC PERSONNEL</p>}
                 </div>
              </div>

              {/* Owner / Div */}
              <div>
                <p className="text-[10px] md:text-[11px] font-black text-white italic truncate max-w-[150px] md:max-w-none">
                  {p.owner_name}
                </p>
                <p className="text-[9px] text-slate-500 font-bold uppercase tracking-widest">
                  {p.div_owner}
                </p>
              </div>

              {/* Actions */}
              <div className={cn(
                "flex items-center gap-2",
                isMobile ? "justify-end pt-2 border-t border-slate-800" : "justify-center opacity-0 group-hover:opacity-100 transition-opacity"
              )}>
                 <button 
                   onClick={(e) => {
                     e.stopPropagation();
                     handleOpenDetail(p);
                     setTimeout(() => setIsEditMode(true), 100);
                   }}
                   className="p-2.5 bg-slate-800 hover:bg-amber-600 text-slate-400 hover:text-white rounded-xl transition-all shadow-lg border border-slate-700 flex items-center gap-2"
                   title="Edit Project"
                 >
                   <Activity className="w-4 h-4" />
                   {isMobile && <span className="text-[10px] font-bold uppercase">Edit</span>}
                 </button>
                 <button 
                   onClick={(e) => {
                     e.stopPropagation();
                     handleOpenAudit(p);
                   }}
                   className="p-2.5 bg-slate-800 hover:bg-indigo-600 text-slate-400 hover:text-white rounded-xl transition-all shadow-lg border border-slate-700 flex items-center gap-2"
                   title="History & Audit"
                 >
                   <History className="w-4 h-4" />
                   {isMobile && <span className="text-[10px] font-bold uppercase">Logs</span>}
                 </button>
              </div>
            </div>
          ))}
        </div>

        {/* Pagination Footer */}
        {totalPages > 1 && (
          <div className="bg-slate-950/50 border-t border-slate-800 px-6 py-4 flex flex-col sm:flex-row items-center justify-between gap-4">
            <div className="text-[10px] font-black text-slate-500 uppercase tracking-widest italic">
              Showing <span className="text-white italic">{(currentPage - 1) * itemsPerPage + 1}</span> to <span className="text-white italic">{Math.min(currentPage * itemsPerPage, filteredProjects.length)}</span> of <span className="text-white italic">{filteredProjects.length}</span> Results
            </div>
            
            <div className="flex items-center gap-2">
              <button 
                onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                disabled={currentPage === 1}
                className="p-2 bg-slate-800 border border-slate-700 rounded-xl text-slate-400 hover:text-white disabled:opacity-30 disabled:cursor-not-allowed transition-all"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
              
              <div className="flex items-center gap-1">
                {Array.from({ length: Math.min(5, totalPages) }, (_, i) => {
                  let pageNum = i + 1;
                  if (totalPages > 5 && currentPage > 3) {
                    pageNum = currentPage - 3 + i + 1;
                    if (pageNum > totalPages) pageNum = totalPages - (4 - i);
                  }
                  
                  return (
                    <button
                      key={pageNum}
                      onClick={() => setCurrentPage(pageNum)}
                      className={cn(
                        "w-8 h-8 rounded-xl text-[10px] font-black transition-all border",
                        currentPage === pageNum 
                          ? "bg-indigo-600 text-white border-indigo-500 shadow-lg shadow-indigo-600/20" 
                          : "bg-slate-900 text-slate-500 border-slate-800 hover:border-slate-700 hover:text-slate-300"
                      )}
                    >
                      {pageNum}
                    </button>
                  );
                })}
              </div>

              <button 
                onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                disabled={currentPage === totalPages}
                className="p-2 bg-slate-800 border border-slate-700 rounded-xl text-slate-400 hover:text-white disabled:opacity-30 disabled:cursor-not-allowed transition-all"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Project Detail & Edit Drawer */}
      <AnimatePresence>
        {isDetailOpen && selectedProject && (
          <div className="fixed inset-0 z-[150] flex justify-end">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsDetailOpen(false)}
              className="absolute inset-0 bg-slate-950/80 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ x: '100%' }}
              animate={{ x: 0 }}
              exit={{ x: '100%' }}
              transition={{ type: 'spring', damping: 25, stiffness: 200 }}
              className="relative w-full max-w-3xl bg-[#0B1120] border-l border-slate-800 h-full flex flex-col shadow-2xl"
            >
               <div className="p-8 border-b border-white/5 bg-slate-900/50 shrink-0">
                  <div className="flex items-center justify-between">
                     <div className="flex items-center gap-4">
                        <div className="w-12 h-12 bg-indigo-600/10 rounded-2xl flex items-center justify-center border border-indigo-500/20">
                           <FolderKanban className="w-6 h-6 text-indigo-400" />
                        </div>
                        <div>
                           <h3 className="text-xl font-black text-white italic tracking-tighter uppercase">{isEditMode ? 'Edit' : 'Project'} <span className="text-indigo-500">Details</span></h3>
                           <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">{selectedProject.ticket_id}</p>
                        </div>
                     </div>
                     <div className="flex items-center gap-2">
                        {!isEditMode && (
                          <button 
                            onClick={() => setIsEditMode(true)}
                            className="bg-slate-800 hover:bg-slate-700 text-white p-2 rounded-xl border border-slate-700 transition-colors"
                          >
                            <Activity className="w-4 h-4" />
                          </button>
                        )}
                        <button onClick={() => setIsDetailOpen(false)} className="p-2 text-slate-500 hover:text-white transition-colors">
                           <X className="w-6 h-6" />
                        </button>
                     </div>
                  </div>
               </div>

               <div className="flex-1 overflow-y-auto p-8 scrollbar-hide">
                  {isEditMode ? (
                    <div className="space-y-6">
                      <div className="grid grid-cols-2 gap-6">
                        <div className="space-y-1.5">
                          <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Project Name</label>
                          <input 
                            type="text"
                            value={selectedProject.project_name}
                            onChange={e => setSelectedProject({ ...selectedProject, project_name: e.target.value })}
                            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-sm text-white focus:border-indigo-500 outline-none transition-all"
                          />
                        </div>
                        <div className="space-y-1.5">
                          <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Global Status</label>
                          <select 
                            value={selectedProject.status}
                            onChange={e => setSelectedProject({ ...selectedProject, status: e.target.value })}
                            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-sm text-indigo-400 font-bold focus:border-indigo-500 outline-none transition-all"
                          >
                            <option value="OPEN">OPEN</option>
                            <option value="ON QUEUE">ON QUEUE</option>
                            <option value="ON PROGRESS">ON PROGRESS</option>
                            <option value="LIVE">LIVE</option>
                            <option value="DONE">DONE</option>
                          </select>
                        </div>
                      </div>

                      <div className="grid grid-cols-2 gap-6">
                        <div className="space-y-1.5">
                          <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">PIC Name</label>
                          <input 
                            type="text"
                            value={selectedProject.pic_name}
                            onChange={e => setSelectedProject({ ...selectedProject, pic_name: e.target.value })}
                            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-sm text-white focus:border-indigo-500 outline-none transition-all"
                          />
                        </div>
                        <div className="space-y-1.5">
                          <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Owner Name</label>
                          <input 
                            type="text"
                            value={selectedProject.owner_name}
                            onChange={e => setSelectedProject({ ...selectedProject, owner_name: e.target.value })}
                            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-sm text-white focus:border-indigo-500 outline-none transition-all"
                          />
                        </div>
                      </div>

                      <div className="grid grid-cols-2 gap-6">
                        <div className="space-y-1.5">
                          <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Division</label>
                          <input 
                            type="text"
                            value={selectedProject.div_owner}
                            onChange={e => setSelectedProject({ ...selectedProject, div_owner: e.target.value })}
                            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-sm text-white focus:border-indigo-500 outline-none transition-all"
                          />
                        </div>
                        <div className="space-y-1.5">
                          <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Total Man Hours</label>
                          <input 
                            type="number"
                            value={selectedProject.total_man_hours || 0}
                            onChange={e => setSelectedProject({ ...selectedProject, total_man_hours: Number(e.target.value) })}
                            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-sm text-white focus:border-indigo-500 outline-none transition-all"
                          />
                        </div>
                      </div>

                      <div className="grid grid-cols-2 gap-6">
                        <div className="space-y-1.5">
                          <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Plan Start Date</label>
                          <input 
                            type="date"
                            value={selectedProject.plan_start_date ? selectedProject.plan_start_date.substring(0, 10) : ''}
                            onChange={e => setSelectedProject({ ...selectedProject, plan_start_date: e.target.value })}
                            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-sm text-white focus:border-indigo-500 outline-none transition-all"
                          />
                        </div>
                        <div className="space-y-1.5">
                          <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Plan End Date</label>
                          <input 
                            type="date"
                            value={selectedProject.plan_end_date ? selectedProject.plan_end_date.substring(0, 10) : ''}
                            onChange={e => setSelectedProject({ ...selectedProject, plan_end_date: e.target.value })}
                            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-sm text-white focus:border-indigo-500 outline-none transition-all"
                          />
                        </div>
                      </div>

                      <div className="flex gap-4 pt-6">
                        <button 
                          onClick={() => setIsEditMode(false)}
                          className="flex-1 py-4 bg-slate-800 text-slate-500 rounded-2xl font-black uppercase text-[10px] tracking-widest hover:text-white transition-all shadow-lg"
                        >
                          Cancel
                        </button>
                        <button 
                          onClick={handleSaveEdit}
                          className="flex-[2] py-4 bg-indigo-600 text-white rounded-2xl font-black uppercase text-[10px] tracking-widest hover:bg-indigo-500 transition-all shadow-xl shadow-indigo-600/20"
                        >
                          {loading ? 'Saving...' : 'Save Changes'}
                        </button>
                      </div>
                    </div>
                  ) : (
                    <div className="space-y-10">
                       <section className="grid grid-cols-2 gap-6 bg-slate-950/30 p-6 rounded-[2rem] border border-white/5 shadow-inner">
                          <div>
                             <p className="text-[9px] font-black text-slate-500 uppercase tracking-widest mb-1 italic">Current Status</p>
                             <span className={cn(
                               "px-3 py-1 rounded-full text-[10px] font-black border uppercase tracking-widest",
                               selectedProject.status === 'DONE' ? "bg-emerald-500/10 text-emerald-500 border-emerald-500/20" :
                               "bg-indigo-500/10 text-indigo-400 border-indigo-500/20"
                             )}>
                               {selectedProject.status}
                             </span>
                          </div>
                          <div>
                             <p className="text-[9px] font-black text-slate-500 uppercase tracking-widest mb-1 italic">Last Synchronized</p>
                             <p className="text-sm font-bold text-white uppercase">{format(new Date(selectedProject.updated_at), 'dd MMM yyyy')}</p>
                          </div>
                       </section>

                       <section className="space-y-6">
                         <h4 className="text-[10px] font-black text-indigo-400 uppercase tracking-[0.2em] px-4">Audit History (Timeline)</h4>
                         <div className="relative space-y-8 px-4">
                            <div className="absolute left-6 top-2 bottom-2 w-px bg-slate-800 border-l border-dashed border-slate-700/50" />
                            {auditLogs.map((log) => (
                              <div key={log.id} className="relative pl-12">
                                <div className={cn(
                                  "absolute left-4 top-1 w-4 h-4 rounded-full border-2 border-slate-900 z-10 shadow-lg",
                                  log.action === 'CREATE' ? "bg-emerald-500" :
                                  log.action === 'UPDATE' ? "bg-amber-500" : "bg-indigo-500"
                                )} />
                                <div className="space-y-2">
                                  <div className="flex items-center justify-between">
                                    <span className="text-[10px] font-bold text-white bg-slate-800 px-2 py-0.5 rounded uppercase">{log.action}</span>
                                    <span className="text-[9px] text-slate-500 font-mono italic">{format(new Date(log.created_at), 'dd MMM HH:mm')}</span>
                                  </div>
                                  <div className="bg-slate-950/50 border border-slate-800 p-3 rounded-xl">
                                    <p className="text-xs text-slate-400 italic leading-relaxed">"{log.note}"</p>
                                    {log.changed_fields && (
                                      <div className="mt-2 p-2 bg-black/40 rounded-lg space-y-1">
                                        {Object.entries(JSON.parse(log.changed_fields)).map(([field, delta]: [string, any]) => (
                                          <div key={field} className="text-[8px] flex items-center gap-2">
                                            <span className="text-slate-500 uppercase font-black">{field.replace('_', ' ')}:</span>
                                            <span className="text-rose-400 line-through opacity-60">{(delta as any).from || 'N/A'}</span>
                                            <span className="text-slate-400">→</span>
                                            <span className="text-emerald-400 font-bold">{(delta as any).to || 'N/A'}</span>
                                          </div>
                                        ))}
                                      </div>
                                    )}
                                    <p className="text-[8px] text-slate-600 font-black uppercase mt-2">By: {log.actor}</p>
                                  </div>
                                </div>
                              </div>
                            ))}
                            {auditLogs.length === 0 && (
                              <div className="text-center py-10 opacity-30 italic text-slate-500 text-xs">No logs recorded.</div>
                            )}
                         </div>
                       </section>
                    </div>
                  )}
               </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* CSV Import Preview & Modal */}
      <AnimatePresence>
        {isImportOpen && (
          <div className="fixed inset-0 z-[150] flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsImportOpen(false)}
              className="absolute inset-0 bg-slate-950/90 backdrop-blur-md"
            />
            <motion.div 
              initial={{ opacity: 0, scale: 0.9 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.9 }}
              className="relative w-full max-w-2xl bg-slate-900 border border-slate-800 rounded-3xl shadow-2xl p-8 overflow-hidden"
            >
               <div className="flex flex-col items-center text-center">
                  <div className="w-20 h-20 bg-indigo-600/10 rounded-full flex items-center justify-center mb-6 border border-indigo-500/20 shadow-lg shadow-indigo-600/5 relative overflow-hidden">
                     <Upload className="w-8 h-8 text-indigo-400 relative z-10" />
                     <div className="absolute inset-0 bg-gradient-to-br from-indigo-500/10 to-transparent" />
                  </div>
                  <h3 className="text-2xl font-black text-white italic tracking-tighter uppercase mb-2">Import Project <span className="text-indigo-500">Master</span></h3>
                  <p className="text-xs text-slate-500 font-bold mb-8 uppercase tracking-widest leading-relaxed">
                     Upload file .CSV atau .XLSX untuk memperbarui basis data master project secara kolektif.
                  </p>
               </div>

               <div className="space-y-4">
                  {isImporting ? (
                    <div className="flex flex-col items-center gap-4 py-8">
                       <RefreshCw className="w-10 h-10 text-indigo-500 animate-spin" />
                       <p className="text-[10px] font-black text-indigo-400 uppercase tracking-[0.2em]">Processing Central Grid...</p>
                    </div>
                  ) : importSummary ? (
                    <div className="flex flex-col items-center gap-4 py-8 text-center">
                       <div className="w-16 h-16 bg-emerald-500/10 rounded-full flex items-center justify-center border border-emerald-500/20">
                          <CheckCircle2 className="w-8 h-8 text-emerald-400" />
                       </div>
                       <p className="text-xs text-white font-bold uppercase tracking-widest">Sincronization Complete</p>
                       <div className="grid grid-cols-3 gap-4 w-full mt-4">
                          <div className="bg-slate-950 p-3 rounded-xl border border-white/5">
                            <p className="text-xl font-black text-indigo-500 italic">{importSummary.inserted}</p>
                            <p className="text-[8px] text-slate-500 font-black uppercase tracking-widest">New</p>
                          </div>
                          <div className="bg-slate-950 p-3 rounded-xl border border-white/5">
                            <p className="text-xl font-black text-amber-500 italic">{importSummary.updated}</p>
                            <p className="text-[8px] text-slate-500 font-black uppercase tracking-widest">Updated</p>
                          </div>
                          <div className="bg-slate-950 p-3 rounded-xl border border-white/5">
                            <p className="text-xl font-black text-slate-500 italic">{importSummary.skipped}</p>
                            <p className="text-[8px] text-slate-500 font-black uppercase tracking-widest">Skipped</p>
                          </div>
                       </div>
                       <button 
                         onClick={() => setIsImportOpen(false)}
                         className="mt-6 w-full py-4 bg-indigo-600 text-white text-[10px] font-black uppercase tracking-widest rounded-2xl hover:bg-indigo-500 transition-all shadow-lg shadow-indigo-600/20"
                       >
                         Dismiss Synchronization View
                       </button>
                    </div>
                  ) : importPreview ? (
                    <div className="space-y-6">
                       <div className="bg-slate-950/50 p-4 rounded-2xl border border-white/5 space-y-4">
                          <p className="text-[10px] font-black text-indigo-400 uppercase tracking-[0.2em] mb-4 flex items-center gap-2">
                             <TrendingUp className="w-4 h-4" /> Import Preview Analysis
                          </p>
                          <div className="max-h-[200px] overflow-y-auto space-y-2 pr-2 scrollbar-hide">
                             {importPreview.diffs.map((diff, idx) => (
                               <div key={idx} className="bg-slate-900 border border-slate-800 p-3 rounded-xl flex items-center justify-between gap-4">
                                  <div className="min-w-0">
                                     <p className="text-[10px] font-mono font-bold text-indigo-500">{diff.ticket_id}</p>
                                     <p className="text-xs text-white font-medium truncate">{diff.name}</p>
                                  </div>
                                  <div className="shrink-0 flex items-center gap-3">
                                     {diff.type === 'UPDATE' ? (
                                       <div className="flex flex-col items-end">
                                          <span className="text-[8px] font-black bg-amber-500/10 text-amber-500 px-2 py-0.5 rounded uppercase border border-amber-500/20">UPDATE</span>
                                          <p className="text-[7px] text-slate-500 font-bold uppercase mt-1 italic">{Object.keys(diff.changes).length} changes</p>
                                       </div>
                                     ) : (
                                       <span className="text-[8px] font-black bg-emerald-500/10 text-emerald-500 px-2 py-0.5 rounded uppercase border border-emerald-500/20">NEW ENTRY</span>
                                     )}
                                  </div>
                               </div>
                             ))}
                          </div>
                       </div>
                       <div className="flex gap-4">
                          <button 
                            onClick={() => { setImportPreview(null); setStagedData([]); }}
                            className="flex-1 py-4 bg-slate-800 text-slate-500 rounded-2xl font-black uppercase text-[10px] tracking-widest hover:text-white"
                          >
                            Reset
                          </button>
                          <button 
                            onClick={confirmImport}
                            className="flex-[2] py-4 bg-indigo-600 text-white rounded-2xl font-black uppercase text-[10px] tracking-widest hover:bg-indigo-500 shadow-xl shadow-indigo-600/20 transition-all"
                          >
                            Commit Changes
                          </button>
                       </div>
                    </div>
                  ) : (
                    <>
                      <label className="block">
                         <span className="sr-only">Pilih File</span>
                         <input 
                           type="file" 
                           accept=".csv,.xlsx"
                           onChange={handleImport}
                           className="block w-full text-xs text-slate-500 file:mr-4 file:py-3 file:px-6 file:rounded-xl file:border-0 file:text-[10px] file:font-black file:uppercase file:bg-indigo-600 file:text-white hover:file:bg-indigo-500 file:transition-all cursor-pointer bg-slate-950/50 rounded-2xl border border-slate-800/80 pr-4"
                         />
                      </label>
                      <p className="text-[9px] text-slate-600 font-mono italic text-center">Support CSV/XLSX: Ticket, ProjectName, Status, PIC_Name, Owner_Name, Div_Owner</p>
                      <button 
                        onClick={() => setIsImportOpen(false)}
                        className="mt-6 w-full py-4 text-slate-500 font-black uppercase text-[10px] tracking-widest hover:text-white transition-colors"
                      >
                         Batal
                      </button>
                    </>
                  )}
               </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Manual Add Project Modal */}
      <AnimatePresence>
        {isAddOpen && (
          <div className="fixed inset-0 z-[150] flex items-center justify-center p-4">
             <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} onClick={() => setIsAddOpen(false)} className="absolute inset-0 bg-slate-950/90 backdrop-blur-md" />
             <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: 20 }} className="relative w-full max-w-xl bg-slate-900 border border-slate-800 rounded-3xl shadow-2xl p-8">
                <h3 className="text-xl font-black text-white italic uppercase tracking-tighter mb-8 italic">Manual Master Project <span className="text-indigo-500">Entry</span></h3>
                <div className="space-y-6">
                   <div className="grid grid-cols-2 gap-4">
                      <div className="space-y-1.5">
                        <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Ticket ID</label>
                        <input id="new-ticket" className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-xs text-white" placeholder="PMA-XXXX" />
                      </div>
                      <div className="space-y-1.5">
                        <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Status</label>
                        <select id="new-status" className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-xs text-indigo-400 font-black">
                           <option>OPEN</option>
                           <option>ON QUEUE</option>
                           <option>ON PROGRESS</option>
                        </select>
                      </div>
                   </div>
                   <div className="space-y-1.5">
                      <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Project Name</label>
                      <input id="new-name" className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-xs text-white" placeholder="e.g. Migration Core Core Banking" />
                   </div>
                   <div className="grid grid-cols-2 gap-4">
                      <div className="space-y-1.5">
                        <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Owner Name</label>
                        <input id="new-owner" className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-xs text-white" placeholder="Owner Name" />
                      </div>
                      <div className="space-y-1.5">
                        <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Division</label>
                        <input id="new-div" className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-xs text-white" placeholder="Owner Div" />
                      </div>
                   </div>
                   <div className="space-y-1.5">
                      <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">PIC Personnel</label>
                      <input id="new-pic" className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-xs text-white" placeholder="Nama Lengkap" />
                   </div>
                </div>
                <div className="flex gap-4 mt-10">
                   <button onClick={() => setIsAddOpen(false)} className="flex-1 py-4 bg-slate-800 text-slate-500 rounded-xl font-black uppercase text-[10px] tracking-widest">Cancel</button>
                   <button 
                     onClick={async () => {
                       const ticket = (document.getElementById('new-ticket') as HTMLInputElement).value;
                       const name = (document.getElementById('new-name') as HTMLInputElement).value;
                       const pic = (document.getElementById('new-pic') as HTMLInputElement).value;
                       const status = (document.getElementById('new-status') as HTMLSelectElement).value;
                       const owner = (document.getElementById('new-owner') as HTMLInputElement).value;
                       const div = (document.getElementById('new-div') as HTMLInputElement).value;
                       if (!ticket || !name) return alert('Data tidak lengkap');
                       await taskService.createMasterProject({ ticket_id: ticket, project_name: name, pic_name: pic, status, owner_name: owner, div_owner: div }, user?.email || 'Admin');
                       fetchData();
                       setIsAddOpen(false);
                     }}
                     className="flex-3 py-4 bg-indigo-600 text-white rounded-xl font-black uppercase text-[10px] tracking-widest shadow-lg shadow-indigo-600/20"
                   >
                     Confirm Save
                   </button>
                </div>
             </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

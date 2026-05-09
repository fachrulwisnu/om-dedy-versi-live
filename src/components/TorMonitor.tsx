import React, { useState, useEffect } from 'react';
import { Project, Task, MasterProject } from '../types';
import { supabase } from '../lib/supabase';
import { getL2PrecisionStatus, TorStatus, calculateGlobalProjectStatus } from '../services/torMonitorService';
import { cn } from '../lib/utils';
import { 
  Rocket, 
  ChevronRight, 
  Activity, 
  AlertCircle, 
  CheckCircle2, 
  Clock,
  RefreshCw,
  Search,
  Settings,
  ShieldCheck
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

interface ProjectMonitorRow extends Project {
  currentPhase?: string;
  globalStatusLabel: string;
  isAutoStatus?: boolean;
}

export function TorMonitor({ user }: { user: any }) {
  const [projects, setProjects] = useState<Project[]>([]);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [masterProjects, setMasterProjects] = useState<MasterProject[]>([]);
  const [loading, setLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState('ALL');
  const [picFilter, setPicFilter] = useState('ALL');
  const [autoStatusSettings, setAutoStatusSettings] = useState<Record<string, boolean>>({});
  const [auditLogs, setAuditLogs] = useState<any[]>([]);
  const [selectedProjectLogs, setSelectedProjectLogs] = useState<{name: string, logs: any[]} | null>(null);

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    setLoading(true);
    try {
      const { data: prjs } = await supabase.from('projects').select('*').order('created_at', { ascending: false });
      const { data: tsk } = await supabase.from('tasks').select('*');
      const { data: masters } = await supabase.from('master_projects').select('*');
      const { data: logs } = await supabase.from('master_project_audit_logs').select('*').order('created_at', { ascending: false });
      
      setProjects(prjs || []);
      setTasks(tsk || []);
      setMasterProjects(masters || []);
      setAuditLogs(logs || []);
      
      const saved = localStorage.getItem('od_auto_status_configs');
      if (saved) {
        setAutoStatusSettings(JSON.parse(saved));
      } else {
        const initial: Record<string, boolean> = {};
        prjs?.forEach(p => initial[p.id] = true);
        setAutoStatusSettings(initial);
      }
    } catch (err) {
      console.error('Fetch failed:', err);
    } finally {
      setLoading(false);
    }
  };

  const toggleAutoStatus = (projectId: string) => {
    const newSettings = {
      ...autoStatusSettings,
      [projectId]: !autoStatusSettings[projectId]
    };
    setAutoStatusSettings(newSettings);
    localStorage.setItem('od_auto_status_configs', JSON.stringify(newSettings));
  };

  const getCalculatedGlobalStatus = (p: Project, pTasks: Task[]) => {
    // Priority: If master_projects has it, use it (it's the source of truth for global)
    const master = masterProjects.find(m => m.ticket_id === p.ticket_id);
    if (master && master.status) return master.status;

    return calculateGlobalProjectStatus(pTasks);
  };

  const getProjectProgress = (pTasks: Task[]) => {
    if (pTasks.length === 0) return 0;
    const finished = pTasks.filter(t => !!t.realized_finish_date).length;
    return Math.round((finished / pTasks.length) * 100);
  };

  const processedData: ProjectMonitorRow[] = projects.map(p => {
    const projectTasks = tasks.filter(t => t.project_id === p.id);
    const isAuto = autoStatusSettings[p.id] ?? true;
    
    const globalStatus = isAuto ? getCalculatedGlobalStatus(p, projectTasks) : p.status;
    const progress = getProjectProgress(projectTasks);
    
    const l1Tasks = projectTasks.filter(t => !t.parent_id).sort((a,b) => new Date(a.start_time).getTime() - new Date(b.start_time).getTime());
    const currentPhaseTask = l1Tasks.find(t => !t.realized_finish_date) || l1Tasks[l1Tasks.length - 1];

    return {
      ...p,
      currentPhase: currentPhaseTask?.title || 'To Do',
      globalStatusLabel: globalStatus,
      progress,
      isAutoStatus: isAuto
    } as any;
  }).filter(p => {
    const matchesSearch = p.name.toLowerCase().includes(searchTerm.toLowerCase()) || p.ticket_id?.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesStatus = statusFilter === 'ALL' || p.globalStatusLabel === statusFilter;
    const matchesPIC = picFilter === 'ALL' || p.pic_name === picFilter;
    return matchesSearch && matchesStatus && matchesPIC;
  });

  // Summary Counts
  const summary = {
    todo: processedData.filter(p => p.globalStatusLabel === 'On Queue' || p.globalStatusLabel === 'To Do').length,
    inProgress: processedData.filter(p => p.globalStatusLabel.includes('On Progress')).length,
    overdue: processedData.filter(p => p.globalStatusLabel === 'Project Late' || p.globalStatusLabel.toLowerCase().includes('overdue')).length,
    live: processedData.filter(p => p.globalStatusLabel === 'LIVE' || p.globalStatusLabel === 'Done').length
  };

  const uniquePICs = Array.from(new Set(projects.map(p => p.pic_name).filter(Boolean)));

  const getStatusColor = (status: string) => {
    if (status.includes('Overdue') || status === 'Project Late') return 'bg-rose-500/10 text-rose-500 border-rose-500/20';
    if (status === 'LIVE' || status === 'Done' || status.includes('Early')) return 'bg-emerald-500/10 text-emerald-500 border-emerald-500/20';
    if (status.includes('On Progress') || status === 'In Progress') return 'bg-blue-500/10 text-blue-400 border-blue-500/20';
    return 'bg-slate-500/10 text-slate-400 border-slate-500/20';
  };

  const openAuditLogs = (p: Project) => {
    const master = masterProjects.find(m => m.ticket_id === p.ticket_id);
    const logs = master ? auditLogs.filter(l => l.master_project_id === master.id) : [];
    setSelectedProjectLogs({ name: p.name, logs });
  };

  return (
    <div className="p-8 max-w-7xl mx-auto space-y-12 min-h-screen pb-32">
      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-end justify-between gap-8">
        <div className="space-y-4">
          <div className="flex items-center gap-4">
            <div className="w-14 h-14 bg-indigo-600 rounded-[2rem] flex items-center justify-center shadow-2xl shadow-indigo-500/40 rotate-6 hover:rotate-0 transition-transform duration-500">
              <Activity className="w-7 h-7 text-white" />
            </div>
            <div>
              <h1 className="text-4xl font-black text-white uppercase tracking-tighter italic leading-none">
                OM DEDY <span className="text-indigo-500">BOX MONITOR</span>
              </h1>
              <p className="text-[11px] text-slate-500 font-bold uppercase tracking-[0.3em] mt-2">
                Predictive & Reactive Operational Logic
              </p>
            </div>
          </div>
        </div>

        {/* Global Filters */}
        <div className="flex flex-wrap gap-4 items-center">
          <div className="relative group">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500 group-focus-within:text-indigo-500 transition-colors" />
            <input 
              type="text"
              placeholder="Project Name / Ticket..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="bg-slate-950 border border-slate-800 rounded-2xl pl-10 pr-4 py-3 text-xs font-bold text-white focus:border-indigo-500 outline-none transition-all w-64 shadow-inner"
            />
          </div>

          <select 
            value={picFilter}
            onChange={(e) => setPicFilter(e.target.value)}
            className="bg-slate-950 border border-slate-800 rounded-2xl px-4 py-3 text-xs font-bold text-white focus:border-indigo-500 outline-none shadow-inner min-w-[140px]"
          >
            <option value="ALL">ALL PIC</option>
            {uniquePICs.map(pic => <option key={pic} value={pic}>{pic}</option>)}
          </select>

          <select 
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
            className="bg-slate-950 border border-slate-800 rounded-2xl px-4 py-3 text-xs font-bold text-white focus:border-indigo-500 outline-none shadow-inner min-w-[140px]"
          >
            <option value="ALL">ALL STATUS</option>
            <option value="FSD On Progress">FSD</option>
            <option value="Development On Progress">DEVELOPMENT</option>
            <option value="SIT On Progress">SIT</option>
            <option value="UAT On Progress">UAT</option>
            <option value="LIVE">LIVE</option>
          </select>
        </div>
      </div>

      {/* Summary Dashboard Boxes */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-6">
        {[
          { label: 'To Do / On Queue', value: summary.todo, color: 'slate', icon: Clock },
          { label: 'In Progress', value: summary.inProgress, color: 'blue', icon: RefreshCw },
          { label: 'Overdue / Late', value: summary.overdue, color: 'rose', icon: AlertCircle },
          { label: 'Done / LIVE', value: summary.live, color: 'emerald', icon: ShieldCheck },
        ].map((box, idx) => (
          <motion.div 
            key={idx}
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: idx * 0.1 }}
            className={cn(
              "p-6 rounded-[2.5rem] border flex flex-col justify-between h-40 shadow-xl transition-all hover:scale-105",
              box.color === 'slate' ? "bg-slate-900/50 border-slate-800 shadow-slate-950/50" :
              box.color === 'blue' ? "bg-blue-500/5 border-blue-500/20 shadow-blue-500/10" :
              box.color === 'rose' ? "bg-rose-500/5 border-rose-500/20 shadow-rose-500/10" :
              "bg-emerald-500/5 border-emerald-500/20 shadow-emerald-500/10"
            )}
          >
            <div className="flex items-center justify-between">
              <div className={cn("w-10 h-10 rounded-2xl flex items-center justify-center", 
                box.color === 'slate' ? "bg-slate-800 text-slate-400" :
                box.color === 'blue' ? "bg-blue-500/10 text-blue-500" :
                box.color === 'rose' ? "bg-rose-500/10 text-rose-500" :
                "bg-emerald-500/10 text-emerald-500"
              )}>
                <box.icon className="w-5 h-5" />
              </div>
              <span className="text-3xl font-black italic tracking-tighter opacity-10">{box.value}</span>
            </div>
            <div>
              <p className="text-[10px] font-black uppercase tracking-widest opacity-60 mb-1">{box.label}</p>
              <p className="text-4xl font-black text-white leading-none">{box.value.toString().padStart(2, '0')}</p>
            </div>
          </motion.div>
        ))}
      </div>

      {/* Grid of Project Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-8">
        <AnimatePresence mode="popLayout">
          {processedData.map((project, idx) => (
            <motion.div 
              layout
              initial={{ opacity: 0, scale: 0.9 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.9 }}
              key={project.id}
              className="bg-slate-900/50 border border-slate-800 rounded-[2.5rem] p-8 space-y-8 flex flex-col justify-between hover:border-indigo-500/50 transition-all shadow-2xl group relative overflow-hidden"
            >
              {/* Card Header */}
              <div className="space-y-4 relative z-10">
                <div className="flex items-center justify-between">
                  <span className="text-[10px] font-black font-mono tracking-widest text-indigo-400 uppercase bg-indigo-500/10 px-3 py-1 rounded-full">
                    {project.ticket_id || 'NO TICKET'}
                  </span>
                  <div className={cn(
                    "px-3 py-1 p-0.5 rounded-full border text-[9px] font-black uppercase tracking-[0.2em] shadow-inner",
                    getStatusColor(project.globalStatusLabel)
                  )}>
                    {project.globalStatusLabel}
                  </div>
                </div>
                <h3 className="text-xl font-black text-white tracking-tight uppercase leading-tight group-hover:text-indigo-400 transition-colors">
                  {project.name}
                </h3>
              </div>

              {/* Card Body: Progress & Phases */}
              <div className="space-y-6 relative z-10">
                <div className="space-y-2">
                  <div className="flex items-center justify-between text-[10px] font-black uppercase tracking-widest text-slate-500">
                    <span>Efficiency Rate</span>
                    <span className="text-white">{(project as any).progress}%</span>
                  </div>
                  <div className="h-2 bg-slate-950 rounded-full overflow-hidden shadow-inner border border-slate-800/50">
                    <motion.div 
                      initial={{ width: 0 }}
                      animate={{ width: `${(project as any).progress}%` }}
                      className={cn(
                        "h-full rounded-full shadow-[0_0_10px_rgba(99,102,241,0.5)]",
                        (project as any).progress === 100 ? "bg-emerald-500 shadow-emerald-500/50" : "bg-indigo-500"
                      )}
                    />
                  </div>
                </div>

                <div className="flex flex-wrap gap-2">
                  {['FSD', 'DEV', 'SIT', 'UAT'].map(phase => {
                    const projectTasks = tasks.filter(t => t.project_id === project.id);
                    const l1Phase = projectTasks.find(t => !t.parent_id && t.title.toUpperCase().includes(phase));
                    const isDone = !!l1Phase?.realized_finish_date;
                    
                    return (
                      <div 
                        key={phase}
                        className={cn(
                          "px-2.5 py-1.5 rounded-xl border text-[8px] font-black tracking-widest transition-all",
                          isDone 
                            ? "bg-emerald-500/10 border-emerald-500/30 text-emerald-400" 
                            : l1Phase 
                              ? "bg-slate-800 border-slate-700 text-slate-500" 
                              : "bg-slate-950/50 border-slate-900 text-slate-800"
                        )}
                      >
                        {phase}
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Card Footer */}
              <div className="pt-8 border-t border-slate-800 flex items-center justify-between relative z-10">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-2xl bg-slate-800 flex items-center justify-center border border-slate-700 overflow-hidden">
                    <div className="text-xs font-black uppercase text-indigo-500">
                      {project.pic_name?.substring(0, 2) || '??'}
                    </div>
                  </div>
                  <div>
                    <p className="text-[10px] font-black text-white uppercase tracking-tight leading-none">{project.pic_name}</p>
                    <p className="text-[8px] text-slate-500 font-bold uppercase tracking-widest mt-1">{project.div_owner || 'OM DEDY'}</p>
                  </div>
                </div>

                <div className="flex items-center gap-3">
                  {/* Auto-Toggle in card */}
                  <button 
                    onClick={(e) => { e.stopPropagation(); toggleAutoStatus(project.id); }}
                    title={project.isAutoStatus ? "Switch to Manual Mode" : "Switch to Auto Status"}
                    className={cn(
                      "p-2 rounded-xl border transition-all",
                      project.isAutoStatus ? "bg-indigo-500/10 border-indigo-500/30 text-indigo-500" : "bg-slate-800 border-slate-700 text-slate-500"
                    )}
                  >
                    <Settings className="w-4 h-4" />
                  </button>

                  <button 
                    onClick={() => openAuditLogs(project)}
                    className="p-3 bg-slate-950 border border-slate-800 rounded-2xl hover:border-indigo-500 transition-all shadow-inner group/btn"
                  >
                    <ChevronRight className="w-4 h-4 text-slate-500 group-hover/btn:text-white group-hover/btn:translate-x-0.5 transition-all" />
                  </button>
                </div>
              </div>

              {/* Background Decoration */}
              <div className="absolute top-0 right-0 -mr-16 -mt-16 w-48 h-48 bg-indigo-500/5 blur-[100px] rounded-full" />
            </motion.div>
          ))}
        </AnimatePresence>
      </div>

      {/* Audit Log Modal */}
      <AnimatePresence>
        {selectedProjectLogs && (
          <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setSelectedProjectLogs(null)}
              className="absolute inset-0 bg-slate-950/80 backdrop-blur-md"
            />
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative w-full max-w-2xl bg-slate-900 border border-slate-800 rounded-[3rem] shadow-2xl overflow-hidden flex flex-col max-h-[80vh]"
            >
              <div className="p-8 border-b border-white/5 flex items-center justify-between">
                <div>
                  <h2 className="text-xl font-black text-white uppercase italic tracking-tight">{selectedProjectLogs.name}</h2>
                  <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest mt-1">Audit Trail & Decision Logs</p>
                </div>
                <button onClick={() => setSelectedProjectLogs(null)} className="p-3 bg-slate-950 border border-slate-800 rounded-2xl text-slate-500 hover:text-white transition-colors">
                  CLOSE
                </button>
              </div>
              <div className="flex-1 overflow-y-auto p-8 space-y-6">
                {selectedProjectLogs.logs.length === 0 ? (
                  <div className="text-center py-20 opacity-30 italic text-slate-500">No logs recorded yet.</div>
                ) : (
                  selectedProjectLogs.logs.map((log: any) => (
                    <div key={log.id} className="relative pl-10 group/log">
                      <div className="absolute left-0 top-0 bottom-0 w-px bg-slate-800 group-last/log:bg-transparent" />
                      <div className="absolute left-[-4px] top-1.5 w-2 h-2 rounded-full bg-indigo-500 shadow-[0_0_10px_rgba(99,102,241,0.5)]" />
                      <div className="space-y-2">
                        <div className="flex items-center gap-3">
                          <span className="text-[10px] font-black text-white uppercase tracking-widest bg-slate-800 px-2 py-0.5 rounded">{log.actor}</span>
                          <span className="text-[8px] text-slate-500 font-bold font-mono">{new Date(log.created_at).toLocaleString()}</span>
                        </div>
                        <p className="text-xs font-bold text-slate-300 italic">{log.action}</p>
                        <div className="p-4 bg-slate-950 rounded-2xl border border-slate-800 text-[10px] font-mono text-slate-400">
                          <pre className="whitespace-pre-wrap">{JSON.stringify(log.new_payload, null, 2)}</pre>
                        </div>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

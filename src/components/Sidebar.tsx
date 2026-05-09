import React, { useState, useEffect } from 'react';
import { 
  LayoutDashboard, 
  BarChart3,
  Activity, 
  LayoutGrid, 
  Calendar, 
  History, 
  Users, 
  FolderKanban, 
  ShieldAlert, 
  ShieldCheck,
  LogOut, 
  ChevronLeft, 
  ChevronRight,
  Zap,
  ChevronDown,
  ChevronUp
} from 'lucide-react';
import { SidebarItem } from './SidebarItem';
import { cn } from '../lib/utils';
import { useNavigate, useLocation } from 'react-router-dom';
import { ThemeToggle } from './ThemeToggle';

interface SidebarProps {
  isOpen: boolean;
  setIsOpen: (open: boolean) => void;
  user: any;
  activeView: string;
  onLogout: () => void;
  pendingRescheduleCount: number;
  selectedProjectId?: string | null;
}

export function Sidebar({ 
  isOpen, 
  setIsOpen, 
  user, 
  activeView, 
  onLogout, 
  pendingRescheduleCount,
  selectedProjectId
}: SidebarProps) {
  const navigate = useNavigate();
  const location = useLocation();
  const userRole = user?.access_level || 'PIC';

  const [isProjectsExpanded, setIsProjectsExpanded] = useState(true);

  // Sub-items for Project List
  const projectSubItems = [
    { id: 'TIMELINE', label: 'Om Dedy Timeline', icon: BarChart3, path: '/timeline' },
    ...(selectedProjectId ? [{ 
      id: 'GANTT_DETAIL', 
      label: 'Om Dedy Detail Timeline', 
      icon: Activity, 
      path: `/detail-timeline/${selectedProjectId}`, 
      showDot: true 
    }] : []),
    { id: 'TOR_MONITOR', label: 'Om Dedy Tor Monitor', icon: ShieldCheck, path: '/tor-monitor', roles: ['Admin', 'Superadmin'] },
  ];

  const isSubMenuActive = projectSubItems.some(item => activeView === item.id) || activeView === 'PROJECTS';

  // Auto-expand if sub-menu is active
  useEffect(() => {
    if (isSubMenuActive) {
      setIsProjectsExpanded(true);
    }
  }, [isSubMenuActive]);

  const mainMenuItems = [
    { id: 'KANBAN', label: 'Status Monitoring', icon: LayoutGrid, path: '/kanban' },
    { id: 'SCHEDULE', label: 'Om Dedy Schedule', icon: Calendar, path: '/schedule' },
    { id: 'RESCHEDULE', label: 'APPROVALS', icon: History, path: '/approvals', roles: ['Admin', 'Superadmin'] },
    { id: 'PERSONEL', label: 'Personel OM DEDY', icon: Users, path: '/personnel', roles: ['Admin', 'Superadmin'] },
    { id: 'MASTER_PROJECT', label: 'Om Dedy Master Project', icon: FolderKanban, path: '/master-project', roles: ['Admin', 'Superadmin'] },
    { id: 'AUDIT', label: 'System Audit Logs', icon: ShieldAlert, path: '/audit', roles: ['Admin', 'Superadmin'] },
  ];

  return (
    <aside className={cn(
      "bg-white dark:bg-slate-900 border-r border-slate-200 dark:border-slate-800 transition-all duration-300 flex flex-col z-40 shrink-0 relative",
      isOpen ? "w-72" : "w-16"
    )}>
      {/* Collapse Toggle Button */}
      <button 
        onClick={() => setIsOpen(!isOpen)}
        className="absolute -right-3 top-24 w-6 h-6 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-full flex items-center justify-center z-50 text-slate-400 hover:text-indigo-600 dark:hover:text-white transition-all shadow-xl"
      >
        {isOpen ? <ChevronLeft className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
      </button>

      {/* Brand Header */}
      <div className="h-20 flex items-center justify-between px-4 border-b border-slate-200 dark:border-slate-800 shrink-0 overflow-hidden">
        <div className="flex items-center">
          <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center shrink-0 shadow-lg shadow-indigo-500/20 group cursor-pointer hover:rotate-12 transition-transform">
            <span className="text-lg font-black text-white">OD</span>
          </div>
          {isOpen && (
            <div className="ml-3">
              <h1 className="font-black text-slate-900 dark:text-white uppercase italic tracking-tighter text-xl leading-none">
                OM <span className="text-indigo-500">DEDY</span>
              </h1>
              <p className="text-[7px] text-slate-500 font-black tracking-[0.05em] uppercase leading-tight mt-1">
                Operational Monitoring Dashboard
              </p>
            </div>
          )}
        </div>
        {isOpen && <ThemeToggle />}
      </div>

      {/* Navigation */}
      <nav className="py-6 space-y-1 overflow-y-auto scrollbar-hide">
        {/* Project List Group */}
        <div className="mb-2">
          <div 
            onClick={() => {
              if (isOpen) {
                setIsProjectsExpanded(!isProjectsExpanded);
              } else {
                setIsOpen(true);
                setIsProjectsExpanded(true);
              }
              // Optionally navigate to projects if clicking the parent
              navigate('/projects');
            }}
            className={cn(
              "w-full flex items-center justify-between py-3 px-4 transition-all relative group cursor-pointer",
              isSubMenuActive 
                ? "text-indigo-600 dark:text-white bg-gradient-to-r from-indigo-600/10 dark:from-indigo-600/20 to-transparent" 
                : "text-slate-500 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200 hover:bg-slate-100 dark:hover:bg-white/5"
            )}
          >
            <div className="flex items-center">
              <LayoutDashboard className={cn(
                "w-5 h-5 shrink-0 transition-colors z-10", 
                isSubMenuActive ? "text-indigo-400" : "text-slate-500 group-hover:text-slate-300"
              )} />
              {isOpen && (
                <span className="ml-4 font-bold text-xs uppercase tracking-widest truncate z-10">
                  Project List
                </span>
              )}
            </div>
            {isOpen && (
              isProjectsExpanded ? <ChevronUp className="w-4 h-4 text-slate-500" /> : <ChevronDown className="w-4 h-4 text-slate-500" />
            )}
            
            {isSubMenuActive && (
              <div className="absolute left-0 top-0 bottom-0 w-1 bg-indigo-500 shadow-[0_0_10px_#6366f1]" />
            )}
            
            {!isOpen && (
              <div className="absolute left-full ml-2 px-3 py-1 bg-slate-800 text-white text-[10px] rounded opacity-0 group-hover:opacity-100 whitespace-nowrap shadow-xl pointer-events-none z-50 transition-opacity">
                Project List
              </div>
            )}

            {/* Glow Effect on Hover */}
            <div className="absolute inset-0 bg-indigo-500/0 group-hover:bg-indigo-500/5 transition-colors" />
          </div>

          {isOpen && isProjectsExpanded && (
            <div className="mt-1 ml-4 border-l border-slate-200 dark:border-slate-800 space-y-1">
              {projectSubItems.map((item) => (
                <SidebarItem
                  key={item.id}
                  to={item.path}
                  label={item.label}
                  icon={item.icon}
                  isOpen={isOpen}
                  isActive={activeView === item.id}
                  roles={item.roles}
                  userRole={userRole}
                  showDot={(item as any).showDot}
                />
              ))}
            </div>
          )}
        </div>

        {/* Separator if needed or just space */}
        <div className="h-px bg-slate-200 dark:bg-slate-800/50 mx-4 my-2" />

        {/* Main Menu Items */}
        {mainMenuItems.map((item) => (
          <SidebarItem
            key={item.id}
            to={item.path}
            label={item.label}
            icon={item.icon}
            isOpen={isOpen}
            isActive={activeView === item.id}
            badgeCount={item.id === 'RESCHEDULE' ? pendingRescheduleCount : undefined}
            roles={item.roles}
            userRole={userRole}
            showDot={(item as any).showDot}
          />
        ))}
      </nav>

      {/* User Footer */}
      <div className="p-4 border-t border-slate-200 dark:border-slate-800 grow-0 mt-8">
        {isOpen && user ? (
          <div className="bg-slate-100 dark:bg-slate-950/50 p-4 rounded-xl border border-slate-200 dark:border-slate-800">
            <div className="flex items-center gap-3 mb-3">
              <div className="w-8 h-8 rounded-full bg-white dark:bg-slate-800 border border-slate-200 dark:border-indigo-500/30 flex items-center justify-center overflow-hidden">
                <span className="text-[10px] font-black text-indigo-600 dark:text-indigo-400 capitalize">{user.email?.charAt(0) || '?'}</span>
              </div>
              <div className="flex-1 min-w-0">
                <p className="text-[10px] font-bold text-slate-800 dark:text-slate-200 truncate">{user?.email}</p>
                <p className="text-[9px] text-slate-500 uppercase tracking-widest font-bold">{userRole}</p>
              </div>
            </div>
            <button 
              onClick={onLogout}
              className="w-full flex items-center justify-center gap-2 py-2 bg-white dark:bg-slate-800 hover:bg-rose-500/10 hover:text-rose-600 dark:hover:text-rose-400 rounded-lg text-[10px] font-bold uppercase transition-all border border-slate-200 dark:border-transparent"
            >
              <LogOut className="w-3.5 h-3.5" /> Log Out
            </button>
          </div>
        ) : user ? (
          <div className="flex flex-col items-center gap-4">
            <ThemeToggle />
            <button 
              onClick={onLogout}
              className="w-full flex items-center justify-center p-2 text-slate-500 hover:text-rose-600 dark:hover:text-rose-400 transition-colors"
              title="Log Out"
            >
              <LogOut className="w-5 h-5" />
            </button>
          </div>
        ) : null}
      </div>
    </aside>
  );
}

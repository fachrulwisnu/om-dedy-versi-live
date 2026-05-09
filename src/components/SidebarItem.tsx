import React from 'react';
import { NavLink } from 'react-router-dom';
import { cn } from '../lib/utils';
import { LucideIcon } from 'lucide-react';

interface SidebarItemProps {
  to: string;
  label: string;
  icon: LucideIcon;
  isOpen: boolean;
  isActive: boolean;
  badgeCount?: number;
  roles?: string[];
  userRole?: string;
  onClick?: () => void;
  showDot?: boolean;
}

export function SidebarItem({ 
  to, 
  label, 
  icon: Icon, 
  isOpen, 
  isActive, 
  badgeCount, 
  roles, 
  userRole, 
  onClick,
  showDot
}: SidebarItemProps) {
  // Role Access Check
  const hasAccess = roles ? roles.includes(userRole || '') : true;

  if (!hasAccess) return null;

  return (
    <NavLink
      to={to}
      onClick={(e) => {
        // e.preventDefault(); // FIXED: Removed to allow navigation
        if (onClick) onClick();
      }}
      className={({ isActive: navActive }) => cn(
        "w-full flex items-center py-3 px-4 transition-all relative group overflow-hidden",
        (navActive || isActive)
          ? "text-indigo-600 dark:text-white bg-gradient-to-r from-indigo-600/10 dark:from-indigo-600/20 to-transparent" 
          : "text-slate-500 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200 hover:bg-slate-100 dark:hover:bg-white/5"
      )}
    >
      <Icon className={cn(
        "w-5 h-5 shrink-0 transition-colors z-10", 
        (isActive) ? "text-indigo-600 dark:text-indigo-400" : "text-slate-400 dark:text-slate-500 group-hover:text-indigo-600 dark:group-hover:text-slate-300"
      )} />
      
      {isOpen && (
        <span className="ml-4 font-bold text-xs uppercase tracking-widest truncate z-10 flex items-center gap-2">
          {showDot && <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.5)] animate-pulse" />}
          {label}
        </span>
      )}

      {badgeCount !== undefined && badgeCount > 0 && (
        <span className={cn(
          "ml-auto bg-rose-500 text-white text-[8px] font-black rounded-full flex items-center justify-center animate-bounce shadow-lg ring-2 ring-white dark:ring-slate-900 transition-all z-10",
          isOpen ? "w-5 h-5" : "w-4 h-4 absolute top-1 right-1"
        )}>
          {badgeCount}
        </span>
      )}

      {!isOpen && (
        <div className="absolute left-full ml-2 px-3 py-1 bg-white dark:bg-slate-800 text-slate-900 dark:text-white text-[10px] border border-slate-200 dark:border-slate-700 rounded opacity-0 group-hover:opacity-100 whitespace-nowrap shadow-xl pointer-events-none z-50 transition-opacity">
          {label}
        </div>
      )}

      {(isActive) && (
        <div className="absolute left-0 top-0 bottom-0 w-1 bg-indigo-500 shadow-[0_0_10px_#6366f1]" />
      )}
      
      {/* Glow Effect on Hover */}
      <div className="absolute inset-0 bg-indigo-500/0 group-hover:bg-indigo-500/5 transition-colors" />
    </NavLink>
  );
}

import React from 'react';
import { Sun, Moon } from 'lucide-react';
import { useTheme } from '../hooks/useTheme';

export const ThemeToggle = () => {
  const { theme, toggleTheme } = useTheme();

  return (
    <button
      onClick={toggleTheme}
      className="relative w-10 h-10 flex items-center justify-center rounded-lg bg-slate-800 hover:bg-slate-700 dark:bg-slate-800 dark:hover:bg-slate-700 light:bg-slate-200 light:hover:bg-slate-300 transition-colors duration-300 border border-slate-700 dark:border-slate-700 light:border-slate-300 shadow-lg"
      aria-label="Toggle Theme"
    >
      <Sun 
        className={`absolute w-5 h-5 text-amber-500 transition-all duration-500 transform ${
          theme === 'dark' ? 'opacity-0 rotate-90 scale-50' : 'opacity-100 rotate-0 scale-100'
        }`} 
      />
      <Moon 
        className={`absolute w-5 h-5 text-indigo-400 transition-all duration-500 transform ${
          theme === 'light' ? 'opacity-0 -rotate-90 scale-50' : 'opacity-100 rotate-0 scale-100'
        }`} 
      />
    </button>
  );
};

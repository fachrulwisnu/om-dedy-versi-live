import React, { useState } from 'react';
import { supabase } from '../lib/supabase';
import { motion } from 'motion/react';
import { Cpu, Mail, Lock, AlertCircle } from 'lucide-react';
import { useAuth } from '../context/AuthContext';
import { useNavigate } from 'react-router-dom';

export function LoginPage() {
  const [emailInput, setEmailInput] = useState('');
  const [passwordInput, setPasswordInput] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const { setCurrentUser } = useAuth();
  const navigate = useNavigate();

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setError(null);
    
    try {
      const email = emailInput.trim().toLowerCase();
      const { data, error: queryError } = await supabase
        .from('users')
        .select('*')
        .ilike('email', email)
        .eq('password', passwordInput)
        .single();

      if (queryError || !data) {
        setError("Invalid login credentials");
      } else {
        // Persistent Session
        localStorage.setItem('omDedyUser', JSON.stringify(data));
        setCurrentUser(data);
        navigate('/portfolio');
      }
    } catch (err: any) {
      setError(err.message || "An unexpected error occurred");
    } finally {
      setLoading(false);
    }
  };

  const handleSignUp = async () => {
    setLoading(true);
    setError(null);
    const { error } = await supabase.auth.signUp({ email: emailInput, password: passwordInput });
    if (error) {
      if (error.status === 400 && error.message.includes('already registered')) {
        setError("Email ini sudah terdaftar. Silakan gunakan menu Login.");
      } else {
        setError(error.message);
      }
    } else {
      setError("Check your email for verification link.");
    }
    setLoading(false);
  };

  return (
    <div className="min-h-screen bg-slate-50 dark:bg-[#020617] flex items-center justify-center p-6 bg-[radial-gradient(circle_at_50%_50%,rgba(79,70,229,0.1),transparent)] transition-colors duration-300">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        className="w-full max-w-md bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl overflow-hidden shadow-xl dark:shadow-[0_0_50px_rgba(79,70,229,0.1)] transition-colors duration-300"
      >
        <div className="p-10 border-b border-slate-200 dark:border-slate-800 bg-indigo-50 dark:bg-indigo-600/5 text-center transition-colors">
          <div className="w-16 h-16 bg-white dark:bg-indigo-600/10 rounded-2xl flex items-center justify-center mx-auto mb-6 border border-slate-200 dark:border-indigo-500/20 shadow-sm transition-all">
            <Cpu className="w-8 h-8 text-indigo-600 dark:text-indigo-500" />
          </div>
          <h1 className="text-3xl font-black text-slate-900 dark:text-white tracking-tighter uppercase italic transition-colors">
            OM <span className="text-indigo-600 dark:text-indigo-500">DEDY</span> <span className="text-slate-400 dark:text-slate-500 italic lowercase text-sm">v2.0</span>
          </h1>
          <p className="text-[10px] text-slate-500 mt-2 font-bold uppercase tracking-[0.2em]">Operational Management & Delivery Deployment Yield</p>
        </div>

        <form onSubmit={handleLogin} className="p-8 space-y-6">
          {error && (
            <div className="bg-rose-500/10 border border-rose-500/20 p-3 rounded-xl flex items-center gap-3 text-rose-400 text-xs">
              <AlertCircle className="w-4 h-4 shrink-0" />
              <span>{error}</span>
            </div>
          )}

          <div className="space-y-4">
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Access Identity</label>
              <div className="relative">
                <Mail className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 dark:text-slate-600" />
                <input 
                  type="email"
                  required
                  value={emailInput}
                  onChange={(e) => setEmailInput(e.target.value)}
                  placeholder="name@company.com"
                  className="w-full bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl pl-12 pr-4 py-3 text-slate-900 dark:text-slate-200 outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-inner dark:shadow-none"
                />
              </div>
            </div>

            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Crypto-Key</label>
              <div className="relative">
                <Lock className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 dark:text-slate-600" />
                <input 
                  type="password"
                  required
                  value={passwordInput}
                  onChange={(e) => setPasswordInput(e.target.value)}
                  placeholder="••••••••"
                  className="w-full bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl pl-12 pr-4 py-3 text-slate-900 dark:text-slate-200 outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-inner dark:shadow-none"
                />
              </div>
            </div>
          </div>

          <div className="flex flex-col gap-3">
            <button 
              type="submit"
              disabled={loading}
              className="w-full py-4 bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 text-white font-black uppercase text-xs tracking-[0.2em] rounded-xl shadow-lg shadow-indigo-600/20 transition-all flex items-center justify-center gap-2 active:scale-98"
            >
              {loading ? "Authenticating..." : "Authorize Access"}
            </button>
            <button 
              type="button"
              onClick={handleSignUp}
              disabled={loading}
              className="w-full py-3 text-slate-500 hover:text-indigo-600 dark:hover:text-slate-300 font-bold uppercase text-[10px] tracking-widest transition-colors"
            >
              Provision New Account
            </button>
          </div>
        </form>

        <div className="p-8 bg-slate-50/50 dark:bg-slate-950/50 border-t border-slate-200 dark:border-slate-800/60 text-center transition-colors">
          <p className="text-[9px] text-slate-400 dark:text-slate-600 leading-relaxed uppercase tracking-widest">
            Authorized Personnel Only • Dynamic Governance Protocol v2.5.1
          </p>
        </div>
      </motion.div>
    </div>
  );
}

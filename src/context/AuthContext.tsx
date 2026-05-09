import React, { createContext, useContext, useEffect, useState } from 'react';
import { User, Session } from '@supabase/supabase-js';
import { supabase } from '../lib/supabase';
import { AppUser } from '../types';

interface AuthContextType {
  user: User | null;
  session: Session | null;
  currentUser: AppUser | null;
  loading: boolean;
  signOut: () => Promise<void>;
  setCurrentUser: (user: AppUser | null) => void;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export function AuthProvider({ children }: { children: React.ReactNode }) {
  const [user, setUser] = useState<User | null>(null);
  const [session, setSession] = useState<Session | null>(null);
  const [currentUser, setCurrentUser] = useState<AppUser | null>(null);
  const [loading, setLoading] = useState(true);

  const fetchCurrentUser = async (email: string) => {
    try {
      const { data, error } = await supabase
        .from('users')
        .select('*')
        .eq('email', email)
        .single();
      
      if (data && !error) {
        localStorage.setItem('omDedyUser', JSON.stringify(data));
        setCurrentUser(data);
      } else {
        // If not found in DB but logged into Supabase Auth, they might be an external viewer
        // OR a newly signed up user.
        setCurrentUser(null);
      }
    } catch (err) {
      console.error('Error fetching current user:', err);
    }
  };

  useEffect(() => {
    // Initial load from localStorage
    const savedUser = localStorage.getItem('omDedyUser');
    if (savedUser) {
      setCurrentUser(JSON.parse(savedUser));
    }

    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      setUser(session?.user ?? null);
      // Only fetch from DB if we don't have a manual session saved
      if (session?.user?.email && !savedUser) {
        fetchCurrentUser(session.user.email);
      }
      setLoading(false);
    });

    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
      setUser(session?.user ?? null);
      
      if (session?.user?.email) {
        fetchCurrentUser(session.user.email);
      } else {
        // Only clear if there's no manual session in localStorage
        const stillHasManualSession = localStorage.getItem('omDedyUser');
        if (!stillHasManualSession) {
          setCurrentUser(null);
        }
      }
      setLoading(false);
    });

    return () => subscription.unsubscribe();
  }, []);

  const signOut = async () => {
    localStorage.removeItem('omDedyUser');
    setCurrentUser(null);
    await supabase.auth.signOut();
  };

  const value = {
    user,
    session,
    currentUser,
    loading,
    signOut,
    setCurrentUser
  };

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
}

export function useAuth() {
  const context = useContext(AuthContext);
  if (context === undefined) {
    throw new Error('useAuth must be used within an AuthProvider');
  }
  return context;
}

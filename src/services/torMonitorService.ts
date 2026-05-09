import { Task, Project, TaskStatus, MasterProject, ProjectStatus } from '../types';
import { differenceInDays, isAfter, isBefore, startOfDay } from 'date-fns';
import { supabase } from '../lib/supabase';
import { taskService } from './taskService';

export interface TorStatus {
  label: string;
  color: 'emerald' | 'rose' | 'slate' | 'blue';
  days?: number;
}

export const getL2PrecisionStatus = (task: Task): TorStatus => {
  const today = startOfDay(new Date());
  const planStart = startOfDay(new Date(task.start_time));
  const planFinish = startOfDay(new Date(task.end_time));
  
  const isFinished = !!task.realized_finish_date;
  const realizedDate = task.realized_finish_date ? startOfDay(new Date(task.realized_finish_date)) : null;

  if (isFinished && realizedDate) {
    const diff = differenceInDays(realizedDate, planFinish);
    if (diff > 0) return { label: `Overdue ${diff} days`, color: 'rose', days: diff };
    if (diff < 0) return { label: `Early by ${Math.abs(diff)} days`, color: 'emerald', days: Math.abs(diff) };
    return { label: 'Done', color: 'emerald' };
  }

  // Not Finished (Realized Finish Date is empty)
  if (isBefore(planFinish, today)) {
    const diff = differenceInDays(today, planFinish);
    return { label: `Overdue ${diff} days`, color: 'rose', days: diff };
  }

  if (isAfter(planStart, today)) {
    return { label: 'To Do', color: 'slate' };
  }

  return { label: 'In Progress', color: 'blue' };
};

export const checkPhaseSequencing = (tasks: Task[], currentTask: Task): boolean => {
  const l1Phases = tasks.filter(t => !t.parent_id && t.project_id === currentTask.project_id);
  const fsdPhase = l1Phases.find(t => t.title.toUpperCase().includes('FSD'));
  const devPhase = l1Phases.find(t => t.title.toUpperCase().includes('DEV'));

  // A task is "Finished" if it has realized_finish_date
  const isFinished = (t: Task) => !!t.realized_finish_date;

  if (currentTask.title.toUpperCase().includes('DEV')) {
    // Development can only be In Progress if FSD is finished
    // Check all L2 tasks under FSD
    if (!fsdPhase) return true;
    const fsdL2 = tasks.filter(t => t.parent_id === fsdPhase.id);
    return fsdL2.length > 0 ? fsdL2.every(isFinished) : isFinished(fsdPhase);
  }

  if (currentTask.title.toUpperCase().includes('SIT') || currentTask.title.toUpperCase().includes('UAT')) {
    const fsdL2 = fsdPhase ? tasks.filter(t => t.parent_id === fsdPhase.id) : [];
    const devL2 = devPhase ? tasks.filter(t => t.parent_id === devPhase.id) : [];
    
    const fsdDone = fsdPhase ? (fsdL2.length > 0 ? fsdL2.every(isFinished) : isFinished(fsdPhase)) : true;
    const devDone = devPhase ? (devL2.length > 0 ? devL2.every(isFinished) : isFinished(devPhase)) : true;
    
    return fsdDone && devDone;
  }

  return true;
};

// Fungsi ini menerima daftar seluruh task (L1/L2) dari sebuah project
export function calculateGlobalProjectStatus(tasks: Task[]) {
  if (!tasks || tasks.length === 0) return "To Do";

  // 1. HELPER: Case-insensitive check for completed statuses
  const isCompleted = (task: Task) => {
    const status = (task.current_status || task.status || '').toUpperCase();
    return !!task.realized_finish_date || status.includes('DONE') || status.includes('OVERDUE') || status.includes('EARLY');
  };

  // 2. HELPER: Case-insensitive phase matching
  const getTasksByPhase = (phaseName: string) => {
    return tasks.filter(t => (t.phase_type || t.title || '').toUpperCase().includes(phaseName.toUpperCase()));
  };

  const fsdTasks = getTasksByPhase('FSD');
  const devTasks = getTasksByPhase('DEV');
  const sitTasks = getTasksByPhase('SIT');
  const uatTasks = getTasksByPhase('UAT');

  // --- DEBUG LOGGER ---
  console.log("--- STATUS CHECK ---");
  console.log("FSD Completed?", fsdTasks.length > 0 ? fsdTasks.every(isCompleted) : "No Tasks");
  console.log("DEV Completed?", devTasks.length > 0 ? devTasks.every(isCompleted) : "No Tasks");

  // 3. SEQUENTIAL EVALUATION LOGIC
  if (fsdTasks.length > 0 && !fsdTasks.every(isCompleted)) return "FSD On Progress";
  if (devTasks.length > 0 && !devTasks.every(isCompleted)) return "Development On Progress";
  if (sitTasks.length > 0 && !sitTasks.every(isCompleted)) return "SIT On Progress";
  if (uatTasks.length > 0 && !uatTasks.every(isCompleted)) return "UAT On Progress";
  if (tasks.length > 0 && tasks.every(isCompleted)) return "LIVE";

  return "To Do";
}

export const torAutoSync = async (projectId: string, tasks: Task[], projects: Project[], actor: string) => {
  const project = projects.find(p => p.id === projectId);
  if (!project) return;

  // Retrieve auto-status config
  const savedConfig = localStorage.getItem('od_auto_status_configs');
  const autoStatusConfigs = savedConfig ? JSON.parse(savedConfig) : {};
  const isAuto = autoStatusConfigs[projectId] ?? true;

  if (!isAuto) return;

  const projectTasks = tasks.filter(t => t.project_id === projectId);
  if (projectTasks.length === 0) return;

  const l1Tasks = projectTasks.filter(t => !t.parent_id);
  
  const isTaskFinished = (t: Task) => !!t.realized_finish_date;
  const today = startOfDay(new Date());

  // Task Status Correction (L2 and L1)
  for (const t of projectTasks) {
    if (!isTaskFinished(t) && !t.is_manual_override) {
      const planStart = startOfDay(new Date(t.start_time));
      let expectedStatus = t.status;
      
      // Sequencing Check: If a phase depends on previous one
      const canStart = checkPhaseSequencing(projectTasks, t);
      
      if (!canStart) {
        expectedStatus = TaskStatus.TODO;
      } else if (isAfter(planStart, today)) {
        expectedStatus = TaskStatus.TODO;
      } else {
        expectedStatus = TaskStatus.IN_PROGRESS;
      }

      // Automatically shift from To Do to In Progress or vice versa if realized is empty
      // and it's currently one of the "active" statuses (not hold/cancel)
      if (t.status === TaskStatus.TODO || t.status === TaskStatus.IN_PROGRESS) {
        if (t.status !== expectedStatus) {
          const oldStatus = t.status;
          await supabase.from('tasks').update({ status: expectedStatus, updated_at: new Date().toISOString() }).eq('id', t.id);
          t.status = expectedStatus; // Update local reference for subsequent logic
          
          await taskService.logAudit({
            project_id: projectId,
            task_id: t.id,
            actor: 'System Auto-Status',
            action: 'Status Correction',
            oldValue: { status: oldStatus },
            newValue: { status: expectedStatus, note: 'Plan date reached or sequencing requirement met' }
          });
        }
      }
    }
  }

  // Level Phase (L1) Logic: DONE only if all L2 under it are Finished
  const newL1Updates: { id: string, realized_finish_date: string | null, status: TaskStatus }[] = [];

  for (const l1 of l1Tasks) {
    const l2Tasks = projectTasks.filter(t => t.parent_id === l1.id);
    if (l2Tasks.length > 0) {
      const allL2Finished = l2Tasks.every(isTaskFinished);
      if (allL2Finished && !l1.realized_finish_date) {
        // Find latest realized finish date from L2 to set as L1's realized finish date
        const latestRealized = l2Tasks.reduce((latest, current) => {
          if (!latest) return current.realized_finish_date;
          if (!current.realized_finish_date) return latest;
          return new Date(current.realized_finish_date) > new Date(latest) ? current.realized_finish_date : latest;
        }, null as string | null);

        newL1Updates.push({ 
          id: l1.id, 
          realized_finish_date: latestRealized || new Date().toISOString(),
          status: TaskStatus.DONE 
        });
      }
    }
  }

  // Apply L1 updates
  for (const update of newL1Updates) {
    await supabase.from('tasks').update({ 
      realized_finish_date: update.realized_finish_date,
      status: update.status,
      updated_at: new Date().toISOString()
    }).eq('id', update.id);

    await taskService.logAudit({
      project_id: projectId,
      task_id: update.id,
      actor: 'System Auto-Update',
      action: 'Phase Finished',
      oldValue: { realized: null },
      newValue: { realized: update.realized_finish_date, note: `All L2 tasks finished` }
    });
  }

  // Calculate new Global Status using common logic
  const newGlobalStatus = calculateGlobalProjectStatus(projectTasks);

  if (newGlobalStatus && project.status !== newGlobalStatus) {
    // 1. Update projects table
    await supabase.from('projects').update({ 
      status: newGlobalStatus as ProjectStatus,
      updated_at: new Date().toISOString()
    }).eq('id', projectId);

    // 2. UPSERT to master_projects using ticket_id
    if (project.ticket_id) {
       const masterPayload: Partial<MasterProject> = {
          ticket_id: project.ticket_id,
          project_name: project.name,
          status: newGlobalStatus,
          global_status: newGlobalStatus, // Sync both fields as per request
          pic_name: project.pic_name || 'Unassigned',
          owner_name: project.owner_name || '',
          div_owner: project.div_owner || '',
          updated_at: new Date().toISOString()
       };

       const { data: existingMaster } = await supabase.from('master_projects').select('id').eq('ticket_id', project.ticket_id).maybeSingle();
       if (existingMaster) {
          await taskService.updateMasterProject(existingMaster.id, masterPayload, 'System Sync');
       } else {
          await supabase.from('master_projects').insert([{ ...masterPayload, created_at: new Date().toISOString() }]);
       }
    }

    await taskService.logAudit({
      project_id: projectId,
      actor: 'System Auto-Update',
      action: 'Global Status Shift',
      oldValue: { status: project.status },
      newValue: { status: newGlobalStatus, note: 'Sequential shift triggered' }
    });
    return newGlobalStatus;
  }
  return null;
};


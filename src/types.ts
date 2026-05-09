export enum TaskStatus {
  TODO = 'To Do',
  IN_PROGRESS = 'In Progress',
  ON_REVIEW = 'On Review',
  DONE = 'Done',
  OVERDUE = 'Overdue',
  EARLY = 'Early',
  HOLD = 'Hold',
  CANCEL = 'Cancel'
}

export type Status = TaskStatus; 

export enum ProjectStatus {
  TODO = 'To Do',
  IN_PROGRESS = 'In Progress',
  FSD_PROGRESS = 'FSD On Progress',
  DEV_PROGRESS = 'Development On Progress',
  SIT_PROGRESS = 'SIT On Progress',
  UAT_PROGRESS = 'UAT On Progress',
  LIVE = 'LIVE',
  HOLD = 'Hold',
  CANCEL = 'Cancel',
  LATE = 'Project Late'
}

export interface Project {
  id: string;
  ticket_id?: string;
  name: string;
  status: ProjectStatus;
  leader_email?: string;
  pic_name?: string; 
  owner_name?: string;
  div_owner?: string;
  start_date?: string;
  end_date?: string;
  created_at: string;
  updated_at: string;
}

export interface AppUser {
  id: string;
  name: string;
  email: string;
  password?: string;
  access_level: string; // Dropdown: Superadmin/Admin/PIC/etc.
  role: string;         // Input Text: Developer/QA/etc.
  created_at: string;
  updated_at?: string;
}

export interface Task {
  id: string;
  parent_id: string | null;
  project_id: string | null;
  title: string;
  assignee: string; // Used as display name (free text)
  developer_name?: string;
  qa_name?: string;
  start_time: string; // ISO string
  end_time: string;   // ISO string
  start_hour: number;  // 0-23
  start_minute: number; // 0-59
  duration_hours: number;
  duration_minutes: number;
  man_hours?: number; // Manual man-hours input
  target_sla_date: string | null;
  created_by_name?: string;
  status: TaskStatus;
  custom_id: string; // Visible ID like #PH-XXXX or #TS-XXXX
  approval_fachrul: string | null;
  suggestion_fachrul: string | null;
  approval_barra: string | null;
  suggestion_barra: string | null;
  created_at: string;
  updated_at: string;
  level?: number;
  realized_finish_date?: string | null;
  is_manual_override?: boolean;
  current_status?: string;
  phase_type?: string;
}

export interface AuditLog {
  id: string;
  task_id?: string;
  project_id?: string;
  user_id?: string;
  actor: string;
  action: string;
  old_payload: any;
  new_payload: any;
  created_at: string;
  updated_at?: string;
}

export type ViewScale = 'MONTH' | 'WEEK' | 'DAY' | 'HOUR';
export type AppView = 'PROJECTS' | 'KANBAN' | 'PERSONEL' | 'AUDIT' | 'GANTT_DETAIL' | 'SCHEDULE' | 'RESCHEDULE' | 'LOGIN' | 'MASTER_PROJECT' | 'TOR_MONITOR' | 'TIMELINE';

export interface Schedule {
  id: string;
  pic_name: string;
  schedule_date: string; // ISO date string
  status: string;
  created_at: string;
  updated_at: string;
}

export interface RescheduleRequest {
  id: string;
  pic_name: string;
  schedule_date: string;
  original_status: string;
  new_status: string;
  reason: string;
  requested_by: string; // email or name
  status: 'Pending' | 'Approved' | 'Rejected';
  swap_date?: string;
  swap_status?: string;
  created_at: string;
  updated_at: string;
}

export interface ProjectRescheduleLog {
  id: string;
  project_id: string;
  changed_by: string;
  old_start_date: string;
  old_end_date: string;
  new_start_date: string;
  new_end_date: string;
  reason: string;
  created_at: string;
}

export interface MasterProject {
  id: string;
  ticket_id: string;
  project_name: string;
  status: string; // Used for "Global Status" (e.g. OPEN, LIVE, etc)
  global_status?: string; // Explicit field as per request
  pic_name: string;
  owner_name: string;
  div_owner: string;
  created_at: string;
  updated_at: string;
}

export interface MasterProjectAuditLog {
  id: string;
  master_project_id: string;
  actor: string;
  action: 'IMPORT' | 'CREATE' | 'UPDATE' | 'DELETE';
  note: string;
  changed_fields?: string;
  created_at: string;
}

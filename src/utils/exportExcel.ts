import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { format } from 'date-fns';
import { PHASE_ORDER } from '../constants';

/**
 * Higher-order export utility for Om Dedy Dashboard
 * Generates a management-ready styled XLSX report
 */
export const exportToExcel = async (project: any, tasks: any[]) => {
  if (!project || !tasks) {
    console.error("Export failed: Missing project or tasks data");
    return;
  }

  // 1. Sort tasks chronologically by PHASE_ORDER
  const sortedTasks = [...tasks].sort((a, b) => {
    const phaseA = (a.phase_type || a.title || a.name || '').toUpperCase();
    const phaseB = (b.phase_type || b.title || b.name || '').toUpperCase();
    
    // Check if the title starts with one of the phases
    const getOrder = (p: string) => {
       for (const key in PHASE_ORDER) {
          if (p.includes(key)) return PHASE_ORDER[key];
       }
       return 99;
    };

    const orderA = getOrder(phaseA);
    const orderB = getOrder(phaseB);
    
    if (orderA !== orderB) return orderA - orderB;
    
    // Secondary sort by date if same phase or unknown
    return new Date(a.start_time || a.plan_start_date).getTime() - new Date(b.start_time || b.plan_start_date).getTime();
  });

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Timeline Approval');

  // 1. SET COLUMN WIDTHS
  sheet.columns = [
    { header: 'Phase / Hierarchy', key: 'phase', width: 35 },
    { header: 'PIC', key: 'pic', width: 25 },
    { header: 'Plan Start', key: 'planStart', width: 18 },
    { header: 'Plan End', key: 'planEnd', width: 18 },
    { header: 'Realized Finish', key: 'realized', width: 18 },
    { header: 'Man-Hours', key: 'manHours', width: 15 },
    { header: 'Man Hour ( Minutes )', key: 'manHoursMin', width: 20 },
    { header: 'Status', key: 'status', width: 25 },
    { header: 'Fachrul Feedback', key: 'feedback1', width: 30 },
    { header: 'Barra Feedback', key: 'feedback2', width: 30 },
  ];

  // 2. ADD TITLE
  sheet.mergeCells('A1:J1');
  const titleCell = sheet.getCell('A1');
  titleCell.value = 'OM DEDY - PROJECT TIMELINE APPROVAL';
  titleCell.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FF1E3A8A' } };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

  sheet.addRow([]); // Spacer

  // 3. ADD SUMMARY INFO
  const addSummaryRow = (label1: string, val1: string, label2: string, val2: string) => {
    const row = sheet.addRow([label1, val1, '', label2, val2]);
    row.getCell(1).font = { bold: true };
    row.getCell(4).font = { bold: true };
    sheet.mergeCells(`B${row.number}:C${row.number}`);
    sheet.mergeCells(`E${row.number}:J${row.number}`);
    return row;
  };

  // Helper to format date safely
  const formatDate = (date: any) => {
    if (!date) return '-';
    try {
      return format(new Date(date), 'MMM dd, yyyy');
    } catch {
      return date.toString();
    }
  };

  // Calculate total man hours dynamically
  const totalHours = tasks.reduce((sum, task) => sum + (parseFloat(task.man_hours) || 0), 0);

  addSummaryRow(
    'Project Name', 
    project.project_name || project.name || 'UNTITLED PROJECT', 
    'Global Status', 
    project.global_status || project.status || 'TO DO'
  );
  addSummaryRow(
    'Start Date', 
    formatDate(project.plan_start_date || project.start_date), 
    'Total Man Hours', 
    `${totalHours.toFixed(1)} Hours`
  );
  addSummaryRow(
    'End Date', 
    formatDate(project.plan_end_date || project.end_date), 
    'PIC', 
    project.pic_name || project.leader_email || 'Unassigned'
  );
  
  sheet.addRow([]); // Spacer before table

  // 4. ADD TABLE HEADERS
  const headerRow = sheet.addRow([
    'Phase / Hierarchy', 'PIC', 'Plan Start', 'Plan End', 'Realized Finish', 'Man-Hours', 'Man Hour ( Minutes )', 'Status', 'Fachrul Feedback', 'Barra Feedback'
  ]);
  
  headerRow.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } }; // Dark Blue
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = { top: { style:'thin' }, left: { style:'thin' }, bottom: { style:'thin' }, right: { style:'thin' } };
  });

  // 5. ADD DATA ROWS WITH CONDITIONAL STATUS COLORS
  sortedTasks.forEach(task => {
    const status = (task.current_status || task.status || 'TO DO').toUpperCase();
    const hours = parseFloat(task.man_hours) || 0;
    const minutes = hours * 60;

    const row = sheet.addRow([
      task.phase_type || task.title || task.name || '-',
      task.pic_name || task.assignee || task.developer_name || '-',
      formatDate(task.plan_start_date || task.start_time),
      formatDate(task.plan_end_date || task.end_time),
      formatDate(task.realized_finish_date),
      `${hours} h`,
      `${minutes} m`,
      status,
      task.fachrul_feedback || task.suggestion_fachrul || '-',
      task.barra_feedback || task.suggestion_barra || '-'
    ]);

    row.eachCell((cell, colNumber) => {
      cell.border = { top: { style:'thin' }, left: { style:'thin' }, bottom: { style:'thin' }, right: { style:'thin' } };
      if (colNumber > 2) cell.alignment = { horizontal: 'center', vertical: 'middle' };
      if (colNumber === 1 || colNumber === 2) cell.alignment = { vertical: 'middle' };
      
      // Status Coloring Logic
      if (colNumber === 8) {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        const stat = status;
        if (stat.includes('DONE') || stat.includes('EARLY') || stat.includes('LIVE')) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF10B981' } }; // Emerald Green
        } else if (stat.includes('PROGRESS') || stat.includes('REVIEW')) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3B82F6' } }; // Blue
        } else if (stat.includes('LATE') || stat.includes('OVERDUE')) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEF4444' } }; // Red
        } else {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF6B7280' } }; // Gray (To Do/Hold/Cancel)
        }
      }
    });
  });

  // 6. GENERATE & DOWNLOAD
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  
  // Format filename: OM_DEDY_Timeline_PROJECT_NAME.xlsx
  const projectName = project.project_name || project.name || 'Export';
  const safeProjectName = projectName.toString().replace(/[^a-zA-Z0-9]/g, '_');
  saveAs(blob, `OM_DEDY_Timeline_${safeProjectName}.xlsx`);
};

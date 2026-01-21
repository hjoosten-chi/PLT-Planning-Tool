/**
 * PLT Planning Tool - Google Apps Script Backend
 * Manages project/activity data for People Leadership Team planning
 */

// Configuration
const CONFIG = {
  SHEET_NAME: 'Sample Template - Source Data',
  SETTINGS_SHEET: 'Settings',
  HEADER_ROW: 1,
  DATA_START_ROW: 2
};

/**
 * Serves the main HTML page
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('PLT Planning Tool')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Include external HTML/CSS/JS files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get all projects from the spreadsheet
 */
function getProjects() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

    if (!sheet) {
      return { error: 'Sheet not found: ' + CONFIG.SHEET_NAME };
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < CONFIG.DATA_START_ROW) {
      return { data: [], headers: [] };
    }

    // Get headers
    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0];

    // Get data
    const dataRange = sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.HEADER_ROW, lastCol);
    const data = dataRange.getValues();

    // Convert to array of objects
    const projects = data.map((row, index) => {
      const project = { rowIndex: index + CONFIG.DATA_START_ROW };
      headers.forEach((header, colIndex) => {
        let value = row[colIndex];
        // Format dates
        if (value instanceof Date) {
          value = formatDate(value);
        }
        project[normalizeHeader(header)] = value;
      });
      return project;
    }).filter(project => project.projectActivityName); // Filter empty rows

    return {
      data: projects,
      headers: headers.map(h => ({ original: h, normalized: normalizeHeader(h) }))
    };
  } catch (error) {
    return { error: error.toString() };
  }
}

/**
 * Get unique values for filter dropdowns
 */
function getFilterOptions() {
  try {
    const result = getProjects();
    if (result.error) return result;

    const projects = result.data;

    const options = {
      categories: [...new Set(projects.map(p => p.category).filter(Boolean))].sort(),
      statuses: [...new Set(projects.map(p => p.status).filter(Boolean))].sort(),
      functionalOwners: [...new Set(projects.map(p => p.functionalOwnerOfDeliverable).filter(Boolean))].sort(),
      programOwners: [...new Set(projects.map(p => p.programOwnerLeadContact).filter(Boolean))].sort(),
      efforts: [...new Set(projects.map(p => p.effort).filter(Boolean))].sort(),
      pltHelpNeeded: ['Yes', 'No']
    };

    return options;
  } catch (error) {
    return { error: error.toString() };
  }
}

/**
 * Update a project's status
 */
function updateProjectStatus(rowIndex, newStatus) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

    // Find the Status column (column E = 5)
    const statusCol = 5;
    sheet.getRange(rowIndex, statusCol).setValue(newStatus);

    return { success: true };
  } catch (error) {
    return { error: error.toString() };
  }
}

/**
 * Update a cell value
 */
function updateCell(rowIndex, columnName, value) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

    // Get headers to find column index
    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = headers.findIndex(h => normalizeHeader(h) === columnName) + 1;

    if (colIndex === 0) {
      return { error: 'Column not found: ' + columnName };
    }

    sheet.getRange(rowIndex, colIndex).setValue(value);
    return { success: true };
  } catch (error) {
    return { error: error.toString() };
  }
}

/**
 * Add a new project
 */
function addProject(projectData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

    // Get headers
    const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Create row data
    const rowData = headers.map(header => {
      const key = normalizeHeader(header);
      return projectData[key] || '';
    });

    // Append row
    sheet.appendRow(rowData);

    return { success: true, rowIndex: sheet.getLastRow() };
  } catch (error) {
    return { error: error.toString() };
  }
}

/**
 * Get summary statistics
 */
function getSummaryStats() {
  try {
    const result = getProjects();
    if (result.error) return result;

    const projects = result.data;

    const stats = {
      total: projects.length,
      byStatus: {},
      byCategory: {},
      byEffort: {},
      pltHelpNeeded: 0,
      completedThisMonth: 0,
      atRisk: 0
    };

    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();

    projects.forEach(project => {
      // By status
      const status = project.status || 'Unknown';
      stats.byStatus[status] = (stats.byStatus[status] || 0) + 1;

      // By category
      const category = project.category || 'Unknown';
      stats.byCategory[category] = (stats.byCategory[category] || 0) + 1;

      // By effort
      const effort = project.effort || 'Unknown';
      stats.byEffort[effort] = (stats.byEffort[effort] || 0) + 1;

      // PLT Help Needed
      if (project.pltHelpNeeded === 'Yes') {
        stats.pltHelpNeeded++;
      }

      // At Risk
      if (status === 'At Risk') {
        stats.atRisk++;
      }
    });

    return stats;
  } catch (error) {
    return { error: error.toString() };
  }
}

/**
 * Get projects for a specific month view
 */
function getProjectsByMonth(month, year) {
  try {
    const result = getProjects();
    if (result.error) return result;

    const projects = result.data.filter(project => {
      if (!project.startDate && !project.endDate) return false;

      const startDate = project.startDate ? parseDate(project.startDate) : null;
      const endDate = project.endDate ? parseDate(project.endDate) : null;

      const monthStart = new Date(year, month - 1, 1);
      const monthEnd = new Date(year, month, 0);

      // Project is in this month if it overlaps with the month
      if (startDate && endDate) {
        return startDate <= monthEnd && endDate >= monthStart;
      } else if (startDate) {
        return startDate.getMonth() + 1 === month && startDate.getFullYear() === year;
      } else if (endDate) {
        return endDate.getMonth() + 1 === month && endDate.getFullYear() === year;
      }

      return false;
    });

    return { data: projects };
  } catch (error) {
    return { error: error.toString() };
  }
}

// Utility Functions

/**
 * Normalize header names to camelCase keys
 */
function normalizeHeader(header) {
  if (!header) return '';
  return header
    .toString()
    .replace(/[^a-zA-Z0-9\s]/g, '')
    .split(/\s+/)
    .map((word, index) => {
      if (index === 0) {
        return word.toLowerCase();
      }
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join('');
}

/**
 * Format date for display
 */
function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;

  const month = date.getMonth() + 1;
  const day = date.getDate();
  const year = date.getFullYear();

  return `${month}/${day}/${year}`;
}

/**
 * Parse date string to Date object
 */
function parseDate(dateStr) {
  if (!dateStr) return null;
  if (dateStr instanceof Date) return dateStr;

  const parts = dateStr.split('/');
  if (parts.length === 3) {
    return new Date(parts[2], parts[0] - 1, parts[1]);
  }
  return new Date(dateStr);
}

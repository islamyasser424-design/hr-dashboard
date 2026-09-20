/**
 * ═══════════════════════════════════════════════════════════════════════════
 * HR DASHBOARD - GOOGLE APPS SCRIPT BACKEND
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * This file handles all backend logic for the HR Analytics Dashboard:
 * - Data persistence and retrieval from Google Sheets
 * - KPI calculations (Headcount, Attrition, Turnover Cost, etc.)
 * - HTML template serving and initialization
 * 
 * Author: HR Analytics Team
 * Last Updated: 2024
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─────────────────────────────────────────────────────────────────────────
// CONFIGURATION & CONSTANTS
// ─────────────────────────────────────────────────────────────────────────

/** @const {string} SHEET_NAME - Name of the active Google Sheet tab */
const SHEET_NAME = 'HR_Data';

/** @const {number} AVERAGE_COST_PER_HIRE - Cost multiplier for turnover calculation */
const AVERAGE_COST_PER_HIRE = 0.5; // 50% of annual salary

/** @const {Array<string>} HEADER_ROW - Column headers for HR database */
const HEADER_ROW = [
  'EmpID',
  'EmployeeName',
  'Department',
  'JobRole',
  'Age',
  'Tenure_Years',
  'JobLevel',
  'Performance',
  'OverTime',
  'Travel',
  'Attrition',
  'Salary',
  'StartDate'
];

// ─────────────────────────────────────────────────────────────────────────
// MAIN ENTRY POINT - HTTP Handler
// ─────────────────────────────────────────────────────────────────────────

/**
 * Main entry point for the web app. Serves the HTML template with embedded 
 * data and configuration.
 * 
 * @param {object} e - Apps Script request object
 * @returns {HtmlOutput} Rendered HTML template
 */
function doGet(e) {
  try {
    // Serve the index.html template
    const template = HtmlService.createTemplateFromFile('index');
    
    // Pass initial dashboard data to frontend
    template.dashboardData = JSON.stringify(getDashboardData());
    template.themes = JSON.stringify(getThemeConfig());
    
    return template
      .evaluate()
      .setWidth(1200)
      .setHeight(800)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    Logger.log('Error in doGet: ' + error);
    return HtmlService.createHtmlOutput('<h1>Error: ' + error + '</h1>');
  }
}

// ─────────────────────────────────────────────────────────────────────────
// DATA INITIALIZATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * Initializes the active Google Sheet with sample HR data (50 employees).
 * This creates a realistic dataset for dashboard demonstration.
 * 
 * Call this function once via the Apps Script editor to populate data.
 * Run > setupDummyData
 */
function setupDummyData() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.setName(SHEET_NAME);
    
    // Clear existing data
    sheet.clear();
    
    // Insert headers
    sheet.getRange(1, 1, 1, HEADER_ROW.length).setValues([HEADER_ROW]);
    
    // Generate 50 realistic employee records
    const sampleData = generateSampleEmployees(50);
    
    // Write data to sheet (starting from row 2)
    if (sampleData.length > 0) {
      sheet.getRange(2, 1, sampleData.length, HEADER_ROW.length).setValues(sampleData);
    }
    
    // Format header row (bold + background)
    const headerRange = sheet.getRange(1, 1, 1, HEADER_ROW.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1f2937');
    headerRange.setFontColor('#ffffff');
    
    // Auto-fit columns
    sheet.autoResizeColumns(1, HEADER_ROW.length);
    
    SpreadsheetApp.getUi().alert('Sample data initialized with 50 employees!');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error);
    Logger.log('setupDummyData error: ' + error);
  }
}

/**
 * Generates an array of realistic sample employee records.
 * 
 * @param {number} count - Number of employees to generate
 * @returns {Array<Array>} Array of employee data rows
 */
function generateSampleEmployees(count) {
  const departments = ['Sales', 'Engineering', 'HR', 'Finance', 'Marketing', 'Operations'];
  const jobRoles = ['Manager', 'Analyst', 'Engineer', 'Coordinator', 'Specialist', 'Director'];
  const employees = [];
  
  for (let i = 1; i <= count; i++) {
    const dept = departments[Math.floor(Math.random() * departments.length)];
    const role = jobRoles[Math.floor(Math.random() * jobRoles.length)];
    const age = Math.floor(Math.random() * (65 - 22 + 1)) + 22;
    const tenure = Math.floor(Math.random() * 15);
    const salary = 40000 + Math.random() * 80000;
    const performance = Math.floor(Math.random() * 5) + 1;
    
    // Higher performance = lower attrition probability
    const attritionProbability = performance >= 4 ? 0.05 : (performance === 3 ? 0.15 : 0.25);
    const isAttritted = Math.random() < attritionProbability ? 'Yes' : 'No';
    
    const startDate = new Date();
    startDate.setFullYear(startDate.getFullYear() - tenure);
    
    employees.push([
      'EMP' + String(i).padStart(4, '0'),                    // EmpID
      'Employee ' + i,                                       // EmployeeName
      dept,                                                  // Department
      role,                                                  // JobRole
      age,                                                   // Age
      tenure,                                                // Tenure_Years
      Math.floor(Math.random() * 5) + 1,                    // JobLevel
      performance,                                           // Performance
      Math.random() > 0.7 ? 'Yes' : 'No',                   // OverTime
      Math.random() > 0.6 ? 'Yes' : 'No',                   // Travel
      isAttritted,                                           // Attrition
      Math.round(salary),                                    // Salary
      Utilities.formatDate(startDate, 'GMT', 'yyyy-MM-dd')  // StartDate
    ]);
  }
  
  return employees;
}

// ─────────────────────────────────────────────────────────────────────────
// DATA RETRIEVAL & PROCESSING
// ─────────────────────────────────────────────────────────────────────────

/**
 * Main data fetching function. Reads the entire HR dataset from the active 
 * Google Sheet and returns structured KPI data.
 * 
 * This function is called via google.script.run from the frontend.
 * 
 * @returns {object} Structured dashboard data including employees and KPIs
 */
function getDashboardData() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      // No data beyond headers
      return {
        employees: [],
        kpis: getEmptyKPIs(),
        departments: [],
        jobRoles: []
      };
    }
    
    // Convert sheet data to array of objects
    const headers = data[0];
    const employees = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const employee = {};
      
      headers.forEach((header, index) => {
        employee[header] = row[index];
      });
      
      employees.push(employee);
    }
    
    // Calculate all KPIs
    const kpis = calculateKPIs(employees);
    
    // Extract unique departments and job roles
    const departments = [...new Set(employees.map(e => e.Department))].filter(Boolean);
    const jobRoles = [...new Set(employees.map(e => e.JobRole))].filter(Boolean);
    
    return {
      employees: employees,
      kpis: kpis,
      departments: departments,
      jobRoles: jobRoles,
      lastUpdated: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log('Error in getDashboardData: ' + error);
    return {
      employees: [],
      kpis: getEmptyKPIs(),
      departments: [],
      jobRoles: [],
      error: error.toString()
    };
  }
}

/**
 * Calculates all dashboard KPIs from the employee dataset.
 * This is the core analytics engine - translates Power BI DAX logic to JavaScript.
 * 
 * KPIs Calculated:
 * - Headcount: Total number of active employees
 * - Attrition Rate: % of employees who left
 * - Turnover Cost: Financial impact of attrition
 * - Regrettable Attrition: Loss of high performers
 * - OT + Travel Risk: Burnout risk indicator
 * 
 * @param {Array<object>} employees - Array of employee objects
 * @returns {object} Object containing all calculated KPIs
 */
function calculateKPIs(employees) {
  if (!employees || employees.length === 0) {
    return getEmptyKPIs();
  }
  
  // Total headcount (all employees including those who left)
  const totalEmployees = employees.length;
  
  // Active employees (still employed)
  const activeEmployees = employees.filter(e => e.Attrition !== 'Yes').length;
  
  // Attritted employees
  const attritedEmployees = employees.filter(e => e.Attrition === 'Yes').length;
  
  // Attrition rate calculation
  const attritionRate = totalEmployees > 0 
    ? parseFloat(((attritedEmployees / totalEmployees) * 100).toFixed(2))
    : 0;
  
  // Average salary calculation
  const totalSalary = employees.reduce((sum, e) => sum + (parseFloat(e.Salary) || 0), 0);
  const averageSalary = totalEmployees > 0 ? totalSalary / totalEmployees : 0;
  
  // Turnover cost calculation (attrite employees × avg salary × cost multiplier)
  const turnoverCost = Math.round(attritedEmployees * averageSalary * AVERAGE_COST_PER_HIRE);
  
  // Regrettable attrition (high performers who left)
  // Performance >= 4 is considered "high performer"
  const regrettableAttrition = employees.filter(
    e => e.Attrition === 'Yes' && parseFloat(e.Performance) >= 4
  ).length;
  
  const regrettableAttritionRate = totalEmployees > 0
    ? parseFloat(((regrettableAttrition / totalEmployees) * 100).toFixed(2))
    : 0;
  
  // OT + Travel Risk: Employees working overtime AND traveling
  const otTravelRisk = employees.filter(
    e => e.OverTime === 'Yes' && e.Travel === 'Yes'
  ).length;
  
  const otTravelRiskRate = totalEmployees > 0
    ? parseFloat(((otTravelRisk / totalEmployees) * 100).toFixed(2))
    : 0;
  
  // Attrition by department (for pie chart)
  const attritionByDept = {};
  employees.forEach(e => {
    if (!attritionByDept[e.Department]) {
      attritionByDept[e.Department] = { total: 0, attrite: 0 };
    }
    attritionByDept[e.Department].total++;
    if (e.Attrition === 'Yes') {
      attritionByDept[e.Department].attrite++;
    }
  });
  
  // Convert to array for charts
  const deptAttritionData = Object.entries(attritionByDept).map(([dept, data]) => ({
    department: dept,
    attritionRate: data.total > 0 ? parseFloat(((data.attrite / data.total) * 100).toFixed(1)) : 0,
    totalEmployees: data.total,
    attriteCount: data.attrite
  }));
  
  // Age distribution (for demographics)
  const ageRanges = {
    '20-30': 0,
    '31-40': 0,
    '41-50': 0,
    '51-60': 0,
    '60+': 0
  };
  
  employees.forEach(e => {
    const age = parseFloat(e.Age);
    if (age >= 20 && age <= 30) ageRanges['20-30']++;
    else if (age >= 31 && age <= 40) ageRanges['31-40']++;
    else if (age >= 41 && age <= 50) ageRanges['41-50']++;
    else if (age >= 51 && age <= 60) ageRanges['51-60']++;
    else if (age > 60) ageRanges['60+']++;
  });
  
  // Tenure distribution
  const tenureRanges = {
    '0-2': 0,
    '3-5': 0,
    '6-10': 0,
    '10+': 0
  };
  
  employees.forEach(e => {
    const tenure = parseFloat(e.Tenure_Years);
    if (tenure >= 0 && tenure <= 2) tenureRanges['0-2']++;
    else if (tenure >= 3 && tenure <= 5) tenureRanges['3-5']++;
    else if (tenure >= 6 && tenure <= 10) tenureRanges['6-10']++;
    else if (tenure > 10) tenureRanges['10+']++;
  });
  
  return {
    // Primary KPIs
    headcount: activeEmployees,
    totalEmployees: totalEmployees,
    attritionRate: attritionRate,
    attritedEmployees: attritedEmployees,
    turnoverCost: turnoverCost,
    averageSalary: Math.round(averageSalary),
    
    // Secondary KPIs
    regrettableAttrition: regrettableAttrition,
    regrettableAttritionRate: regrettableAttritionRate,
    otTravelRisk: otTravelRisk,
    otTravelRiskRate: otTravelRiskRate,
    
    // Aggregated data for charts
    departmentAttrition: deptAttritionData,
    ageDistribution: ageRanges,
    tenureDistribution: tenureRanges
  };
}

/**
 * Returns empty KPI structure for when no data is available.
 * Used as fallback in error scenarios.
 * 
 * @returns {object} KPI object with all values set to 0
 */
function getEmptyKPIs() {
  return {
    headcount: 0,
    totalEmployees: 0,
    attritionRate: 0,
    attritedEmployees: 0,
    turnoverCost: 0,
    averageSalary: 0,
    regrettableAttrition: 0,
    regrettableAttritionRate: 0,
    otTravelRisk: 0,
    otTravelRiskRate: 0,
    departmentAttrition: [],
    ageDistribution: {},
    tenureDistribution: {}
  };
}

// ─────────────────────────────────────────────────────────────────────────
// THEME & CONFIGURATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * Returns theme configuration for light and dark modes.
 * Color values are passed to frontend for CSS variable injection.
 * 
 * @returns {object} Theme configuration object
 */
function getThemeConfig() {
  return {
    light: {
      background: '#ffffff',
      surface: '#f9fafb',
      primary: '#3b82f6',
      secondary: '#8b5cf6',
      text: '#1f2937',
      textSecondary: '#6b7280',
      border: '#e5e7eb'
    },
    dark: {
      background: '#111827',
      surface: '#1f2937',
      primary: '#60a5fa',
      secondary: '#a78bfa',
      text: '#f3f4f6',
      textSecondary: '#d1d5db',
      border: '#374151'
    }
  };
}

/**
 * Hybrid Solar + Wind SaaS Platform (Google Apps Script)
 * -------------------------------------------------------
 * Production-oriented architecture for:
 * 1) Calculation engine (solar + wind + finance)
 * 2) Monthly projection engine
 * 3) PDF proposal generation via HTML template
 * 4) Email automation
 * 5) CRM persistence workflow
 * 6) Utilities, logging, and configuration management
 */

/*******************************
 * SECTION 1: CONFIG
 *******************************/

const APP_CONFIG = Object.freeze({
  SHEETS: {
    INPUT: 'INPUT',
    RESULT: 'RESULT',
    MONTHLY: 'MONTHLY',
    CLIENTS: 'CLIENTS',
    CONFIG: 'CONFIG',
    LOGS: 'LOGS'
  },
  INPUT_FIELDS: [
    'Client Name',
    'Client Email',
    'Location',
    'Project Code',
    'Solar Capacity (kW)',
    'No. of Wind Turbines',
    'Wind Turbine Capacity (kW each)',
    'Avg Solar Hours/Day',
    'Avg Wind Capacity Factor (%)',
    'Tariff (₹/kWh)',
    'Monthly Consumption (kWh)',
    'Total Project Cost (₹)',
    'Subsidy / Incentive (₹)',
    'Accelerated Depreciation Applicable (YES/NO)',
    'Proposal Notes'
  ],
  RESULT_FIELDS: [
    'Solar Capacity (kW)',
    'Wind Capacity (kW)',
    'Hybrid Capacity (kW)',
    'Daily Generation (kWh)',
    'Monthly Generation (kWh)',
    'Annual Generation (kWh)',
    'Annual Savings (₹)',
    'GST @5% (₹)',
    'GST @18% (₹)',
    'Gross Project Cost (₹)',
    'Net Project Cost (₹)',
    'Depreciation Benefit Year-1 (₹)',
    'Simple ROI (%)',
    'Payback Period (Years)'
  ],
  CLIENT_FIELDS: ['Date', 'Client Name', 'Email', 'Project Code', 'Deal Value', 'Status', 'Location', 'Owner'],
  CLIENT_STATUS: Object.freeze({ NEW: 'New', PROPOSAL_SENT: 'Proposal Sent', CLOSED: 'Closed' }),
  EMAIL_MODE: Object.freeze({ BRIEF: 'brief', DETAILED: 'detailed' }),
  DEFAULTS: Object.freeze({
    solarPerformanceRatio: 0.78,
    windHoursPerDay: 24,
    daysPerMonth: 30.4375,
    daysPerYear: 365,
    gst5EligibleRatio: 0.7,
    gst5Rate: 0.05,
    gst18Rate: 0.18,
    depreciationRate: 0.4,
    corporateTaxRate: 0.25,
    discountRateForNPV: 0.1
  }),
  CURRENCY_LOCALE: 'en-IN',
  CURRENCY_CODE: 'INR',
  DATE_FORMAT: 'yyyy-MM-dd HH:mm:ss',
  MONTHS: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
});

/*******************************
 * SECTION 2: ENTRYPOINTS
 *******************************/

/**
 * Adds a custom menu on spreadsheet open.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Hybrid SaaS')
    .addItem('Initialize / Repair Sheets', 'initializeSystem')
    .addItem('Generate Project Code', 'generateAndSetProjectCode')
    .addItem('Run Full Calculation', 'runFullCalculation')
    .addItem('Generate PDF Proposal', 'generateProposalPdfFromInput')
    .addItem('Send Proposal Email', 'sendProposalEmailFromInput')
    .addItem('Save / Update CRM', 'upsertClientFromInput')
    .addToUi();
}

/**
 * One-click setup for all required sheets and headers.
 */
function initializeSystem() {
  try {
    ensureSheetStructure_();
    seedConfigIfMissing_();
    SpreadsheetApp.getActive().toast('System initialized successfully.', 'Hybrid SaaS', 5);
    logEvent_('initializeSystem', 'INFO', 'Sheet structure verified and config seeded.');
  } catch (error) {
    handleAndRethrow_('initializeSystem', error);
  }
}

/**
 * Executes the entire business flow:
 * - Validate input
 * - Calculate technical + financial outputs
 * - Write results + monthly table
 * - Save/update CRM
 */
function runFullCalculation() {
  try {
    ensureSheetStructure_();
    const input = readInputData_();
    validateInput_(input);

    const calc = calculateHybridSystem_(input);
    const monthlyRows = calculateMonthlyProjection_(input, calc);

    writeResultSheet_(calc);
    writeMonthlySheet_(monthlyRows);
    upsertClient_(input, calc.netProjectCost);

    SpreadsheetApp.getActive().toast('Calculation completed.', 'Hybrid SaaS', 5);
    logEvent_('runFullCalculation', 'INFO', `Calculation completed for ${input.projectCode}`);
  } catch (error) {
    handleAndRethrow_('runFullCalculation', error);
  }
}

/**
 * Wrapper to auto-generate and place a project code in INPUT sheet.
 */
function generateAndSetProjectCode() {
  try {
    const inputSheet = getSheet_(APP_CONFIG.SHEETS.INPUT);
    const generatedCode = generateProjectCode_();
    const map = readInputMap_();
    map['Project Code'].setValue(generatedCode);
    SpreadsheetApp.getActive().toast(`Project code generated: ${generatedCode}`, 'Hybrid SaaS', 5);
    logEvent_('generateAndSetProjectCode', 'INFO', generatedCode);
  } catch (error) {
    handleAndRethrow_('generateAndSetProjectCode', error);
  }
}

/**
 * Generates proposal PDF and stores in Drive.
 * @returns {GoogleAppsScript.Drive.File} generated PDF file
 */
function generateProposalPdfFromInput() {
  try {
    const input = readInputData_();
    validateInput_(input);
    const calc = calculateHybridSystem_(input);
    const monthlyRows = calculateMonthlyProjection_(input, calc);

    const proposalData = buildProposalData_(input, calc, monthlyRows);
    const pdfFile = buildProposalPdf_(proposalData);

    logEvent_('generateProposalPdfFromInput', 'INFO', `PDF generated: ${pdfFile.getName()}`);
    SpreadsheetApp.getActive().toast('Proposal PDF generated successfully.', 'Hybrid SaaS', 5);

    return pdfFile;
  } catch (error) {
    handleAndRethrow_('generateProposalPdfFromInput', error);
  }
}

/**
 * Sends proposal email using PDF attachment.
 */
function sendProposalEmailFromInput() {
  try {
    const input = readInputData_();
    validateInput_(input);

    const pdfFile = generateProposalPdfFromInput();
    sendProposalEmail_(input, pdfFile, APP_CONFIG.EMAIL_MODE.DETAILED);
    upsertClient_(input, null, APP_CONFIG.CLIENT_STATUS.PROPOSAL_SENT);

    SpreadsheetApp.getActive().toast('Proposal email sent.', 'Hybrid SaaS', 5);
    logEvent_('sendProposalEmailFromInput', 'INFO', `Proposal sent to ${input.clientEmail}`);
  } catch (error) {
    handleAndRethrow_('sendProposalEmailFromInput', error);
  }
}

/**
 * Manual entrypoint for CRM upsert.
 */
function upsertClientFromInput() {
  try {
    const input = readInputData_();
    validateInput_(input);
    const calc = calculateHybridSystem_(input);
    upsertClient_(input, calc.netProjectCost);
    SpreadsheetApp.getActive().toast('CRM updated.', 'Hybrid SaaS', 5);
    logEvent_('upsertClientFromInput', 'INFO', `CRM updated for ${input.projectCode}`);
  } catch (error) {
    handleAndRethrow_('upsertClientFromInput', error);
  }
}

/*******************************
 * SECTION 3: CALCULATION ENGINE
 *******************************/

/**
 * Central business logic engine.
 * @param {Object} input
 * @returns {Object}
 */
function calculateHybridSystem_(input) {
  const rates = getRatesConfig_();

  const solarGenerationDaily = input.solarKw * input.avgSolarHours * rates.solarPerformanceRatio;
  const windTotalKw = input.windTurbines * input.windTurbineKw;
  const windGenerationDaily = windTotalKw * rates.windHoursPerDay * (input.avgWindCapacityFactor / 100);

  const hybridKw = input.solarKw + windTotalKw;
  const dailyGeneration = solarGenerationDaily + windGenerationDaily;
  const monthlyGeneration = dailyGeneration * rates.daysPerMonth;
  const annualGeneration = dailyGeneration * rates.daysPerYear;

  const annualSavings = annualGeneration * input.tariff;

  const gstSplit = calculateGstSplit_(input.totalProjectCost, rates.gst5EligibleRatio, rates.gst5Rate, rates.gst18Rate);
  const grossProjectCost = input.totalProjectCost + gstSplit.gst5 + gstSplit.gst18;
  const netProjectCost = Math.max(grossProjectCost - input.subsidy, 0);

  const depreciationBenefit = calculateDepreciationBenefit_(netProjectCost, rates.depreciationRate, rates.corporateTaxRate);
  const adjustedNetCost = Math.max(netProjectCost - depreciationBenefit, 0);

  const simpleRoi = adjustedNetCost === 0 ? 0 : (annualSavings / adjustedNetCost) * 100;
  const paybackYears = annualSavings === 0 ? 0 : adjustedNetCost / annualSavings;

  return {
    solarKw: input.solarKw,
    windKw: windTotalKw,
    hybridKw,
    dailyGeneration,
    monthlyGeneration,
    annualGeneration,
    annualSavings,
    gst5: gstSplit.gst5,
    gst18: gstSplit.gst18,
    grossProjectCost,
    netProjectCost,
    depreciationBenefit,
    simpleRoi,
    paybackYears
  };
}

/**
 * GST computation with split rates.
 */
function calculateGstSplit_(baseCost, gst5EligibleRatio, gst5Rate, gst18Rate) {
  const costAt5 = baseCost * gst5EligibleRatio;
  const costAt18 = baseCost - costAt5;
  return {
    gst5: costAt5 * gst5Rate,
    gst18: costAt18 * gst18Rate
  };
}

/**
 * Accelerated depreciation benefit approximation (year-1 tax shield).
 */
function calculateDepreciationBenefit_(netCost, depreciationRate, taxRate) {
  return netCost * depreciationRate * taxRate;
}

/*************************************
 * SECTION 4: MONTHLY PROJECTION ENGINE
 *************************************/

/**
 * Uses seasonal multipliers for realistic generation projection.
 * @param {Object} input
 * @param {Object} calc
 * @returns {Array<Array<*>>}
 */
function calculateMonthlyProjection_(input, calc) {
  const multipliers = getMonthlySeasonality_();
  const monthlyConsumption = input.monthlyConsumption;

  const rows = [];
  for (let i = 0; i < APP_CONFIG.MONTHS.length; i++) {
    const month = APP_CONFIG.MONTHS[i];
    const generation = calc.monthlyGeneration * multipliers[i];
    const savings = generation * input.tariff;
    const coverage = monthlyConsumption === 0 ? 0 : (generation / monthlyConsumption) * 100;

    rows.push([
      month,
      round_(monthlyConsumption, 2),
      round_(generation, 2),
      round_(savings, 2),
      round_(coverage, 2)
    ]);
  }

  return rows;
}

/*************************************
 * SECTION 5: PDF + EMAIL AUTOMATION
 *************************************/

/**
 * Constructs view model for proposal HTML rendering.
 */
function buildProposalData_(input, calc, monthlyRows) {
  return {
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy, HH:mm'),
    client: {
      name: input.clientName,
      email: input.clientEmail,
      location: input.location,
      projectCode: input.projectCode
    },
    system: {
      solarKw: calc.solarKw,
      windKw: calc.windKw,
      hybridKw: calc.hybridKw,
      dailyGeneration: calc.dailyGeneration,
      annualGeneration: calc.annualGeneration
    },
    financial: {
      annualSavings: calc.annualSavings,
      grossProjectCost: calc.grossProjectCost,
      netProjectCost: calc.netProjectCost,
      depreciationBenefit: calc.depreciationBenefit,
      roi: calc.simpleRoi,
      payback: calc.paybackYears
    },
    monthlyRows,
    notes: input.proposalNotes || 'Designed for optimum hybrid generation and long-term savings.'
  };
}

/**
 * Renders HTML template and exports as PDF.
 */
function buildProposalPdf_(proposalData) {
  const template = HtmlService.createTemplateFromFile('ProposalTemplate');
  template.data = proposalData;
  template.formatCurrency = formatCurrency_;
  template.formatNumber = (n) => round_(n, 2);

  const html = template.evaluate().getContent();
  const blob = Utilities.newBlob(html, 'text/html', 'proposal.html').getAs('application/pdf');

  const fileName = `Proposal_${proposalData.client.projectCode}.pdf`;
  const file = DriveApp.createFile(blob).setName(fileName);
  return file;
}

/**
 * Sends proposal email with brief or detailed body.
 */
function sendProposalEmail_(input, pdfFile, mode) {
  const calc = calculateHybridSystem_(input);

  const subject = `Hybrid Solar + Wind Proposal | ${input.projectCode} | ${input.clientName}`;
  const briefBody = [
    `Dear ${input.clientName},`,
    '',
    'Please find attached your hybrid solar + wind system proposal.',
    `Estimated annual savings: ${formatCurrency_(calc.annualSavings)}`,
    `Estimated payback: ${round_(calc.paybackYears, 2)} years`,
    '',
    'Regards,',
    'Energy Advisory Team'
  ].join('\n');

  const detailedBody = [
    `Dear ${input.clientName},`,
    '',
    'Thank you for the opportunity to share your tailored hybrid energy proposal.',
    '',
    'Executive Summary:',
    `• Hybrid Capacity: ${round_(calc.hybridKw, 2)} kW`,
    `• Annual Generation: ${round_(calc.annualGeneration, 0)} kWh`,
    `• Annual Savings: ${formatCurrency_(calc.annualSavings)}`,
    `• Net Project Cost: ${formatCurrency_(calc.netProjectCost)}`,
    `• Payback Period: ${round_(calc.paybackYears, 2)} years`,
    '',
    'The attached PDF contains complete technical design assumptions and financial analysis.',
    '',
    'Warm regards,',
    'Energy Advisory Team'
  ].join('\n');

  MailApp.sendEmail({
    to: input.clientEmail,
    subject,
    body: mode === APP_CONFIG.EMAIL_MODE.BRIEF ? briefBody : detailedBody,
    attachments: [pdfFile.getBlob()],
    name: 'Hybrid Energy SaaS'
  });
}

/*******************************
 * SECTION 6: CRM ENGINE
 *******************************/

/**
 * Upsert record in CLIENTS sheet by Project Code.
 */
function upsertClient_(input, dealValue, status) {
  const clientsSheet = getSheet_(APP_CONFIG.SHEETS.CLIENTS);
  const existingRow = findRowByProjectCode_(clientsSheet, input.projectCode);

  const rowData = [
    new Date(),
    input.clientName,
    input.clientEmail,
    input.projectCode,
    dealValue != null ? dealValue : input.totalProjectCost,
    status || APP_CONFIG.CLIENT_STATUS.NEW,
    input.location,
    Session.getActiveUser().getEmail() || 'system'
  ];

  if (existingRow > 0) {
    clientsSheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    clientsSheet.appendRow(rowData);
  }
}

/*******************************
 * SECTION 7: IO + VALIDATION
 *******************************/

/**
 * Reads input from INPUT sheet using key-value layout:
 * Column A = field name, Column B = value.
 */
function readInputData_() {
  const map = readInputMap_();

  const projectCodeCell = map['Project Code'];
  let projectCode = projectCodeCell.getDisplayValue();
  if (!projectCode) {
    projectCode = generateProjectCode_();
    projectCodeCell.setValue(projectCode);
  }

  return {
    clientName: asString_(map['Client Name'].getDisplayValue()),
    clientEmail: asString_(map['Client Email'].getDisplayValue()),
    location: asString_(map['Location'].getDisplayValue()),
    projectCode,
    solarKw: asNumber_(map['Solar Capacity (kW)'].getValue()),
    windTurbines: asNumber_(map['No. of Wind Turbines'].getValue()),
    windTurbineKw: asNumber_(map['Wind Turbine Capacity (kW each)'].getValue()),
    avgSolarHours: asNumber_(map['Avg Solar Hours/Day'].getValue()),
    avgWindCapacityFactor: asNumber_(map['Avg Wind Capacity Factor (%)'].getValue()),
    tariff: asNumber_(map['Tariff (₹/kWh)'].getValue()),
    monthlyConsumption: asNumber_(map['Monthly Consumption (kWh)'].getValue()),
    totalProjectCost: asNumber_(map['Total Project Cost (₹)'].getValue()),
    subsidy: asNumber_(map['Subsidy / Incentive (₹)'].getValue()),
    depreciationApplicable: /^yes$/i.test(asString_(map['Accelerated Depreciation Applicable (YES/NO)'].getDisplayValue())),
    proposalNotes: asString_(map['Proposal Notes'].getDisplayValue())
  };
}

/**
 * Returns map of input labels to value cells.
 */
function readInputMap_() {
  const inputSheet = getSheet_(APP_CONFIG.SHEETS.INPUT);
  const lastRow = Math.max(inputSheet.getLastRow(), APP_CONFIG.INPUT_FIELDS.length);
  const rows = inputSheet.getRange(1, 1, lastRow, 2).getValues();

  const map = {};
  rows.forEach((row, idx) => {
    const key = row[0];
    if (key && APP_CONFIG.INPUT_FIELDS.indexOf(key) > -1) {
      map[key] = inputSheet.getRange(idx + 1, 2);
    }
  });

  APP_CONFIG.INPUT_FIELDS.forEach((field) => {
    if (!map[field]) {
      throw new Error(`Missing INPUT field in sheet: ${field}`);
    }
  });

  return map;
}

/**
 * Validates business-critical inputs.
 */
function validateInput_(input) {
  const errors = [];

  if (!input.clientName) errors.push('Client Name is required.');
  if (!input.clientEmail || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(input.clientEmail)) errors.push('Valid Client Email is required.');
  if (!input.location) errors.push('Location is required.');

  if (input.solarKw < 0) errors.push('Solar Capacity cannot be negative.');
  if (input.windTurbines < 0) errors.push('No. of Wind Turbines cannot be negative.');
  if (input.windTurbineKw < 0) errors.push('Wind Turbine Capacity cannot be negative.');
  if (input.avgSolarHours <= 0 || input.avgSolarHours > 24) errors.push('Avg Solar Hours/Day must be between 0 and 24.');
  if (input.avgWindCapacityFactor < 0 || input.avgWindCapacityFactor > 100) errors.push('Avg Wind Capacity Factor must be between 0 and 100.');
  if (input.tariff <= 0) errors.push('Tariff must be greater than zero.');
  if (input.monthlyConsumption < 0) errors.push('Monthly Consumption cannot be negative.');
  if (input.totalProjectCost <= 0) errors.push('Total Project Cost must be greater than zero.');
  if (input.subsidy < 0) errors.push('Subsidy cannot be negative.');

  if ((input.solarKw + input.windTurbines * input.windTurbineKw) <= 0) {
    errors.push('At least one generation source (solar or wind) must be greater than zero.');
  }

  if (errors.length) {
    throw new Error(`Input validation failed:\n- ${errors.join('\n- ')}`);
  }
}

/**
 * Writes output metrics in RESULT sheet.
 */
function writeResultSheet_(calc) {
  const resultSheet = getSheet_(APP_CONFIG.SHEETS.RESULT);
  const rows = [
    ['Solar Capacity (kW)', calc.solarKw],
    ['Wind Capacity (kW)', calc.windKw],
    ['Hybrid Capacity (kW)', calc.hybridKw],
    ['Daily Generation (kWh)', calc.dailyGeneration],
    ['Monthly Generation (kWh)', calc.monthlyGeneration],
    ['Annual Generation (kWh)', calc.annualGeneration],
    ['Annual Savings (₹)', calc.annualSavings],
    ['GST @5% (₹)', calc.gst5],
    ['GST @18% (₹)', calc.gst18],
    ['Gross Project Cost (₹)', calc.grossProjectCost],
    ['Net Project Cost (₹)', calc.netProjectCost],
    ['Depreciation Benefit Year-1 (₹)', calc.depreciationBenefit],
    ['Simple ROI (%)', calc.simpleRoi],
    ['Payback Period (Years)', calc.paybackYears]
  ];

  resultSheet.clear();
  resultSheet.getRange(1, 1, rows.length, 2).setValues(rows);
  resultSheet.getRange(1, 1, rows.length, 1).setFontWeight('bold');
  resultSheet.autoResizeColumns(1, 2);

  applyCurrencyFormat_(resultSheet, [7, 8, 9, 10, 11, 12]);
}

/**
 * Writes month-wise projections.
 */
function writeMonthlySheet_(monthlyRows) {
  const monthlySheet = getSheet_(APP_CONFIG.SHEETS.MONTHLY);
  const headers = ['Month', 'Consumption (kWh)', 'Projected Generation (kWh)', 'Savings (₹)', 'Coverage %'];

  monthlySheet.clear();
  monthlySheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  monthlySheet.getRange(2, 1, monthlyRows.length, headers.length).setValues(monthlyRows);
  monthlySheet.autoResizeColumns(1, headers.length);

  monthlySheet.getRange(2, 4, monthlyRows.length, 1).setNumberFormat('₹#,##,##0.00');
  monthlySheet.getRange(2, 5, monthlyRows.length, 1).setNumberFormat('0.00%');
  // Coverage is computed in percentage units (e.g. 75 means 75%), so adjust display.
  const coverageRange = monthlySheet.getRange(2, 5, monthlyRows.length, 1);
  const coverageValues = coverageRange.getValues().map((r) => [r[0] / 100]);
  coverageRange.setValues(coverageValues);
}

/*******************************
 * SECTION 8: SHEET BOOTSTRAP
 *******************************/

function ensureSheetStructure_() {
  ensureInputSheet_();
  ensureResultSheet_();
  ensureMonthlySheet_();
  ensureClientsSheet_();
  ensureConfigSheet_();
  ensureLogsSheet_();
}

function ensureInputSheet_() {
  const sh = getOrCreateSheet_(APP_CONFIG.SHEETS.INPUT);
  sh.clear();
  const rows = APP_CONFIG.INPUT_FIELDS.map((field) => [field, '']);
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, rows.length, 1).setFontWeight('bold');
  sh.autoResizeColumns(1, 2);
}

function ensureResultSheet_() {
  const sh = getOrCreateSheet_(APP_CONFIG.SHEETS.RESULT);
  if (sh.getLastRow() === 0) {
    const rows = APP_CONFIG.RESULT_FIELDS.map((field) => [field, '']);
    sh.getRange(1, 1, rows.length, 2).setValues(rows);
    sh.getRange(1, 1, rows.length, 1).setFontWeight('bold');
    sh.autoResizeColumns(1, 2);
  }
}

function ensureMonthlySheet_() {
  const sh = getOrCreateSheet_(APP_CONFIG.SHEETS.MONTHLY);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 5).setValues([
      ['Month', 'Consumption (kWh)', 'Projected Generation (kWh)', 'Savings (₹)', 'Coverage %']
    ]);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold');
    sh.autoResizeColumns(1, 5);
  }
}

function ensureClientsSheet_() {
  const sh = getOrCreateSheet_(APP_CONFIG.SHEETS.CLIENTS);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, APP_CONFIG.CLIENT_FIELDS.length).setValues([APP_CONFIG.CLIENT_FIELDS]);
    sh.getRange(1, 1, 1, APP_CONFIG.CLIENT_FIELDS.length).setFontWeight('bold');
    sh.autoResizeColumns(1, APP_CONFIG.CLIENT_FIELDS.length);
  }
}

function ensureConfigSheet_() {
  const sh = getOrCreateSheet_(APP_CONFIG.SHEETS.CONFIG);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 2).setValues([['Config Key', 'Value']]).setFontWeight('bold');
    sh.autoResizeColumns(1, 2);
  }
}

function ensureLogsSheet_() {
  const sh = getOrCreateSheet_(APP_CONFIG.SHEETS.LOGS);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Function', 'Severity', 'Message']]).setFontWeight('bold');
    sh.autoResizeColumns(1, 4);
  }
}

function seedConfigIfMissing_() {
  const sh = getSheet_(APP_CONFIG.SHEETS.CONFIG);
  const existing = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 0), 2).getValues();
  const map = {};
  existing.forEach((r) => {
    if (r[0]) map[r[0]] = r[1];
  });

  const defaults = {
    solarPerformanceRatio: APP_CONFIG.DEFAULTS.solarPerformanceRatio,
    windHoursPerDay: APP_CONFIG.DEFAULTS.windHoursPerDay,
    daysPerMonth: APP_CONFIG.DEFAULTS.daysPerMonth,
    daysPerYear: APP_CONFIG.DEFAULTS.daysPerYear,
    gst5EligibleRatio: APP_CONFIG.DEFAULTS.gst5EligibleRatio,
    gst5Rate: APP_CONFIG.DEFAULTS.gst5Rate,
    gst18Rate: APP_CONFIG.DEFAULTS.gst18Rate,
    depreciationRate: APP_CONFIG.DEFAULTS.depreciationRate,
    corporateTaxRate: APP_CONFIG.DEFAULTS.corporateTaxRate,
    monthlyMultipliersCsv: '0.92,0.95,1.00,1.03,1.05,0.98,0.90,0.88,0.94,1.00,1.02,0.96'
  };

  const toInsert = [];
  Object.keys(defaults).forEach((key) => {
    if (map[key] == null || map[key] === '') {
      toInsert.push([key, defaults[key]]);
    }
  });

  if (toInsert.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toInsert.length, 2).setValues(toInsert);
  }
}

/*******************************
 * SECTION 9: UTILITIES
 *******************************/

function getRatesConfig_() {
  const cfg = getConfigMap_();
  return {
    solarPerformanceRatio: asNumber_(cfg.solarPerformanceRatio, APP_CONFIG.DEFAULTS.solarPerformanceRatio),
    windHoursPerDay: asNumber_(cfg.windHoursPerDay, APP_CONFIG.DEFAULTS.windHoursPerDay),
    daysPerMonth: asNumber_(cfg.daysPerMonth, APP_CONFIG.DEFAULTS.daysPerMonth),
    daysPerYear: asNumber_(cfg.daysPerYear, APP_CONFIG.DEFAULTS.daysPerYear),
    gst5EligibleRatio: asNumber_(cfg.gst5EligibleRatio, APP_CONFIG.DEFAULTS.gst5EligibleRatio),
    gst5Rate: asNumber_(cfg.gst5Rate, APP_CONFIG.DEFAULTS.gst5Rate),
    gst18Rate: asNumber_(cfg.gst18Rate, APP_CONFIG.DEFAULTS.gst18Rate),
    depreciationRate: asNumber_(cfg.depreciationRate, APP_CONFIG.DEFAULTS.depreciationRate),
    corporateTaxRate: asNumber_(cfg.corporateTaxRate, APP_CONFIG.DEFAULTS.corporateTaxRate)
  };
}

function getMonthlySeasonality_() {
  const cfg = getConfigMap_();
  const csv = asString_(cfg.monthlyMultipliersCsv);
  const arr = csv.split(',').map((x) => asNumber_(x.trim(), 1));
  if (arr.length !== 12) {
    throw new Error('CONFIG.monthlyMultipliersCsv must contain exactly 12 comma-separated values.');
  }
  return arr;
}

function getConfigMap_() {
  const sh = getSheet_(APP_CONFIG.SHEETS.CONFIG);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  const rows = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  const map = {};
  rows.forEach((r) => {
    if (r[0]) map[r[0]] = r[1];
  });
  return map;
}

function generateProjectCode_() {
  const datePart = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const randomPart = Math.floor(Math.random() * 9000 + 1000);
  return `HSP-${datePart}-${randomPart}`;
}

function logEvent_(fnName, severity, message) {
  try {
    const sh = getSheet_(APP_CONFIG.SHEETS.LOGS);
    sh.appendRow([new Date(), fnName, severity, message]);
  } catch (err) {
    console.error('Logging failed:', err);
  }
}

function handleAndRethrow_(fnName, error) {
  const message = `${error && error.message ? error.message : error}`;
  logEvent_(fnName, 'ERROR', message);
  throw new Error(message);
}

function findRowByProjectCode_(sheet, projectCode) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const values = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(projectCode).trim()) {
      return i + 2;
    }
  }
  return -1;
}

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error(`Sheet not found: ${name}. Run initializeSystem first.`);
  return sheet;
}

function asString_(value, fallback) {
  if (value == null) return fallback || '';
  return String(value).trim();
}

function asNumber_(value, fallback) {
  if (value == null || value === '') return fallback != null ? fallback : 0;
  const num = Number(value);
  return Number.isNaN(num) ? (fallback != null ? fallback : 0) : num;
}

function round_(num, digits) {
  const factor = Math.pow(10, digits || 0);
  return Math.round(num * factor) / factor;
}

function formatCurrency_(amount) {
  return new Intl.NumberFormat(APP_CONFIG.CURRENCY_LOCALE, {
    style: 'currency',
    currency: APP_CONFIG.CURRENCY_CODE,
    maximumFractionDigits: 2
  }).format(asNumber_(amount));
}

function applyCurrencyFormat_(sheet, rowIndexes) {
  rowIndexes.forEach((r) => {
    sheet.getRange(r, 2).setNumberFormat('₹#,##,##0.00');
  });
}

/**
 * Optional web app endpoint for future SaaS expansion.
 */
function doGet() {
  return HtmlService.createHtmlOutput('<h3>Hybrid SaaS backend is active.</h3>');
}

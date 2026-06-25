/**
 * Xenin Design Studio ERP - Google Apps Script Backend
 * Complete ERP + CRM + Accounts Management System
 *
 * Architecture:
 * - Google Sheets as Database
 * - Google Apps Script as API
 * - Token-based Authentication (REST API mode)
 * - Session-based Authentication (GAS-served mode)
 * - Activity Logging
 *
 * Deployment Modes:
 * 1. GAS Web App: Frontend served by GAS, uses google.script.run
 * 2. REST API: Frontend hosted externally, uses JSONP/HTTP fetch
 */

const CONFIG = {
  APP_NAME: 'Xenin Design Studio ERP',
  COMPANY: 'Xenin Design Studio',
  TOKEN_EXPIRY_HOURS: 24,
  SHEET_NAMES: {
    USERS: 'Users',
    CLIENTS: 'Clients',
    PROJECTS: 'Projects',
    INCOME: 'Income',
    EXPENSES: 'Expenses',
    ACTIVITIES: 'Activities',
    SETTINGS: 'Settings',
    TOKENS: 'Tokens',
    LOANS: 'Loans',
    BALANCES: 'Balances'
  },
  HEADERS: {
    USERS: ['ID', 'Username', 'Password', 'Name', 'Email', 'Role', 'CreatedAt', 'LastLogin'],
    CLIENTS: ['ID', 'Name', 'Company', 'Phone', 'WhatsApp', 'Email', 'Address', 'Source', 'Notes', 'Status', 'Photo', 'CreatedAt', 'UpdatedAt'],
    PROJECTS: ['ID', 'Name', 'ClientName', 'ClientID', 'Type', 'Location', 'StartDate', 'EndDate', 'Budget', 'Status', 'CreatedAt', 'UpdatedAt', 'QuotationItems', 'QuotationDate', 'QuotationNotes'],
    INCOME: ['ID', 'Date', 'ProjectName', 'ProjectID', 'ClientName', 'Amount', 'PaymentMethod', 'Reference', 'Notes', 'CreatedAt'],
    EXPENSES: ['ID', 'Date', 'ProjectName', 'ProjectID', 'Category', 'Purpose', 'Vendor', 'Amount', 'Notes', 'CreatedAt'],
    ACTIVITIES: ['ID', 'Timestamp', 'User', 'Action', 'Details', 'IP'],
    SETTINGS: ['Key', 'Value', 'UpdatedAt'],
    TOKENS: ['Token', 'Username', 'UserData', 'CreatedAt', 'ExpiresAt'],
    LOANS: ['ID', 'Date', 'Lender', 'Amount', 'Interest', 'Status', 'Notes', 'CreatedAt'],
    BALANCES: ['ID', 'Date', 'AccountType', 'Type', 'Description', 'Amount', 'Notes', 'CreatedAt']
  },
  PROJECT_TYPES: ['Interior Design', 'Architecture', 'Construction', 'Renovation', 'Consultancy'],
  PROJECT_STATUSES: ['Lead', 'Quotation', 'Running', 'Completed', 'Cancelled'],
  EXPENSE_CATEGORIES: ['Materials', 'Labor', 'Transport', 'Site Visit', 'Office Expense', 'Marketing', 'Utility', 'Others'],
  PAYMENT_METHODS: ['Cash', 'Bank', 'Mobile Banking']
};

function initializeSheets() {
  const ss = getSpreadsheet();
  const sheets = CONFIG.SHEET_NAMES;

  Object.keys(sheets).forEach(key => {
    const name = sheets[key];
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(CONFIG.HEADERS[key]);
      if (key === 'USERS') {
        sheet.appendRow(['1', 'admin', 'admin123', 'Administrator', 'admin@xenindesign.com', 'Admin', new Date().toISOString(), '']);
      }
      if (key === 'SETTINGS') {
        sheet.appendRow(['company_name', CONFIG.COMPANY, new Date().toISOString()]);
        sheet.appendRow(['currency', 'INR', new Date().toISOString()]);
        sheet.appendRow(['timezone', 'Asia/Kolkata', new Date().toISOString()]);
        sheet.appendRow(['theme', 'light', new Date().toISOString()]);
        sheet.appendRow(['tax_rate', '18', new Date().toISOString()]);
        sheet.appendRow(['language', 'en', new Date().toISOString()]);
        sheet.appendRow(['logo', '', new Date().toISOString()]);
      }
    }
  });

  return { success: true };
}

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    initializeSheets();
    sheet = ss.getSheetByName(name);
  }
  return sheet;
}

function getNextId(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 1;
  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const numPart = String(data[i][0]).replace(/^[A-Z]+#/, '');
    const id = parseInt(numPart);
    if (!isNaN(id) && id > maxId) maxId = id;
  }
  return maxId + 1;
}

function sanitizeInput(value) {
  if (typeof value === 'string') {
    return value.trim().replace(/<[^>]*>/g, '');
  }
  return value;
}

function validateRequired(value, fieldName) {
  if (!value || (typeof value === 'string' && value.trim() === '')) {
    throw new Error(fieldName + ' is required');
  }
  return sanitizeInput(value);
}

function getCurrentDateString() {
  return new Date().toISOString();
}

function createResponse(success, data, message) {
  var resp = { success: success };
  if (data) {
    Object.keys(data).forEach(function(k) { resp[k] = data[k]; });
  }
  if (message) resp.message = message;
  return resp;
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleJsonpRequest(e);
  }
  return serveHtmlApp();
}

function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    return handleApiRequest(params);
  } catch (error) {
    return sendJsonResponse({ success: false, message: 'Invalid request: ' + error.toString() });
  }
}

function serveHtmlApp() {
  var html = HtmlService.createTemplateFromFile('App');
  return html.evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function handleJsonpRequest(e) {
  var params = e.parameter;
  var callback = params.callback || 'callback';
  var action = params.action;
  var cleanParams = {};
  Object.keys(params).forEach(function(key) {
    if (key !== 'action' && key !== 'callback') {
      cleanParams[key] = params[key];
    }
  });
  var result;
  try {
    result = routeAction(action, cleanParams);
  } catch (error) {
    result = createResponse(false, {}, error.toString());
  }
  var output = callback + '(' + JSON.stringify(result) + ')';
  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function handleApiRequest(params) {
  var action = params.action;
  delete params.action;
  var result;
  try {
    result = routeAction(action, params);
  } catch (error) {
    result = createResponse(false, {}, error.toString());
  }
  return sendJsonResponse(result);
}

function sendJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function routeAction(action, params) {
  switch (action) {
    case 'login': return login(params.username, params.password);
    case 'logout': return logout(params);
    case 'checkSession': return checkSession(params);
    case 'changePassword': return changePassword(params.username, params.oldPassword, params.newPassword);
    case 'getClients': return getClients(params);
    case 'addClient': return addClient(params);
    case 'updateClient': return updateClient(params);
    case 'deleteClient': return deleteClient(params);
    case 'getProjects': return getProjects(params);
    case 'addProject': return addProject(params);
    case 'updateProject': return updateProject(params);
    case 'deleteProject': return deleteProject(params);
    case 'getIncome': return getIncome(params);
    case 'addIncome': return addIncome(params);
    case 'updateIncome': return updateIncome(params);
    case 'deleteIncome': return deleteIncome(params);
    case 'generateReceipt': return generateReceipt(params);
    case 'getExpenses': return getExpenses(params);
    case 'addExpense': return addExpense(params);
    case 'updateExpense': return updateExpense(params);
    case 'deleteExpense': return deleteExpense(params);
    case 'getLoans': return getLoans(params);
    case 'addLoan': return addLoan(params);
    case 'updateLoan': return updateLoan(params);
    case 'deleteLoan': return deleteLoan(params);
    case 'getBalances': return getBalances(params);
    case 'addBalance': return addBalance(params);
    case 'updateBalance': return updateBalance(params);
    case 'deleteBalance': return deleteBalance(params);
    case 'getDashboardData': return getDashboardData();
    case 'getReport': return getReport(params);
    case 'exportPDF': return exportPDF(params);
    case 'exportCSV': return exportCSV(params);
    case 'getActivities': return getActivities(params);
    case 'getSettings': return getSettings();
    case 'updateSettings': return updateSettings(params);
    default:
      return createResponse(false, {}, 'Unknown action: ' + action);
  }
}

function login(username, password) {
  try {
    var sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === username && data[i][2] === password) {
        var user = {
          id: data[i][0],
          username: data[i][1],
          name: data[i][3],
          email: data[i][4],
          role: data[i][5]
        };

        sheet.getRange(i + 1, 8).setValue(getCurrentDateString());
        logActivity('Login', 'User ' + username + ' logged in');

        var token = generateToken(username, user);

        try {
          PropertiesService.getUserProperties().setProperty('session_user', JSON.stringify(user));
          PropertiesService.getUserProperties().setProperty('session_key', ScriptApp.getTemporaryActiveUserKey());
        } catch (e) {}

        return createResponse(true, { user: user, token: token }, 'Login successful');
      }
    }

    return createResponse(false, {}, 'Invalid username or password');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function generateToken(username, user) {
  var token = Utilities.getUuid();
  var now = new Date();
  var expires = new Date(now.getTime() + CONFIG.TOKEN_EXPIRY_HOURS * 60 * 60 * 1000);
  var tokenSheet = getSheet(CONFIG.SHEET_NAMES.TOKENS);
  var tokenId = getNextId(CONFIG.SHEET_NAMES.TOKENS);
  tokenSheet.appendRow([token, username, JSON.stringify(user), now.toISOString(), expires.toISOString()]);
  return token;
}

function validateToken(token) {
  if (!token) return { authenticated: false };
  try {
    var sheet = getSheet(CONFIG.SHEET_NAMES.TOKENS);
    var data = sheet.getDataRange().getValues();
    var now = new Date();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === token) {
        var expires = new Date(data[i][4]);
        if (expires > now) {
          return {
            authenticated: true,
            username: data[i][1],
            user: JSON.parse(data[i][2])
          };
        }
        sheet.deleteRow(i + 1);
        return { authenticated: false, message: 'Token expired' };
      }
    }
  } catch (e) {
    return { authenticated: false, message: e.toString() };
  }
  return { authenticated: false, message: 'Invalid token' };
}

function checkSession(params) {
  try {
    var sessionUser = PropertiesService.getUserProperties().getProperty('session_user');
    if (sessionUser) {
      return createResponse(true, { user: JSON.parse(sessionUser) });
    }
  } catch (e) {}

  if (params && params.token) {
    var session = validateToken(params.token);
    if (session.authenticated) {
      return createResponse(true, { user: session.user });
    }
  }

  return createResponse(false, {}, 'No active session');
}

function logout(params) {
  try {
    logActivity('Logout', 'User logged out');

    try {
      PropertiesService.getUserProperties().deleteProperty('session_user');
      PropertiesService.getUserProperties().deleteProperty('session_key');
    } catch (e) {}

    if (params && params.token) {
      var sheet = getSheet(CONFIG.SHEET_NAMES.TOKENS);
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === params.token) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
    }

    return createResponse(true, {}, 'Logged out successfully');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function changePassword(username, oldPassword, newPassword) {
  try {
    var sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === username && data[i][2] === oldPassword) {
        sheet.getRange(i + 1, 3).setValue(newPassword);
        logActivity('Password Change', 'Password changed for user ' + username);
        return createResponse(true, {}, 'Password changed successfully');
      }
    }

    return createResponse(false, {}, 'Current password is incorrect');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function addClient(data) {
  try {
    var id = 'XNC#' + getNextId(CONFIG.SHEET_NAMES.CLIENTS);
    var now = getCurrentDateString();

    var rowData = [
      id,
      validateRequired(data.name, 'Client Name'),
      sanitizeInput(data.company || ''),
      sanitizeInput(data.phone || ''),
      sanitizeInput(data.whatsapp || ''),
      sanitizeInput(data.email || ''),
      sanitizeInput(data.address || ''),
      sanitizeInput(data.source || ''),
      sanitizeInput(data.notes || ''),
      sanitizeInput(data.status || 'Active'),
      data.photo || '',
      now,
      now
    ];

    var sheet = getSheet(CONFIG.SHEET_NAMES.CLIENTS);
    sheet.appendRow(rowData);

    logActivity('Add Client', 'Added client: ' + data.name);

    return createResponse(true, {
      client: {
        id: id, name: data.name, company: data.company, phone: data.phone,
        whatsapp: data.whatsapp, email: data.email, address: data.address,
        source: data.source, notes: data.notes, status: data.status || 'Active',
        photo: data.photo || ''
      }
    }, 'Client added successfully');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function updateClient(params) {
  try {
    var id = typeof params === 'object' ? params.id : params;
    var data = params.data || params;
    var sheet = getSheet(CONFIG.SHEET_NAMES.CLIENTS);
    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        var row = i + 1;
        var updates = {};
        if (data.name !== undefined) updates[2] = sanitizeInput(data.name);
        if (data.company !== undefined) updates[3] = sanitizeInput(data.company || '');
        if (data.phone !== undefined) updates[4] = sanitizeInput(data.phone || '');
        if (data.whatsapp !== undefined) updates[5] = sanitizeInput(data.whatsapp || '');
        if (data.email !== undefined) updates[6] = sanitizeInput(data.email || '');
        if (data.address !== undefined) updates[7] = sanitizeInput(data.address || '');
        if (data.source !== undefined) updates[8] = sanitizeInput(data.source || '');
        if (data.notes !== undefined) updates[9] = sanitizeInput(data.notes || '');
        if (data.status !== undefined) updates[10] = sanitizeInput(data.status || 'Active');
        if (data.photo !== undefined) updates[11] = data.photo || '';
        updates[13] = getCurrentDateString();

        Object.keys(updates).forEach(function(col) {
          sheet.getRange(row, parseInt(col)).setValue(updates[col]);
        });

        logActivity('Update Client', 'Updated client: ' + (data.name || id));
        return createResponse(true, {}, 'Client updated successfully');
      }
    }

    return createResponse(false, {}, 'Client not found');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function deleteClient(params) {
  try {
    var id = typeof params === 'object' ? params.id : params;
    var sheet = getSheet(CONFIG.SHEET_NAMES.CLIENTS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        logActivity('Delete Client', 'Deleted client: ' + data[i][1]);
        return createResponse(true, {}, 'Client deleted successfully');
      }
    }

    return createResponse(false, {}, 'Client not found');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function getClients(filter) {
  try {
    var sheet = getSheet(CONFIG.SHEET_NAMES.CLIENTS);
    var data = sheet.getDataRange().getValues();
    var clients = [];

    for (var i = 1; i < data.length; i++) {
      var client = {
        id: data[i][0],
        name: data[i][1],
        company: data[i][2],
        phone: data[i][3],
        whatsapp: data[i][4],
        email: data[i][5],
        address: data[i][6],
        source: data[i][7],
        notes: data[i][8],
        status: data[i][9],
        photo: data[i][10],
        createdAt: data[i][11],
        updatedAt: data[i][12]
      };

      if (filter && filter.search) {
        var search = filter.search.toLowerCase();
        var matches = Object.values(client).some(function(val) {
          return String(val).toLowerCase().includes(search);
        });
        if (!matches) continue;
      }
      if (filter && filter.status && client.status !== filter.status) continue;

      clients.push(client);
    }

    return createResponse(true, { clients: clients });
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function addProject(data) {
  try {
    var id = 'XNP#' + getNextId(CONFIG.SHEET_NAMES.PROJECTS);
    var now = getCurrentDateString();

    var rowData = [
      id,
      validateRequired(data.name, 'Project Name'),
      sanitizeInput(data.clientName || ''),
      sanitizeInput(data.clientID || ''),
      sanitizeInput(data.type || ''),
      sanitizeInput(data.location || ''),
      sanitizeInput(data.startDate || ''),
      sanitizeInput(data.endDate || ''),
      parseFloat(data.budget) || 0,
      sanitizeInput(data.status || 'Lead'),
      now,
      now,
      typeof data.quotationItems === 'string' ? data.quotationItems : (data.quotationItems ? JSON.stringify(data.quotationItems) : ''),
      sanitizeInput(data.quotationDate || ''),
      sanitizeInput(data.quotationNotes || '')
    ];

    var sheet = getSheet(CONFIG.SHEET_NAMES.PROJECTS);
    sheet.appendRow(rowData);

    logActivity('Add Project', 'Added project: ' + data.name);

    return createResponse(true, { project: { id: id, name: data.name } }, 'Project added successfully');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function updateProject(params) {
  try {
    var id = typeof params === 'object' ? params.id : params;
    var data = params.data || params;
    var sheet = getSheet(CONFIG.SHEET_NAMES.PROJECTS);
    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        var row = i + 1;
        var updates = {};
        if (data.name !== undefined) updates[2] = sanitizeInput(data.name);
        if (data.clientName !== undefined) updates[3] = sanitizeInput(data.clientName || '');
        if (data.clientID !== undefined) updates[4] = sanitizeInput(data.clientID || '');
        if (data.type !== undefined) updates[5] = sanitizeInput(data.type || '');
        if (data.location !== undefined) updates[6] = sanitizeInput(data.location || '');
        if (data.startDate !== undefined) updates[7] = sanitizeInput(data.startDate || '');
        if (data.endDate !== undefined) updates[8] = sanitizeInput(data.endDate || '');
        if (data.budget !== undefined) updates[9] = parseFloat(data.budget) || 0;
        if (data.status !== undefined) updates[10] = sanitizeInput(data.status || 'Lead');
        if (data.quotationItems !== undefined) updates[13] = typeof data.quotationItems === 'string' ? data.quotationItems : JSON.stringify(data.quotationItems);
        if (data.quotationDate !== undefined) updates[14] = sanitizeInput(data.quotationDate || '');
        if (data.quotationNotes !== undefined) updates[15] = sanitizeInput(data.quotationNotes || '');
        updates[12] = getCurrentDateString();

        Object.keys(updates).forEach(function(col) {
          sheet.getRange(row, parseInt(col)).setValue(updates[col]);
        });

        logActivity('Update Project', 'Updated project: ' + (data.name || id));
        return createResponse(true, {}, 'Project updated successfully');
      }
    }

    return createResponse(false, {}, 'Project not found');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function deleteProject(params) {
  try {
    var id = typeof params === 'object' ? params.id : params;
    var sheet = getSheet(CONFIG.SHEET_NAMES.PROJECTS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        logActivity('Delete Project', 'Deleted project: ' + data[i][1]);
        return createResponse(true, {}, 'Project deleted successfully');
      }
    }

    return createResponse(false, {}, 'Project not found');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function getProjects(filter) {
  try {
    var sheet = getSheet(CONFIG.SHEET_NAMES.PROJECTS);
    var data = sheet.getDataRange().getValues();
    var projects = [];

    for (var i = 1; i < data.length; i++) {
      var project = {
        id: data[i][0],
        name: data[i][1],
        clientName: data[i][2],
        clientID: data[i][3],
        type: data[i][4],
        location: data[i][5],
        startDate: data[i][6],
        endDate: data[i][7],
        budget: data[i][8],
        status: data[i][9],
        createdAt: data[i][10],
        updatedAt: data[i][11],
        quotationItems: data[i][12] ? (typeof data[i][12] === 'string' && data[i][12].charAt(0) === '[' ? JSON.parse(data[i][12]) : data[i][12]) : [],
        quotationDate: data[i][13] || '',
        quotationNotes: data[i][14] || ''
      };

      if (filter && filter.search) {
        var search = filter.search.toLowerCase();
        var matches = Object.values(project).some(function(val) {
          return String(val).toLowerCase().includes(search);
        });
        if (!matches) continue;
      }
      if (filter && filter.status && project.status !== filter.status) continue;
      if (filter && filter.type && project.type !== filter.type) continue;
      if (filter && filter.client && project.clientName !== filter.client) continue;

      projects.push(project);
    }

    return createResponse(true, { projects: projects });
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function addIncome(data) {
  try {
    var id = 'XNI#' + getNextId(CONFIG.SHEET_NAMES.INCOME);
    var now = getCurrentDateString();
    var date = data.date || formatDateFromISO(now);

    var rowData = [
      id,
      sanitizeInput(date),
      sanitizeInput(data.projectName || ''),
      sanitizeInput(data.projectID || ''),
      sanitizeInput(data.clientName || ''),
      parseFloat(data.amount) || 0,
      sanitizeInput(data.paymentMethod || 'Cash'),
      sanitizeInput(data.reference || ''),
      sanitizeInput(data.notes || ''),
      now
    ];

    var sheet = getSheet(CONFIG.SHEET_NAMES.INCOME);
    sheet.appendRow(rowData);

    logActivity('Income Entry', 'Added income: ' + data.projectName + ' - ' + data.amount);

    var receiptHtml = generateReceiptHtml({
      id: 'RCP-' + String(id).slice(-4),
      date: date,
      clientName: data.clientName,
      projectName: data.projectName,
      amount: data.amount,
      paymentMethod: data.paymentMethod
    });

    return createResponse(true, {
      income: { id: id, date: date, projectName: data.projectName, projectID: data.projectID, clientName: data.clientName, amount: parseFloat(data.amount) || 0, paymentMethod: data.paymentMethod || 'Cash', reference: data.reference || '', notes: data.notes || '' },
      receiptHtml: receiptHtml
    }, 'Income added successfully');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function updateIncome(params) {
  try {
    var id = typeof params === 'object' ? params.id : params;
    var data = params.data || params;
    var sheet = getSheet(CONFIG.SHEET_NAMES.INCOME);
    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        var row = i + 1;
        var updates = {};
        if (data.date !== undefined) updates[2] = sanitizeInput(data.date);
        if (data.projectName !== undefined) updates[3] = sanitizeInput(data.projectName || '');
        if (data.projectID !== undefined) updates[4] = sanitizeInput(data.projectID || '');
        if (data.clientName !== undefined) updates[5] = sanitizeInput(data.clientName || '');
        if (data.amount !== undefined) updates[6] = parseFloat(data.amount) || 0;
        if (data.paymentMethod !== undefined) updates[7] = sanitizeInput(data.paymentMethod || 'Cash');
        if (data.reference !== undefined) updates[8] = sanitizeInput(data.reference || '');
        if (data.notes !== undefined) updates[9] = sanitizeInput(data.notes || '');

        Object.keys(updates).forEach(function(col) {
          sheet.getRange(row, parseInt(col)).setValue(updates[col]);
        });

        logActivity('Update Income', 'Updated income entry: ' + id);
        return createResponse(true, {}, 'Income updated successfully');
      }
    }

    return createResponse(false, {}, 'Income entry not found');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function deleteIncome(params) {
  try {
    var id = typeof params === 'object' ? params.id : params;
    var sheet = getSheet(CONFIG.SHEET_NAMES.INCOME);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        logActivity('Delete Income', 'Deleted income entry: ' + id);
        return createResponse(true, {}, 'Income entry deleted successfully');
      }
    }

    return createResponse(false, {}, 'Income entry not found');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function getIncome(filter) {
  try {
    var sheet = getSheet(CONFIG.SHEET_NAMES.INCOME);
    var data = sheet.getDataRange().getValues();
    var incomeList = [];

    for (var i = 1; i < data.length; i++) {
      var income = {
        id: data[i][0],
        date: data[i][1],
        projectName: data[i][2],
        projectID: data[i][3],
        clientName: data[i][4],
        amount: data[i][5],
        paymentMethod: data[i][6],
        reference: data[i][7],
        notes: data[i][8],
        createdAt: data[i][9]
      };

      if (filter && filter.search) {
        var search = filter.search.toLowerCase();
        var matches = Object.values(income).some(function(val) {
          return String(val).toLowerCase().includes(search);
        });
        if (!matches) continue;
      }
      if (filter && filter.project && income.projectName !== filter.project) continue;
      if (filter && filter.client && income.clientName !== filter.client) continue;
      if (filter && filter.fromDate && income.date < filter.fromDate) continue;
      if (filter && filter.toDate && income.date > filter.toDate) continue;

      incomeList.push(income);
    }

    return createResponse(true, { incomeList: incomeList });
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function generateReceiptHtml(data) {
  var companyName = CONFIG.COMPANY;
  var receiptNo = data.id || 'RCP-' + Date.now();
  var date = data.date || formatDateFromISO(new Date().toISOString());
  var amount = parseFloat(data.amount) || 0;
  var amountWords = numberToWords(amount);

  return '' +
    '<div style="font-family: Segoe UI, Arial, sans-serif; max-width: 700px; margin: 0 auto; padding: 40px; background: #fff; border: 2px solid #1a1a2e;">' +
    '<div style="text-align: center; border-bottom: 3px double #1a1a2e; padding-bottom: 20px; margin-bottom: 20px;">' +
    '<h1 style="margin: 0; color: #1a1a2e; font-size: 28px; letter-spacing: 2px;">' + companyName + '</h1>' +
    '<p style="margin: 5px 0; color: #666; font-size: 14px;">Interior Design | Architecture | Construction</p></div>' +
    '<div style="display: flex; justify-content: space-between; margin-bottom: 30px;">' +
    '<div><h2 style="margin: 0; color: #1a1a2e; font-size: 24px;">RECEIPT</h2><p style="margin: 5px 0; color: #666;">Receipt #: ' + receiptNo + '</p></div>' +
    '<div style="text-align: right;"><p style="margin: 5px 0; color: #666;">Date: ' + date + '</p></div></div>' +
    '<table style="width: 100%; border-collapse: collapse; margin-bottom: 30px;">' +
    '<tr><td style="padding: 10px 0; color: #666; width: 150px;">Received from</td><td style="padding: 10px 0; font-weight: bold; border-bottom: 1px solid #ddd;">' + (data.clientName || 'N/A') + '</td></tr>' +
    '<tr><td style="padding: 10px 0; color: #666;">Project</td><td style="padding: 10px 0; font-weight: bold; border-bottom: 1px solid #ddd;">' + (data.projectName || 'N/A') + '</td></tr>' +
    '<tr><td style="padding: 10px 0; color: #666;">Amount</td><td style="padding: 10px 0; font-weight: bold; font-size: 18px; color: #1a1a2e; border-bottom: 1px solid #ddd;">\u09F3 ' + amount.toLocaleString('en-IN') + '/-</td></tr>' +
    '<tr><td style="padding: 10px 0; color: #666;">In Words</td><td style="padding: 10px 0; font-style: italic; border-bottom: 1px solid #ddd;">' + amountWords + '</td></tr>' +
    '<tr><td style="padding: 10px 0; color: #666;">Payment Method</td><td style="padding: 10px 0; border-bottom: 1px solid #ddd;">' + (data.paymentMethod || 'N/A') + '</td></tr></table>' +
    '<div style="margin-top: 50px; display: flex; justify-content: space-between;">' +
    '<div style="text-align: center;"><div style="border-top: 1px solid #333; padding-top: 5px; width: 200px;">Client Signature</div></div>' +
    '<div style="text-align: center;"><div style="border-top: 1px solid #333; padding-top: 5px; width: 200px;">Authorized Signature</div></div></div>' +
    '<div style="margin-top: 30px; text-align: center; color: #999; font-size: 12px; border-top: 1px solid #eee; padding-top: 15px;">' +
    '<p>This is a computer-generated receipt. No signature required.</p>' +
    '<p>' + companyName + ' | Thank you for your business!</p></div></div>';
}

function generateReceipt(params) {
  try {
    var incomeId = typeof params === 'object' ? params.incomeId : params;
    var sheet = getSheet(CONFIG.SHEET_NAMES.INCOME);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(incomeId)) {
        var receiptData = {
          id: 'RCP-' + String(data[i][0]).slice(-4),
          date: data[i][1],
          clientName: data[i][4],
          projectName: data[i][2],
          amount: data[i][5],
          paymentMethod: data[i][6]
        };
        return createResponse(true, { receiptHtml: generateReceiptHtml(receiptData) });
      }
    }

    return createResponse(false, {}, 'Income entry not found');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function addExpense(data) {
  try {
    var id = 'XNE#' + getNextId(CONFIG.SHEET_NAMES.EXPENSES);
    var now = getCurrentDateString();
    var date = data.date || formatDateFromISO(now);

    var rowData = [
      id,
      sanitizeInput(date),
      sanitizeInput(data.projectName || ''),
      sanitizeInput(data.projectID || ''),
      sanitizeInput(data.category || 'Others'),
      sanitizeInput(data.purpose || ''),
      sanitizeInput(data.vendor || ''),
      parseFloat(data.amount) || 0,
      sanitizeInput(data.notes || ''),
      now
    ];

    var sheet = getSheet(CONFIG.SHEET_NAMES.EXPENSES);
    sheet.appendRow(rowData);

    logActivity('Expense Entry', 'Added expense: ' + data.category + ' - ' + data.amount);

    return createResponse(true, { expense: { id: id, date: date, projectName: data.projectName, projectID: data.projectID, category: data.category || 'Others', purpose: data.purpose || '', vendor: data.vendor || '', amount: parseFloat(data.amount) || 0, notes: data.notes || '' } }, 'Expense added successfully');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function updateExpense(params) {
  try {
    var id = typeof params === 'object' ? params.id : params;
    var data = params.data || params;
    var sheet = getSheet(CONFIG.SHEET_NAMES.EXPENSES);
    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        var row = i + 1;
        var updates = {};
        if (data.date !== undefined) updates[2] = sanitizeInput(data.date);
        if (data.projectName !== undefined) updates[3] = sanitizeInput(data.projectName || '');
        if (data.projectID !== undefined) updates[4] = sanitizeInput(data.projectID || '');
        if (data.category !== undefined) updates[5] = sanitizeInput(data.category || 'Others');
        if (data.purpose !== undefined) updates[6] = sanitizeInput(data.purpose || '');
        if (data.vendor !== undefined) updates[7] = sanitizeInput(data.vendor || '');
        if (data.amount !== undefined) updates[8] = parseFloat(data.amount) || 0;
        if (data.notes !== undefined) updates[9] = sanitizeInput(data.notes || '');

        Object.keys(updates).forEach(function(col) {
          sheet.getRange(row, parseInt(col)).setValue(updates[col]);
        });

        logActivity('Update Expense', 'Updated expense: ' + id);
        return createResponse(true, {}, 'Expense updated successfully');
      }
    }

    return createResponse(false, {}, 'Expense not found');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function deleteExpense(params) {
  try {
    var id = typeof params === 'object' ? params.id : params;
    var sheet = getSheet(CONFIG.SHEET_NAMES.EXPENSES);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        logActivity('Delete Expense', 'Deleted expense: ' + id);
        return createResponse(true, {}, 'Expense deleted successfully');
      }
    }

    return createResponse(false, {}, 'Expense not found');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function getExpenses(filter) {
  try {
    var sheet = getSheet(CONFIG.SHEET_NAMES.EXPENSES);
    var data = sheet.getDataRange().getValues();
    var expenses = [];

    for (var i = 1; i < data.length; i++) {
      var expense = {
        id: data[i][0],
        date: data[i][1],
        projectName: data[i][2],
        projectID: data[i][3],
        category: data[i][4],
        purpose: data[i][5],
        vendor: data[i][6],
        amount: data[i][7],
        notes: data[i][8],
        createdAt: data[i][9]
      };

      if (filter && filter.search) {
        var search = filter.search.toLowerCase();
        var matches = Object.values(expense).some(function(val) {
          return String(val).toLowerCase().includes(search);
        });
        if (!matches) continue;
      }
      if (filter && filter.category && expense.category !== filter.category) continue;
      if (filter && filter.project && expense.projectName !== filter.project) continue;
      if (filter && filter.fromDate && expense.date < filter.fromDate) continue;
      if (filter && filter.toDate && expense.date > filter.toDate) continue;

      expenses.push(expense);
    }

    return createResponse(true, { expenses: expenses });
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function addLoan(data) {
  try {
    var id = 'XNL#' + getNextId(CONFIG.SHEET_NAMES.LOANS);
    var now = getCurrentDateString();
    var date = data.date || formatDateFromISO(now);

    var rowData = [
      id,
      sanitizeInput(date),
      sanitizeInput(data.lender || ''),
      parseFloat(data.amount) || 0,
      sanitizeInput(data.interest || '0'),
      sanitizeInput(data.status || 'Active'),
      sanitizeInput(data.notes || ''),
      now
    ];

    var sheet = getSheet(CONFIG.SHEET_NAMES.LOANS);
    sheet.appendRow(rowData);

    logActivity('Loan Entry', 'Added loan: ' + data.lender + ' - ' + data.amount);

    return createResponse(true, { loan: { id: id, date: date, lender: data.lender || '', amount: parseFloat(data.amount) || 0, interest: data.interest || '0', status: data.status || 'Active', notes: data.notes || '' } }, 'Loan added successfully');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function updateLoan(params) {
  try {
    var id = typeof params === 'object' ? params.id : params;
    var data = params.data || params;
    var sheet = getSheet(CONFIG.SHEET_NAMES.LOANS);
    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        var row = i + 1;
        var updates = {};
        if (data.date !== undefined) updates[2] = sanitizeInput(data.date);
        if (data.lender !== undefined) updates[3] = sanitizeInput(data.lender || '');
        if (data.amount !== undefined) updates[4] = parseFloat(data.amount) || 0;
        if (data.interest !== undefined) updates[5] = sanitizeInput(data.interest || '0');
        if (data.status !== undefined) updates[6] = sanitizeInput(data.status || 'Active');
        if (data.notes !== undefined) updates[7] = sanitizeInput(data.notes || '');

        Object.keys(updates).forEach(function(col) {
          sheet.getRange(row, parseInt(col)).setValue(updates[col]);
        });

        logActivity('Update Loan', 'Updated loan entry: ' + id);
        return createResponse(true, {}, 'Loan updated successfully');
      }
    }

    return createResponse(false, {}, 'Loan not found');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function deleteLoan(params) {
  try {
    var id = typeof params === 'object' ? params.id : params;
    var sheet = getSheet(CONFIG.SHEET_NAMES.LOANS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        logActivity('Delete Loan', 'Deleted loan entry: ' + id);
        return createResponse(true, {}, 'Loan deleted successfully');
      }
    }

    return createResponse(false, {}, 'Loan not found');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function getLoans(filter) {
  try {
    var sheet = getSheet(CONFIG.SHEET_NAMES.LOANS);
    var data = sheet.getDataRange().getValues();
    var loans = [];

    for (var i = 1; i < data.length; i++) {
      var loan = {
        id: data[i][0],
        date: data[i][1],
        lender: data[i][2],
        amount: data[i][3],
        interest: data[i][4],
        status: data[i][5],
        notes: data[i][6],
        createdAt: data[i][7]
      };

      if (filter && filter.search) {
        var search = filter.search.toLowerCase();
        var matches = Object.values(loan).some(function(val) {
          return String(val).toLowerCase().includes(search);
        });
        if (!matches) continue;
      }

      loans.push(loan);
    }

    return createResponse(true, { loans: loans });
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function addBalance(data) {
  try {
    var id = 'XNB#' + getNextId(CONFIG.SHEET_NAMES.BALANCES);
    var now = getCurrentDateString();
    var date = data.date || formatDateFromISO(now);

    var rowData = [
      id,
      sanitizeInput(date),
      sanitizeInput(data.accountType || 'cash'),
      sanitizeInput(data.type || 'deposit'),
      sanitizeInput(data.description || ''),
      parseFloat(data.amount) || 0,
      sanitizeInput(data.notes || ''),
      now
    ];

    var sheet = getSheet(CONFIG.SHEET_NAMES.BALANCES);
    sheet.appendRow(rowData);

    logActivity('Balance Entry', 'Added ' + data.type + ': ' + data.description + ' - ' + data.amount);

    return createResponse(true, { balance: { id: id, date: date, accountType: data.accountType || 'cash', type: data.type || 'deposit', description: data.description || '', amount: parseFloat(data.amount) || 0, notes: data.notes || '' } }, 'Balance transaction added successfully');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function updateBalance(params) {
  try {
    var id = typeof params === 'object' ? params.id : params;
    var data = params.data || params;
    var sheet = getSheet(CONFIG.SHEET_NAMES.BALANCES);
    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        var row = i + 1;
        var updates = {};
        if (data.date !== undefined) updates[2] = sanitizeInput(data.date);
        if (data.accountType !== undefined) updates[3] = sanitizeInput(data.accountType || 'cash');
        if (data.type !== undefined) updates[4] = sanitizeInput(data.type || 'deposit');
        if (data.description !== undefined) updates[5] = sanitizeInput(data.description || '');
        if (data.amount !== undefined) updates[6] = parseFloat(data.amount) || 0;
        if (data.notes !== undefined) updates[7] = sanitizeInput(data.notes || '');

        Object.keys(updates).forEach(function(col) {
          sheet.getRange(row, parseInt(col)).setValue(updates[col]);
        });

        logActivity('Update Balance', 'Updated balance entry: ' + id);
        return createResponse(true, {}, 'Transaction updated successfully');
      }
    }

    return createResponse(false, {}, 'Transaction not found');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function deleteBalance(params) {
  try {
    var id = typeof params === 'object' ? params.id : params;
    var sheet = getSheet(CONFIG.SHEET_NAMES.BALANCES);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        logActivity('Delete Balance', 'Deleted balance entry: ' + id);
        return createResponse(true, {}, 'Transaction deleted successfully');
      }
    }

    return createResponse(false, {}, 'Transaction not found');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function getBalances(filter) {
  try {
    var sheet = getSheet(CONFIG.SHEET_NAMES.BALANCES);
    var data = sheet.getDataRange().getValues();
    var balances = [];

    for (var i = 1; i < data.length; i++) {
      var balance = {
        id: data[i][0],
        date: data[i][1],
        accountType: data[i][2],
        type: data[i][3],
        description: data[i][4],
        amount: data[i][5],
        notes: data[i][6],
        createdAt: data[i][7]
      };

      if (filter && filter.search) {
        var search = filter.search.toLowerCase();
        var matches = Object.values(balance).some(function(val) {
          return String(val).toLowerCase().includes(search);
        });
        if (!matches) continue;
      }

      balances.push(balance);
    }

    return createResponse(true, { balances: balances });
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function getDashboardData() {
  try {
    var clients = getClients().clients || [];
    var projects = getProjects().projects || [];
    var incomeList = getIncome().incomeList || [];
    var expensesList = getExpenses().expenses || [];
    var loansList = getLoans().loans || [];
    var balancesList = getBalances().balances || [];

    var totalClients = clients.length;
    var totalProjects = projects.length;
    var activeProjects = projects.filter(function(p) { return p.status === 'Running'; }).length;
    var completedProjects = projects.filter(function(p) { return p.status === 'Completed'; }).length;

    var totalIncome = incomeList.reduce(function(sum, i) { return sum + (parseFloat(i.amount) || 0); }, 0);
    var totalExpense = expensesList.reduce(function(sum, e) { return sum + (parseFloat(e.amount) || 0); }, 0);
    var netProfit = totalIncome - totalExpense;

    var totalLoans = loansList.reduce(function(sum, l) { return sum + (parseFloat(l.amount) || 0); }, 0);

    var totalBalance = 0;
    balancesList.forEach(function(b) {
      var amt = parseFloat(b.amount) || 0;
      if (b.type === 'deposit') totalBalance += amt;
      else totalBalance -= amt;
    });

    var pendingPayments = incomeList.filter(function(i) {
      var parts = i.date.split('/');
      if (parts.length !== 3) return false;
      var incDate = new Date(parts[2], parts[1] - 1, parts[0]);
      var now = new Date();
      return (now - incDate) > 30 * 24 * 60 * 60 * 1000;
    }).length;

    var monthlyData = {};
    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    months.forEach(function(m) { monthlyData[m] = { income: 0, expense: 0, profit: 0 }; });

    incomeList.forEach(function(i) {
      if (i.date) {
        var parts = i.date.split('/');
        if (parts.length === 3) {
          var monthIndex = parseInt(parts[1]) - 1;
          if (monthIndex >= 0 && monthIndex < 12) {
            monthlyData[months[monthIndex]].income += parseFloat(i.amount) || 0;
          }
        }
      }
    });

    expensesList.forEach(function(e) {
      if (e.date) {
        var parts = e.date.split('/');
        if (parts.length === 3) {
          var monthIndex = parseInt(parts[1]) - 1;
          if (monthIndex >= 0 && monthIndex < 12) {
            monthlyData[months[monthIndex]].expense += parseFloat(e.amount) || 0;
          }
        }
      }
    });

    Object.keys(monthlyData).forEach(function(m) {
      monthlyData[m].profit = monthlyData[m].income - monthlyData[m].expense;
    });

    var revenueByProject = {};
    incomeList.forEach(function(i) {
      if (i.projectName) {
        revenueByProject[i.projectName] = (revenueByProject[i.projectName] || 0) + (parseFloat(i.amount) || 0);
      }
    });

    var expenseByCategory = {};
    expensesList.forEach(function(e) {
      if (e.category) {
        expenseByCategory[e.category] = (expenseByCategory[e.category] || 0) + (parseFloat(e.amount) || 0);
      }
    });

    return createResponse(true, {
      data: {
        totalClients: totalClients,
        totalProjects: totalProjects,
        activeProjects: activeProjects,
        completedProjects: completedProjects,
        totalIncome: totalIncome,
        totalExpense: totalExpense,
        netProfit: netProfit,
        pendingPayments: pendingPayments,
        totalLoans: totalLoans,
        totalBalance: totalBalance,
        monthlyData: monthlyData,
        revenueByProject: revenueByProject,
        expenseByCategory: expenseByCategory,
        recentIncome: incomeList.slice(-5).reverse(),
        recentExpenses: expensesList.slice(-5).reverse(),
        recentLoans: loansList.slice(-5).reverse().map(function(l) { return { date: l.date, lender: l.lender, amount: l.amount }; }),
        recentBalances: balancesList.slice(-5).reverse().map(function(b) { return { date: b.date, accountType: b.accountType, description: b.description, amount: b.amount, type: b.type }; })
      }
    });
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function logActivity(action, details) {
  try {
    var sheet = getSheet(CONFIG.SHEET_NAMES.ACTIVITIES);
    var id = getNextId(CONFIG.SHEET_NAMES.ACTIVITIES);
    var now = getCurrentDateString();
    var username = 'Unknown';

    try {
      var sessionUser = PropertiesService.getUserProperties().getProperty('session_user');
      if (sessionUser) username = JSON.parse(sessionUser).username;
    } catch (e) {}

    sheet.appendRow([id, now, username, action, details, '']);
    return true;
  } catch (e) {
    console.error('Failed to log activity: ' + e.toString());
    return false;
  }
}

function getActivities(filter) {
  try {
    var sheet = getSheet(CONFIG.SHEET_NAMES.ACTIVITIES);
    var data = sheet.getDataRange().getValues();
    var activities = [];

    for (var i = 1; i < data.length; i++) {
      var activity = {
        id: data[i][0],
        timestamp: data[i][1],
        user: data[i][2],
        action: data[i][3],
        details: data[i][4]
      };

      if (filter && filter.search) {
        var search = filter.search.toLowerCase();
        var matches = Object.values(activity).some(function(val) {
          return String(val).toLowerCase().includes(search);
        });
        if (!matches) continue;
      }

      activities.push(activity);
    }

    return createResponse(true, { activities: activities.reverse() });
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function getReport(params) {
  try {
    var type = typeof params === 'object' ? params.type : params;
    var format = typeof params === 'object' ? (params.format || 'json') : 'json';
    var filters = typeof params === 'object' ? (params.filters || {}) : {};
    var reportData = [];
    var headers = [];
    var filename = '';

    switch (type) {
      case 'daily':
        headers = ['Date', 'Income', 'Expense', 'Profit'];
        var today = formatDateFromISO(getCurrentDateString());
        var dailyIncome = getIncome({ fromDate: today, toDate: today }).incomeList || [];
        var dailyExpense = getExpenses({ fromDate: today, toDate: today }).expenses || [];
        reportData.push([today,
          dailyIncome.reduce(function(s, i) { return s + parseFloat(i.amount); }, 0),
          dailyExpense.reduce(function(s, e) { return s + parseFloat(e.amount); }, 0),
          dailyIncome.reduce(function(s, i) { return s + parseFloat(i.amount); }, 0) - dailyExpense.reduce(function(s, e) { return s + parseFloat(e.amount); }, 0)
        ]);
        filename = 'Daily_Report';
        break;

      case 'monthly':
        headers = ['Month', 'Income', 'Expense', 'Profit'];
        var monthlyIncome = getIncome(filters).incomeList || [];
        var monthlyExpense = getExpenses(filters).expenses || [];
        var monthGroups = {};
        monthlyIncome.forEach(function(i) {
          var key = i.date ? i.date.substring(3) : 'Unknown';
          if (!monthGroups[key]) monthGroups[key] = { income: 0, expense: 0 };
          monthGroups[key].income += parseFloat(i.amount) || 0;
        });
        monthlyExpense.forEach(function(e) {
          var key = e.date ? e.date.substring(3) : 'Unknown';
          if (!monthGroups[key]) monthGroups[key] = { income: 0, expense: 0 };
          monthGroups[key].expense += parseFloat(e.amount) || 0;
        });
        Object.keys(monthGroups).forEach(function(key) {
          reportData.push([key, monthGroups[key].income, monthGroups[key].expense, monthGroups[key].income - monthGroups[key].expense]);
        });
        filename = 'Monthly_Report';
        break;

      case 'yearly':
        headers = ['Year', 'Income', 'Expense', 'Profit'];
        var yearlyIncome = getIncome(filters).incomeList || [];
        var yearlyExpense = getExpenses(filters).expenses || [];
        var yearGroups = {};
        yearlyIncome.forEach(function(i) {
          var key = i.date ? i.date.substring(6) : 'Unknown';
          if (!yearGroups[key]) yearGroups[key] = { income: 0, expense: 0 };
          yearGroups[key].income += parseFloat(i.amount) || 0;
        });
        yearlyExpense.forEach(function(e) {
          var key = e.date ? e.date.substring(6) : 'Unknown';
          if (!yearGroups[key]) yearGroups[key] = { income: 0, expense: 0 };
          yearGroups[key].expense += parseFloat(e.amount) || 0;
        });
        Object.keys(yearGroups).forEach(function(key) {
          reportData.push([key, yearGroups[key].income, yearGroups[key].expense, yearGroups[key].income - yearGroups[key].expense]);
        });
        filename = 'Yearly_Report';
        break;

      case 'project':
        headers = ['Project', 'Total Income', 'Total Expense', 'Profit'];
        var projectIncome = getIncome(filters).incomeList || [];
        var projectExpense = getExpenses(filters).expenses || [];
        var projectGroups = {};
        projectIncome.forEach(function(i) {
          var key = i.projectName || 'Unknown';
          if (!projectGroups[key]) projectGroups[key] = { income: 0, expense: 0 };
          projectGroups[key].income += parseFloat(i.amount) || 0;
        });
        projectExpense.forEach(function(e) {
          var key = e.projectName || 'Unknown';
          if (!projectGroups[key]) projectGroups[key] = { income: 0, expense: 0 };
          projectGroups[key].expense += parseFloat(e.amount) || 0;
        });
        Object.keys(projectGroups).forEach(function(key) {
          reportData.push([key, projectGroups[key].income, projectGroups[key].expense, projectGroups[key].income - projectGroups[key].expense]);
        });
        filename = 'Project_Wise_Report';
        break;

      case 'client':
        headers = ['Client', 'Total Income', 'Total Expense', 'Profit'];
        var clientIncome = getIncome(filters).incomeList || [];
        var clientExpense = getExpenses(filters).expenses || [];
        var clientGroups = {};
        clientIncome.forEach(function(i) {
          var key = i.clientName || 'Unknown';
          if (!clientGroups[key]) clientGroups[key] = { income: 0, expense: 0 };
          clientGroups[key].income += parseFloat(i.amount) || 0;
        });
        clientExpense.forEach(function(e) {
          var key = e.projectName || 'Unknown';
          var clientMatch = clientIncome.find(function(ci) { return ci.projectName === e.projectName; });
          var clientKey = clientMatch ? clientMatch.clientName : 'Unknown';
          if (!clientGroups[clientKey]) clientGroups[clientKey] = { income: 0, expense: 0 };
          clientGroups[clientKey].expense += parseFloat(e.amount) || 0;
        });
        Object.keys(clientGroups).forEach(function(key) {
          reportData.push([key, clientGroups[key].income, clientGroups[key].expense, clientGroups[key].income - clientGroups[key].expense]);
        });
        filename = 'Client_Wise_Report';
        break;
    }

    if (format === 'csv') {
      return createResponse(true, { csv: generateCSV(headers, reportData), filename: filename });
    } else if (format === 'pdf') {
      var html = generateReportHTML(type, headers, reportData, filename);
      var blob = Utilities.newBlob(html, 'text/html', filename + '.html');
      var file = DriveApp.createFile(blob);
      return createResponse(true, { url: file.getUrl(), html: html });
    } else {
      return createResponse(true, { data: reportData, headers: headers, filename: filename });
    }
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function generateCSV(headers, data) {
  var csvRows = [headers.join(',')];
  data.forEach(function(row) {
    csvRows.push(row.map(function(val) { return '"' + String(val).replace(/"/g, '""') + '"'; }).join(','));
  });
  return csvRows.join('\n');
}

function generateReportHTML(type, headers, data, title) {
  var rows = data.map(function(row) {
    return '<tr>' + row.map(function(val) { return '<td>' + val + '</td>'; }).join('') + '</tr>';
  }).join('');

  return '' +
    '<html><head><style>' +
    'body { font-family: Arial, sans-serif; padding: 40px; }' +
    'h1 { color: #1a1a2e; border-bottom: 2px solid #1a1a2e; padding-bottom: 10px; }' +
    'table { width: 100%; border-collapse: collapse; margin-top: 20px; }' +
    'th { background: #1a1a2e; color: #fff; padding: 12px; text-align: left; }' +
    'td { padding: 10px; border-bottom: 1px solid #ddd; }' +
    'tr:hover { background: #f5f5f5; }' +
    '.footer { margin-top: 30px; color: #666; font-size: 12px; text-align: center; }' +
    '</style></head><body>' +
    '<h1>' + CONFIG.COMPANY + ' - ' + title + '</h1>' +
    '<p>Generated: ' + new Date().toLocaleString() + '</p>' +
    '<table><tr>' + headers.map(function(h) { return '<th>' + h + '</th>'; }).join('') + '</tr>' + rows + '</table>' +
    '<div class="footer">' + CONFIG.APP_NAME + ' | Automated Report</div>' +
    '</body></html>';
}

function exportCSV(params) {
  var type = typeof params === 'object' ? params.type : params;
  var filters = typeof params === 'object' ? (params.filters || {}) : {};
  return getReport({ type: type, format: 'csv', filters: filters });
}

function exportPDF(params) {
  var type = typeof params === 'object' ? params.type : params;
  var filters = typeof params === 'object' ? (params.filters || {}) : {};
  return getReport({ type: type, format: 'pdf', filters: filters });
}

function getSettings() {
  try {
    var sheet = getSheet(CONFIG.SHEET_NAMES.SETTINGS);
    var data = sheet.getDataRange().getValues();
    var settings = {};

    for (var i = 1; i < data.length; i++) {
      settings[data[i][0]] = data[i][1];
    }

    return createResponse(true, { settings: settings });
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function updateSettings(data) {
  try {
    var sheet = getSheet(CONFIG.SHEET_NAMES.SETTINGS);
    var now = getCurrentDateString();
    var existingData = sheet.getDataRange().getValues();
    var existingKeys = {};

    for (var i = 1; i < existingData.length; i++) {
      existingKeys[existingData[i][0]] = i;
    }

    Object.keys(data).forEach(function(key) {
      if (existingKeys[key] !== undefined) {
        sheet.getRange(existingKeys[key] + 1, 2).setValue(data[key]);
        sheet.getRange(existingKeys[key] + 1, 3).setValue(now);
      } else {
        sheet.appendRow([key, data[key], now]);
      }
    });

    logActivity('Settings Update', 'System settings updated');
    return createResponse(true, {}, 'Settings updated successfully');
  } catch (error) {
    return createResponse(false, {}, error.toString());
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function formatDateFromISO(isoStr) {
  if (!isoStr) return '';
  var d = new Date(isoStr);
  var day = String(d.getDate()).padStart(2, '0');
  var month = String(d.getMonth() + 1).padStart(2, '0');
  var year = d.getFullYear();
  return day + '/' + month + '/' + year;
}

function numberToWords(num) {
  if (num === 0) return 'Zero';

  var ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine',
    'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  var tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

  function convertLessThanThousand(n) {
    if (n === 0) return '';
    if (n < 20) return ones[n];
    if (n < 100) return tens[Math.floor(n / 10)] + (n % 10 ? ' ' + ones[n % 10] : '');
    return ones[Math.floor(n / 100)] + ' Hundred' + (n % 100 ? ' ' + convertLessThanThousand(n % 100) : '');
  }

  function convert(n) {
    if (n === 0) return '';
    var crore = Math.floor(n / 10000000);
    var lakh = Math.floor((n % 10000000) / 100000);
    var thousand = Math.floor((n % 100000) / 1000);
    var hundred = Math.floor((n % 1000) / 100);
    var remainder = n % 100;

    var result = '';
    if (crore) result += convertLessThanThousand(crore) + ' Crore ';
    if (lakh) result += convertLessThanThousand(lakh) + ' Lakh ';
    if (thousand) result += convertLessThanThousand(thousand) + ' Thousand ';
    if (hundred) result += convertLessThanThousand(hundred) + ' Hundred ';
    if (remainder) {
      if (result) result += 'and ';
      result += convertLessThanThousand(remainder);
    }
    return result.trim();
  }

  var rupees = Math.floor(num);
  var paise = Math.round((num - rupees) * 100);

  var words = convert(rupees);
  if (words) words += ' Rupees';
  if (paise > 0) words += ' and ' + convert(paise) + ' Paise';
  words += ' Only';

  return words;
}

function onOpen() {
  initializeSheets();
}

function setup() {
  initializeSheets();
  return createResponse(true, {}, 'System initialized successfully');
}

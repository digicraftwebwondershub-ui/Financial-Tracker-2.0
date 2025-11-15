// --- Global Utility Functions ---
const SS = SpreadsheetApp.getActiveSpreadsheet();

/**
 * Serves the main web app HTML template.
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  
  // Pass the user's name for a personalized touch
  const userEmail = Session.getActiveUser().getEmail();
  template.userName = userEmail ? userEmail.split('@')[0] : 'User';

  return template.evaluate()
    .setTitle('Fin-Track: Personal Finance Tracker')
    .setFaviconUrl('https://ssl.gstatic.com/docs/spreadsheets/forms/favicon_48.png') // Generic Sheet Icon
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Includes the content of a separate HTML file (e.g., Sidebar.html, Style.html, Script.html) 
 * into the main template using a "force printing" syntax in Index.html: <?!= include('Filename') ?>.
 * @param {string} filename The name of the HTML file.
 * @return {string} The evaluated HTML content.
 */
function include(filename) {
  // Use createHtmlOutputFromFile().getContent() to get the raw content, 
  // which is then inserted unfiltered via <?!= ?> in the main template.
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Gets a specific sheet by name.
 * @param {string} sheetName The name of the sheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The sheet object.
 */
function getSheet(sheetName) {
  return SS.getSheetByName(sheetName);
}

/**
 * Creates a custom menu in the spreadsheet UI when the sheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Fin-Track Tools')
      .addItem('Recalculate Card Balances', 'recalculateAndUpdateCardBalances')
      .addToUi();
}

// --- ID & Config Functions ---

/**
 * Sets up initial ID counters in the Config sheet. Run once after setup.
 */
function setupInitialConfig() {
  const sheetUi = SpreadsheetApp.getUi(); 
  const configSheet = getSheet('Config');
  if (!configSheet) {
    throw new Error("Config sheet not found. Please create it.");
  }
  
  configSheet.clearContents();
  configSheet.getRange('A1:B1').setValues([['KEY', 'VALUE']]);
  
  configSheet.appendRow(['NEXT_TRANSACTION_ID', 1000]);
  configSheet.appendRow(['NEXT_CARD_ID', 2000]);
  configSheet.appendRow(['NEXT_GOAL_ID', 3000]);
  configSheet.appendRow(['NEXT_REMINDER_ID', 4000]);
  
  sheetUi.alert('Initial configuration complete! Run fixConfigIDs() next to align IDs with existing data.'); 
}

/**
 * FIX FUNCTION: Aligns the NEXT_ID values in the Config sheet 
 * based on the last row ID found in each data sheet.
 */
function fixConfigIDs() {
    const sheetUi = SpreadsheetApp.getUi();
    const configSheet = getSheet('Config');
    
    const sheetMaps = {
        'Transactions': { key: 'NEXT_TRANSACTION_ID', prefix: 'TR-' },
        'Credit_Cards': { key: 'NEXT_CARD_ID', prefix: 'CARD-' },
        'Goals': { key: 'NEXT_GOAL_ID', prefix: 'GOAL-' },
        'Reminders': { key: 'NEXT_REMINDER_ID', prefix: 'REM-' }
    };
    
    for (let sheetName in sheetMaps) {
        const sheet = getSheet(sheetName);
        if (sheet.getLastRow() > 1) {
            const lastIdString = sheet.getRange(sheet.getLastRow(), 1).getValue();
            let lastIdNumber = parseInt(String(lastIdString).replace(sheetMaps[sheetName].prefix, '')) || 0;
            
            if (lastIdNumber > 0) {
                const nextId = lastIdNumber + 1;
                
                const configValues = configSheet.getDataRange().getValues();
                for (let i = 0; i < configValues.length; i++) {
                    if (configValues[i][0] === sheetMaps[sheetName].key) {
                        configSheet.getRange(i + 1, 2).setValue(nextId);
                        break;
                    }
                }
            }
        }
    }
    
    sheetUi.alert('Configuration IDs successfully aligned with existing data!');
}

/**
 * Generates a unique ID for a given prefix (TR, CARD, GOAL, REM).
 * @param {string} prefix The ID prefix.
 * @return {string} The new unique ID.
 */
function generateUniqueId(prefix) {
  const configSheet = getSheet('Config');
  const keyMap = {
    'TR': 'NEXT_TRANSACTION_ID',
    'CARD': 'NEXT_CARD_ID',
    'GOAL': 'NEXT_GOAL_ID',
    'REM': 'NEXT_REMINDER_ID'
  };
  const key = keyMap[prefix];
  
  if (!key) throw new Error("Invalid ID prefix.");
  
  const range = configSheet.getRange('A:A');
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === key) {
      let nextId = parseInt(configSheet.getRange(i + 1, 2).getValue());
      configSheet.getRange(i + 1, 2).setValue(nextId + 1); // Increment for next time
      return `${prefix}-${nextId}`;
    }
  }
  throw new Error(`Configuration key ${key} not found.`);
}

// --- Data Loading Functions ---

/**
 * Loads all data from a specified sheet.
 * @param {string} sheetName The name of the sheet.
 * @return {Array<Object>} An array of objects, where each object is a row with header keys.
 */
function loadData(sheetName) {
  Logger.log(`Starting data load for sheet: ${sheetName}`);
  const sheet = getSheet(sheetName);
  if (!sheet) {
      Logger.log(`Error: Sheet ${sheetName} not found.`);
      return [];
  }
  
  // Safely handle sheets with only a header row or empty sheets.
  if (sheet.getLastRow() < 2) {
    Logger.log(`Sheet ${sheetName} is empty.`);
    return []; 
  }
  
  const range = sheet.getDataRange();
  const [headers, ...data] = range.getValues();
  
  // Define keys that represent numeric values and need cleanup/conversion from string
  const NUMERIC_KEYS = ['AMOUNT', 'LIMIT', 'BALANCE', 'LASTPAYMENT', 'APR', 'TARGETAMOUNT', 'SAVEDAMOUNT', 'MONTHLYSAVINGS', 'DAYSLEFT'];
  const DATE_KEYS = ['DATE', 'LASTPAYMENTDATE', 'DUEDATE', 'STATEMENTDATE', 'TARGETDATE'];

  let results = [];
  data.forEach((row, rowIndex) => {
    let obj = {};
    headers.forEach((header, i) => {
      // Clean header name: AMOUNT(₱) becomes AMOUNT
      const key = header.replace(/[^a-zA-Z0-9_]/g, '').toUpperCase();
      let value = row[i];
      
      try {
          // 1. Numerical Cleanup
          if (NUMERIC_KEYS.includes(key)) {
              if (typeof value === 'string') {
                  // Remove commas/thousand separators and try to parse
                  value = parseFloat(value.replace(/,/g, '')) || 0;
              } else if (value instanceof Date) {
                  // Sometimes numerical data is read as a date (e.g., small numbers)
                  value = parseFloat(value.getTime()) || 0;
              } else {
                  value = parseFloat(value) || 0;
              }
          }
          
          // 2. Date Cleanup (Convert date objects to string for client-side consumption)
          if (DATE_KEYS.includes(key) && value instanceof Date) {
              value = value.toISOString().split('T')[0];
          }

          obj[key] = value;
      } catch (e) {
          Logger.log(`Error processing row ${rowIndex + 2} in ${sheetName}: Key=${key}, Value=${value}. Error: ${e.message}`);
          obj[key] = null; // Set to null on error to prevent total script crash
      }
    });
    results.push(obj);
  });
  
  Logger.log(`Successfully loaded ${results.length} rows from ${sheetName}.`);
  return results;
}

/**
 * Loads all data for a specific type of tab.
 */
function loadTransactions() { return loadData('Transactions'); }

function loadCreditCards() {
  const cards = loadData('Credit_Cards');
  const transactions = loadData('Transactions');

  // Calculate live balances from the transaction log
  const cardBalances = {};
  transactions.forEach(t => {
    const cardId = t.ACCOUNT;
    if (cardId && cardId.startsWith('CARD-')) {
      if (!cardBalances[cardId]) {
        cardBalances[cardId] = 0;
      }
      const amount = t.AMOUNT || 0;
      if (t.CATEGORY === 'Credit Card Payment') {
        cardBalances[cardId] -= amount;
      } else if (t.TYPE === 'Expense') {
        cardBalances[cardId] += amount;
      }
    }
  });

  // Map the live balances back onto the card data, replacing the static balance
  const updatedCards = cards.map(card => {
    const liveBalance = cardBalances[card.CARD_ID];
    if (liveBalance !== undefined) {
      card.BALANCE = liveBalance;
    }
    return card;
  });

  return updatedCards;
}

function loadGoals() { return loadData('Goals'); }
function loadReminders() { return loadData('Reminders'); }

// --- Dashboard Calculation Functions ---

/**
 * Calculates all dashboard key performance indicators (KPIs).
 */
function getDashboardData() {
  try {
    const transactions = loadTransactions();
    const cards = loadCreditCards();
    const goals = loadGoals();

    let totalIncome = 0;
    let totalExpenses = 0;
    let totalSavingsDeposits = 0;
    let totalCreditLimit = 0;
    let totalCardBalance = 0;

    // 1. Transaction-based calculations
    transactions.forEach(t => {
      // t.AMOUNT is now guaranteed to be a number due to loadData fix
      const amount = t.AMOUNT || 0;
      if (t.TYPE === 'Income') {
        totalIncome += amount;
      } else if (t.TYPE === 'Expense') {
        totalExpenses += amount;
      }
      
      if (t.CATEGORY === 'Savings Deposit') {
        totalSavingsDeposits += amount;
      }
    });

    // 2. Credit Card calculations
    // The `cards` object from `loadCreditCards()` now contains live balances.
    const cardStats = cards.map(c => {
      const limit = c.LIMIT || 0;
      const balance = c.BALANCE || 0;
      totalCreditLimit += limit;
      totalCardBalance += balance;
      
      return {
        name: c.CARDNAME,
        balance: balance,
        limit: limit,
        usage: limit > 0 ? (balance / limit) : 0,
        id: c.CARD_ID
      };
    });
    
    // 3. Goals calculations
    const goalsProgress = goals.map(g => {
      const target = g.TARGETAMOUNT || 0;
      const saved = g.SAVEDAMOUNT || 0;
      const progress = target > 0 ? (saved / target) : 0;
      
      return {
        name: g.GOALNAME,
        saved: saved,
        target: target,
        progress: progress,
        priority: g.PRIORITYLEVEL
      };
    });

    const netIncome = totalIncome - totalExpenses;
    const savingsRate = totalIncome > 0 ? (totalSavingsDeposits / totalIncome) : 0;
    const creditUsage = totalCreditLimit > 0 ? (totalCardBalance / totalCreditLimit) : 0;
    const availableCredit = totalCreditLimit - totalCardBalance;

    Logger.log('Dashboard data calculated successfully.');
    return {
      netIncome: netIncome,
      totalExpenses: totalExpenses,
      totalIncome: totalIncome,
      savingsRate: savingsRate,
      creditUsage: creditUsage,
      availableCredit: availableCredit,
      cardStats: cardStats,
      goalsProgress: goalsProgress,
      motivationalMessage: netIncome >= 0 ? 
        "Great job! Your net income is positive—keep building financial freedom!" : 
        "Review expenses this month. Every small saving counts!"
    };
  } catch (e) {
    Logger.log(`FATAL ERROR in getDashboardData: ${e.message}`);
    // Return empty data structure to prevent front-end crash
    return {
      netIncome: 0, totalExpenses: 0, totalIncome: 0, savingsRate: 0, creditUsage: 0, availableCredit: 0, cardStats: [], goalsProgress: [], motivationalMessage: "Data failed to load. Check Apps Script logs."
    };
  }
}

/**
 * RECALCULATION FUNCTION: Iterates through all transactions to
 * accurately update each credit card's balance and last payment info.
 */
function recalculateAndUpdateCardBalances() {
  const sheetUi = SpreadsheetApp.getUi();
  sheetUi.alert('Starting card balance recalculation... This may take a moment.');

  try {
    const transactions = loadData('Transactions');
    const cards = loadData('Credit_Cards');
    const cardsSheet = getSheet('Credit_Cards');

    if (!cardsSheet || cards.length === 0) {
      sheetUi.alert('No credit cards found to update.');
      return;
    }

    // Create a map for quick lookups
    const cardUpdates = {};
    cards.forEach(card => {
      cardUpdates[card.CARD_ID] = {
        balance: 0,
        lastPayment: 0,
        lastPaymentDate: null,
        rowNum: -1 // We'll need to find this later
      };
    });

    // 1. Process all transactions
    transactions.forEach(t => {
      const cardId = t.ACCOUNT;
      if (cardUpdates[cardId]) {
        const amount = t.AMOUNT || 0;

        if (t.TYPE === 'Expense') {
          cardUpdates[cardId].balance += amount;
        } else if (t.CATEGORY === 'Credit Card Payment') {
          // Payments should also be treated as expenses against the card balance from a transactional view
          cardUpdates[cardId].balance -= amount; // Payment reduces the balance

          // Check if this is the most recent payment
          const paymentDate = new Date(t.DATE);
          if (!cardUpdates[cardId].lastPaymentDate || paymentDate > cardUpdates[cardId].lastPaymentDate) {
            cardUpdates[cardId].lastPayment = amount;
            cardUpdates[cardId].lastPaymentDate = paymentDate;
          }
        }
      }
    });

    // 2. Prepare data for batch update
    const [headers, ...rows] = cardsSheet.getDataRange().getValues();
    const balanceColIndex = headers.indexOf('BALANCE');
    const lastPaymentColIndex = headers.indexOf('LASTPAYMENT');
    const lastPaymentDateColIndex = headers.indexOf('LASTPAYMENTDATE');

    if (balanceColIndex === -1 || lastPaymentColIndex === -1 || lastPaymentDateColIndex === -1) {
        throw new Error('Could not find required columns (BALANCE, LASTPAYMENT, LASTPAYMENTDATE) in Credit_Cards sheet.');
    }

    const updateRangeData = [];
    rows.forEach((row, i) => {
        const cardId = row[0];
        if (cardUpdates[cardId]) {
            const update = cardUpdates[cardId];
            row[balanceColIndex] = update.balance;
            row[lastPaymentColIndex] = update.lastPayment;
            row[lastPaymentDateColIndex] = update.lastPaymentDate ? update.lastPaymentDate.toLocaleDateString('en-US') : '';
        }
        updateRangeData.push(row);
    });

    // 3. Perform the batch update
    const dataRange = cardsSheet.getRange(2, 1, updateRangeData.length, headers.length);
    dataRange.setValues(updateRangeData);

    sheetUi.alert('Success! All credit card balances and last payment details have been recalculated and updated.');
  } catch (e) {
    Logger.log(`ERROR during recalculation: ${e.toString()}`);
    sheetUi.alert(`An error occurred: ${e.message}. Please check the logs.`);
  }
}

// --- Action & Update Functions ---

/**
 * Finds and updates a row based on its unique ID.
 * @param {string} sheetName The name of the sheet.
 * @param {string} id The unique ID to match.
 * @param {Object} data The data object with updated values.
 * @return {boolean} True if update was successful.
 */
function updateRecordById(sheetName, id, data) {
  const sheet = getSheet(sheetName);
  const [headers, ...rows] = sheet.getDataRange().getValues();
  const idColIndex = 0; // The first column is always the ID
  
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][idColIndex] === id) {
      const rowNum = i + 2; // +1 for 0-index, +1 for header row
      
      headers.forEach((header, colIndex) => {
        const key = header.replace(/[^a-zA-Z0-9_]/g, '').toUpperCase();
        if (data.hasOwnProperty(key)) {
          // Special handling for numerical fields
          let value = data[key];
          if (['AMOUNT', 'LIMIT', 'BALANCE', 'TARGETAMOUNT', 'SAVEDAMOUNT', 'MONTHLYSAVINGS', 'APR'].includes(key)) {
            // Ensure any string input is cleaned/parsed again before saving back to sheet
            if (typeof value === 'string') {
                value = parseFloat(value.replace(/,/g, '')) || 0;
            } else {
                value = parseFloat(value) || 0;
            }
          }
          sheet.getRange(rowNum, colIndex + 1).setValue(value);
        }
      });
      return true;
    }
  }
  return false;
}

/**
 * Adds a new transaction and handles the relational updates (A & B).
 * @param {Object} formData Transaction data.
 * @return {Object} Status and message.
 */
function addTransaction(formData) {
  try {
    const transactionId = generateUniqueId('TR');
    const sheet = getSheet('Transactions');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = [];

    // Map form data to column order
    const dataMap = {
      'TRANSACTION_ID': transactionId,
      'DATE': formData.DATE || new Date().toLocaleDateString('en-US'),
      'TYPE': formData.TYPE,
      'CATEGORY': formData.CATEGORY,
      // Clean amount before processing and saving
      'AMOUNT(₱)': parseFloat(String(formData.AMOUNT).replace(/,/g, '')) || 0, 
      'DESCRIPTION': formData.DESCRIPTION,
      'PAYMENTMETHOD': formData.PAYMENTMETHOD,
      'ACCOUNT': formData.ACCOUNT, // Could be a CARD_ID
      'RELATED_ID': formData.RELATED_ID // Could be a GOAL_ID
    };
    
    headers.forEach(header => {
      const key = header.replace(/[^a-zA-Z0-9_]/g, '').toUpperCase();
      newRow.push(dataMap[key]);
    });

    sheet.appendRow(newRow);

    // --- Relational Logic A: Credit Card Transactions ---
    if (formData.PAYMENTMETHOD === 'Credit Card' && formData.ACCOUNT) {
      updateCreditCardBalance(formData.ACCOUNT, parseFloat(String(formData.AMOUNT).replace(/,/g, '')) || 0, formData.TYPE, formData.CATEGORY);
    }

    // --- Relational Logic B: Savings Goal Deposits ---
    if (formData.CATEGORY === 'Savings Deposit' && formData.RELATED_ID) {
      updateSavingsGoal(formData.RELATED_ID, parseFloat(String(formData.AMOUNT).replace(/,/g, '')) || 0);
    }

    return { status: 'success', message: `Transaction ${transactionId} added successfully.` };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * Updates a credit card's balance based on a transaction. (Logic A)
 */
function updateCreditCardBalance(cardId, amount, type, category) {
  const cards = loadCreditCards();
  const card = cards.find(c => c.CARD_ID === cardId);

  if (card) {
    let newBalance = card.BALANCE;
    let updateData = {};
    
    // Expenses (including payment) increase the balance, Income reduces it
    if (type === 'Expense') {
      newBalance += amount;
    } else if (type === 'Income') {
      // This case is unlikely for cards, but included for completeness
      newBalance -= amount; 
    }
    
    updateData.BALANCE = newBalance;
    
    // Update Last Payment info if it's a Credit Card Payment
    if (category === 'Credit Card Payment') {
      updateData.LASTPAYMENT = amount;
      updateData.LASTPAYMENTDATE = new Date().toLocaleDateString('en-US');
    }

    // Update the record in the sheet
    updateRecordById('Credit_Cards', cardId, updateData);
  }
}

/**
 * Increases the saved amount for a goal. (Logic B)
 */
function updateSavingsGoal(goalId, depositAmount) {
  const goals = loadGoals();
  const goal = goals.find(g => g.GOAL_ID === goalId);
  
  if (goal) {
    const newSavedAmount = (goal.SAVEDAMOUNT || 0) + depositAmount;
    updateRecordById('Goals', goalId, { SAVEDAMOUNT: newSavedAmount });
  }
}

/**
 * Creates a transaction from a paid reminder and resets the reminder. (Logic C)
 * @param {string} reminderId The ID of the reminder marked as paid.
 * @param {Object} formData Reminder data.
 * @return {Object} Status and message.
 */
function markReminderPaid(reminderId, formData) {
  try {
    const reminder = loadReminders().find(r => r.REMINDER_ID === reminderId);
    if (!reminder) {
      return { status: 'error', message: 'Reminder not found.' };
    }

    // 1. Automatically create a new transaction (Expense)
    const transactionData = {
      DATE: new Date().toLocaleDateString('en-US'),
      TYPE: 'Expense',
      CATEGORY: reminder.CATEGORY || 'Bill Payment',
      AMOUNT: reminder.AMOUNT || 0,
      DESCRIPTION: `Payment for Reminder: ${reminder.DESCRIPTION}`,
      PAYMENTMETHOD: reminder.PAYMENTCHANNEL || 'Other',
      ACCOUNT: '',
      RELATED_ID: reminderId
    };
    addTransaction(transactionData);

    // 2. Update reminder status
    let updateData = { STATUS: 'Paid' };
    
    // 3. Reset next due date if recurring
    if (reminder.RECURRING === 'Yes') {
      const currentDueDate = new Date(reminder.DUEDATE);
      let nextDueDate = new Date(currentDueDate);
      
      // Simple recurring logic: assumes monthly for a clean example
      nextDueDate.setMonth(nextDueDate.getMonth() + 1);
      updateData.DUEDATE = nextDueDate.toLocaleDateString('en-US');
    }
    
    updateRecordById('Reminders', reminderId, updateData);
    
    return { status: 'success', message: `Reminder ${reminderId} marked paid and transaction created.` };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * Generically adds a new record to any sheet.
 * @param {string} sheetName The name of the sheet.
 * @param {Object} formData The data to add.
 * @param {string} prefix The ID prefix (TR, CARD, GOAL, REM).
 * @return {Object} Status and message.
 */
function addRecord(sheetName, formData, prefix) {
  try {
    if (sheetName === 'Transactions') return addTransaction(formData); // Use the specialized one for logic
    
    const recordId = generateUniqueId(prefix);
    const sheet = getSheet(sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = [];
    
    headers.forEach((header, i) => {
      const key = header.replace(/[^a-zA-Z0-9_]/g, '').toUpperCase();
      let value = key === `${prefix}_ID` ? recordId : formData[key];
      
      // Handle the DAYLEFT calculation for Reminders (simple example)
      if (sheetName === 'Reminders' && key === 'DAYSLEFT' && formData.DUEDATE) {
        const dueDate = new Date(formData.DUEDATE);
        const today = new Date();
        const diffTime = dueDate.getTime() - today.getTime();
        value = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      }
      
      newRow.push(value);
    });
    
    sheet.appendRow(newRow);
    return { status: 'success', message: `${sheetName} record ${recordId} added successfully.` };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// --- Data Export Function (Extra Feature) ---
/**
 * Exports data from a sheet to a CSV file.
 * @param {string} sheetName The name of the sheet.
 * @return {string} A Base64 encoded string of the CSV data.
 */
function exportToCSV(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) throw new Error("Sheet not found.");
  
  const data = sheet.getDataRange().getValues();
  let csv = data.map(row => row.map(cell => {
    // Escape double quotes by doubling them, then wrap in quotes
    let processedCell = String(cell).replace(/"/g, '""');
    return `"${processedCell}"`;
  }).join(',')).join('\n');

  // Convert to Base64 to send to the frontend for download
  return Utilities.base64Encode(csv);
}

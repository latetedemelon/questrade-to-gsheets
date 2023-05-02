// Replace YOUR_API_KEY with your actual Questrade API key.
const YOUR_INITIAL_TOKEN = 'OcSYm8MZLj6dL_CZ9-aai6Hgwz4Dpv3g0';

// Set the sheet names you want to use for storing transactions and balances.
const TRANSACTIONS_SHEET_NAME = 'Transactions';
const BALANCES_SHEET_NAME = 'Balances';

function setup() {
  const accessToken = getAccessToken(YOUR_INITIAL_TOKEN);
  Logger.log('Access token:', accessToken);
}

function getAccessToken(refreshToken) {
  const url = 'https://login.questrade.com/oauth2/token?grant_type=refresh_token&refresh_token=' + refreshToken;
  const options = {
    method: 'GET',
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    const contentType = response.getHeaders()['Content-Type'];
    let errorMessage = 'Error getting access token (HTTP ' + response.getResponseCode() + '):\n';

    if (contentType && contentType.includes('application/json')) {
      const errorData = JSON.parse(response.getContentText());
      errorMessage += 'Error message: ' + errorData.message + '\n';
      errorMessage += 'Error code: ' + errorData.code + '\n';
    } else {
      errorMessage += response.getContentText();
    }

    Logger.log(errorMessage);
    throw new Error('Error getting access token. Check the logs for more details.');
  }

  const data = JSON.parse(response.getContentText());
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('access_token', data.access_token);
  scriptProperties.setProperty('refresh_token', data.refresh_token);
  scriptProperties.setProperty('api_server', data.api_server);

  return data.access_token;
}

function updateBalances() {
  const refreshToken = PropertiesService.getScriptProperties().getProperty('refresh_token');
  const apiServer = PropertiesService.getScriptProperties().getProperty('api_server');
  const accessToken = getAccessToken(refreshToken);
  const accounts = getAccounts(accessToken, apiServer);

  const balancesData = [];
  
  for (const account of accounts) {
    const balances = getBalances(accessToken, account.number, apiServer);
    for (const balance of balances) {
      balance.accountNumber = account.number; // Add account number to balance object
      balancesData.push(balance);
    }
  }

  writeToSheet(BALANCES_SHEET_NAME, balancesData);
}

function getAccounts(accessToken, apiServer) {
  const response = UrlFetchApp.fetch(apiServer + 'v1/accounts', {
    headers: {
      Authorization: 'Bearer ' + accessToken,
    },
  });
  
  const data = JSON.parse(response.getContentText());
  return data.accounts;
}

function getTransactions(accessToken, accountNumber, apiServer) {
  const startDate = new Date('2022-01-01');
  const endDate = new Date('2023-01-28');
  
  const oneDay = 24 * 60 * 60 * 1000;
  const maxInterval = 30 * oneDay;
  const allActivities = [];
  
  for (let currentStart = startDate; currentStart < endDate; currentStart = new Date(currentStart.getTime() + maxInterval)) {
    const currentEnd = new Date(Math.min(currentStart.getTime() + maxInterval, endDate.getTime()));
    
    const startTimeString = currentStart.toISOString().replace(/\.\d{3}Z/, 'Z');
    const endTimeString = currentEnd.toISOString().replace(/\.\d{3}Z/, 'Z');
    
    const response = UrlFetchApp.fetch(apiServer + 'v1/accounts/' + accountNumber + '/activities?startTime=' + startTimeString + '&endTime=' + endTimeString, {
      headers: {
        Authorization: 'Bearer ' + accessToken,
      },
    });

    const data = JSON.parse(response.getContentText());
    allActivities.push(...data.activities);
  }
  
  return allActivities;
}

function getBalances(accessToken, accountNumber, apiServer) {
  const response = UrlFetchApp.fetch(apiServer + 'v1/accounts/' + accountNumber + '/balances', {
    headers: {
      Authorization: 'Bearer ' + accessToken,
    },
  });
  
  const data = JSON.parse(response.getContentText());
  return data.perCurrencyBalances;
}

function importTransactions() {
  const refreshToken = PropertiesService.getScriptProperties().getProperty('refresh_token');
  const apiServer = PropertiesService.getScriptProperties().getProperty('api_server');
  const accessToken = getAccessToken(refreshToken);
  const accounts = getAccounts(accessToken, apiServer);
  
  // Clear the Transactions sheet before importing transactions for all accounts.
  const sheet = getSheet(TRANSACTIONS_SHEET_NAME);
  sheet.clearContents();
  
  for (const account of accounts) {
    const transactions = getTransactions(accessToken, account.number, apiServer);
    // Add the account number to each transaction.
    transactions.forEach(transaction => transaction.accountNumber = account.number);
    writeToSheet(TRANSACTIONS_SHEET_NAME, transactions, false);
  }
}

function writeToSheet(sheetName, data, clearContents = true) {
  if (!data || data.length === 0) {
    return;
  }
  
  const sheet = getSheet(sheetName);
  
  // Assumes the first object in the data array represents the structure for all objects.
  const headers = Object.keys(data[0]);
  
  // Add headers only if the sheet is empty.
  if (sheet.getLastRow() === 0) {
    // Clear the sheet's contents only if clearContents is true.
    if (clearContents) {
      sheet.clearContents();
    }
    
    sheet.appendRow(headers);
  }
  
  for (const item of data) {
    const row = [];
    for (const header of headers) {
      row.push(item[header]);
    }
    sheet.appendRow(row);
  }
}





function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  return sheet;
}

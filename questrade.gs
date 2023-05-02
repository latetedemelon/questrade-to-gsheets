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

function importTransactions() {
  const refreshToken = PropertiesService.getScriptProperties().getProperty('refresh_token');
  const apiServer = PropertiesService.getScriptProperties().getProperty('api_server');
  const accessToken = getAccessToken(refreshToken);
  const accounts = getAccounts(accessToken, apiServer);
  
  for (const account of accounts) {
    const transactions = getTransactions(accessToken, account.number, apiServer);
    writeToSheet(TRANSACTIONS_SHEET_NAME, transactions);
  }
}

function updateBalances() {
  const refreshToken = PropertiesService.getScriptProperties().getProperty('refresh_token');
  const apiServer = PropertiesService.getScriptProperties().getProperty('api_server');
  const accessToken = getAccessToken(refreshToken);
  const accounts = getAccounts(accessToken, apiServer);
  
  for (const account of accounts) {
    const balances = getBalances(accessToken, account.number, apiServer);
    writeToSheet(BALANCES_SHEET_NAME, balances);
  }
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
  // Set your desired date range for transactions.
  const startDate = '2022-01-01';
  const endDate = '2023-12-31';
  
  const response = UrlFetchApp.fetch(apiServer + 'v1/accounts/' + accountNumber + '/transactions?startTime=' + startDate + '&endTime=' + endDate, {
    headers: {
      Authorization: 'Bearer ' + accessToken,
    },
  });
  
  const data = JSON.parse(response.getContentText());
  return data.transactions;
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

function writeToSheet(sheetName, data) {
  const sheet = getSheet(sheetName);
  sheet.clearContents();

  // Assumes the first object in the data array represents the structure for all objects.
  const headers = Object.keys(data[0]);
  sheet.appendRow(headers);
  
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

[![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://paypal.me/latetedemelon) [![Donate](https://img.shields.io/badge/Donate-Buy%20Me%20a%20Coffee-yellow)](https://buymeacoffee.com/latetedemelon) [![Donate](https://img.shields.io/badge/Donate-Ko--Fi-ff69b4)](https://ko-fi.com/latetedemelon)
# questrade-to-gsheets

## Table of Contents

- [Features](#features)
- [Setup](#setup)
- [Usage](#usage)
- [Functions](#functions)

## Features

- Import account transactions
- Import account balances
- Import account positions
- Automatically refresh access tokens

## Setup

1. Open a new Google Sheet.
2. Click on `Extensions` > `Apps Script`.
3. Delete the default `Code.gs` file.
4. Click on `File` > `New` > `Script` and name the file `QuestradeIntegration`.
5. Copy the entire script from the provided code and paste it into the `QuestradeIntegration.gs` file.
6. Replace the `YOUR_INITIAL_TOKEN` placeholder with your actual Questrade API key.
7. Save the script by clicking on the floppy disk icon or pressing `Ctrl + S` (or `Cmd + S` on macOS).
8. Close the Apps Script editor.

## Usage

1. After setting up the script, go back to your Google Sheet and refresh the page via the browser.
2. Access the script editor by clicking on `Extensions` > `Apps Script`.
3. Click on `Select function` and choose the function you want to run, such as `importTransactions`, `updateBalances`, or `importPositions`.
4. Click on the play button to run the selected function.
5. Check your Google Sheet for the imported data in the respective sheets.

## Functions

- `setup`: Initializes the script and retrieves an access token using your API key.
- `getAccessToken`: Retrieves an access token using a refresh token.
- `updateBalances`: Imports account balances into the specified sheet.
- `getAccounts`: Retrieves a list of accounts associated with the API key.
- `getTransactions`: Retrieves account transactions for a specified date range.
- `importPositions`: Imports account positions into the specified sheet.
- `getPositions`: Retrieves account positions for a specified account number.
- `getBalances`: Retrieves account balances for a specified account number.
- `importTransactions`: Imports account transactions into the specified sheet.
- `writeToSheet`: Writes data to the specified sheet.
- `getSheet`: Retrieves or creates a sheet with the specified name.

## Notes

- Ensure you have a valid Questrade API key and replace the `YOUR_INITIAL_TOKEN` placeholder with your actual API key.
- You may need to adjust the date range in the `getTransactions` function to match your desired transaction history period.
- For a more customized experience, you can modify the script to suit your specific needs, such as changing the sheet names or adding new functionality.

## Contributing

Contributions to the Google Sheets Passiv Integration project are welcome! If you have improvements, bug fixes, or new features you'd like to see added, please submit a Pull Request.

## Donations

If you find this integration helpful and would like to support its development, consider making a donation. Every little bit helps!

<a href='https://paypal.me/latetedemelon' target='_blank'><img src="https://github.com/stefan-niedermann/paypal-donate-button/blob/master/paypal-donate-button.png" width="270" height="105" alt='Donate via PayPal' />

<a href='https://ko-fi.com/latetedemelon' target='_blank'><img height='35' style='border:0px;height:46px;' src='https://az743702.vo.msecnd.net/cdn/kofi3.png?v=0' border='0' alt='Buy Me a Coffee at ko-fi.com' />

[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/yellow_img.png)](https://www.buymeacoffee.com/latetedemelon)

## License

This project is licensed under the MIT License. See the [LICENSE](https://github.com/latetedemelon/passiv-to-gsheets/blob/main/LICENSE) file for details.

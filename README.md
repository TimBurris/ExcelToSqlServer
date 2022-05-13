# ExcelToSqlServer
The purpose of this utility is to simply provide a mechanism for getting data from excel into sql server for further manipulation.

It's just a console app with a bunch of settings in appSettings.json to control how/what gets brought in

# A Tutorial
I recorded a video tutorial outlining some use cases and reviewing the options and functionality here:
[Excel To Sql Server: a utility to simplify your life
](https://youtu.be/50DOu4CwZyg)

## Run
Execute the exe and either pass the excel file in as a command-line param or it will prompt you for the path to the excel file.
 All the other settings are controlled by appSettings.json

## Features
* Process all worksheets in the document, or just the first
* Explicitly ignore specific worksheets
* Explicitly include specifict worksheets
* Bring all worksheets into a single table, or separate tables per worksheet
* Will make you bagels in the morning (this may or may not be true)
* Optionally strips field names down to simply a-z0-9 characters to make them safe for sql server column names 
* Helps reduce duplicate columns which differ by spaces or special chars (e.g. Account Number vs Account_Number)

# Additional Settings
All the options and settings (e.g. SkipBlankRows) are documented in the appsettings.json

# Want to say thank you?
If something in this repo has helped you solve a problem, or made you more efficient, I welcome your support!

<a href="https://www.buymeacoffee.com/timburris" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/default-orange.png" alt="Buy Me A Coffee" height="41" width="174"></a>

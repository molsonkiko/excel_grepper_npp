# Notepad++ Excel grepping tool
Tired of people responding *Why would you use Notepad++ to view Excel files? They're binary!* when you ask about using Notepad++ with Excel? Well, now you can (sort of) read Excel files in Notepad++!

The original version supported only choosing text patterns in Excel files, but now you can filter filename patterns and sheet names, and you can also list worksheets names in each file.

To use in Notepad++, you must have [PythonScript 3.0.15](https://github.com/bruderstein/PythonScript/releases/tag/v3.0.15) or higher installed, and add `grep_in_excel.py` to the `plugins\PythonScript\scripts` folder of your Notepad++ installation. You can then execute the plugin with `Plugins->PythonScript->Scripts->grep_in_excel` from the main menu.

The user interface is VERY SIMPLE (because I didn't feel like making something fancier).

There is also a CLI tool that just prints JSON to the terminal. It has the same options as the Notepad++ add-in shown below.

![user interface of excel grepping tool](/UI%20example.PNG)
![example of results from successful grep](/results%20example.PNG)

## CHANGES ##

## [0.2.0] (2023-02-07)

### Added

- Filename filtering (by glob)
- sheet name filtering (by regex, case-insensitive only)
- option to list sheets that match filename and sheet name filters
- now automatically starts in directory of currently open file (in Notepad++) or current directory (in CLI)

### Fixed

Previously, the sheet names given didn't match the names as they would appear in an Excel workbook (e.g., the first sheet was `sheet1` even if it was actually named `foo bar`)

## [0.1.0] (2023-02-04)

Basic functionality added. Only filter text, no filename or sheet name filtering, no option to list only sheet names.
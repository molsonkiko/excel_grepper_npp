# Notepad++ Excel grepping tool
Tired of people responding *Why would you use Notepad++ to view Excel files? They're binary!* when you ask about using Notepad++ with Excel? Well, now you can (sort of) read Excel files in Notepad++!

To use in Notepad++, you must have [PythonScript 3.0.15](https://github.com/bruderstein/PythonScript/releases/tag/v3.0.15) or higher installed, and add `grep_in_excel.py` to the `plugins\PythonScript\scripts` folder of your Notepad++ installation. You can then execute the plugin with `Plugins->PythonScript->Scripts->grep_in_excel` from the main menu.

The user interface is VERY SIMPLE (because I didn't feel like making something fancier).

There is also a CLI tool that just prints JSON to the terminal. It has the same options as the Notepad++ add-in shown below.

![user interface of excel grepping tool](/UI%20example.PNG)
![example of results from successful grep](/results%20example.PNG)
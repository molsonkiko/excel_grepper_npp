import glob
import logging
import os
import re
import traceback
from xml.etree import ElementTree as ET
from zipfile import ZipFile
# yay, nothing outside standard library required!

__version__ =  '0.2.0'

logging.basicConfig(level=logging.WARNING)

class NotepadNotFound(Exception): pass

def grep_in_one_sheet(zf, sheet, is_match, shared_strings):
    '''
    zf: a pointer to an opened zip archive
    sheet: a file in a zip archive representing the data in an Excel worksheet
    is_match: a function that takes a string as input and returns
        true/false, whether that string matches a pattern/contains a substring
    shared_strings: a list of all the unique strings in the Excel workbook
    returns: a dict mapping 
    '''
    in_formulas = {}
    in_texts = {}
    with zf.open(sheet) as shf:
        text = str(shf.read(), encoding='utf-8')
        root = ET.fromstring(text)
        data = [elt for elt in root if elt.tag.endswith('sheetData')][0]
        for row in data:
            if not row.tag.endswith('row'):
                continue
            logging.debug('row number ' + row.attrib['r'])
            for col in row:
                cell_address = col.attrib['r']
                celltype = col.attrib.get('t')
                # t seems to be 'str' if the cell is a formula that
                # returns text
                # otherwise, t is (usually) 's' and the cell's 'v' element
                # (i.e., its value) is a number.
                # (sometimes t is 'e', which corresponds to a formula error)
                # The number is the index of the string in the
                # sharedStrings list.
                # Essentially, Excel cleverly figured out that it
                # could compress its data by keeping a list of
                # every text cell's value and then having each
                # text cell maintain a pointer to the appropriate
                # element in that list, rather than storing the text.
                # that way if you have like 100 cells that all
                # contain the same long text value (e.g., "Bob's Blazin' Burgers")
                # the cell's value will internally be "50" if that
                # happened to be the 50th unique string value in
                # the document.
                if not celltype:
                    continue # it's a number and we aren't interested
                formula = [cell.text for cell in col if cell.tag.endswith('f')]
                if formula:
                    formula_text = formula[0]
                    logging.debug(f'{cell_address = }, {celltype = }, {formula_text = }')
                    if is_match(formula_text):
                        logging.debug('adding text to list')
                        in_formulas[cell_address] = formula_text
                value = [cell.text for cell in col if cell.tag.endswith('v')]
                if value:
                    cell_text = value[0]
                    if celltype == 's':
                        # it's not the result of evaluating a formula
                        cell_text = shared_strings[int(cell_text)]
                        logging.debug(f'{cell_address = }, {celltype = }, {cell_text = }')
                    if is_match(cell_text):
                        logging.debug('adding text to list')
                        in_texts[cell_address] = cell_text
    sheet_results = {}
    if in_formulas:
        sheet_results['formulas'] = in_formulas
    if in_texts:
        sheet_results['text'] = in_texts
    return sheet_results

def grep_in_one_file(zf: ZipFile, is_match, sheet_name_regex, sheet_names_only) -> dict[str, dict[str, dict[str, str]]]:
    '''
    zf: a pointer to an open zip file (an excel workbook)
    is_match: a function taking a string as input and returning a boolean (whether
        the string matches a pattern)
    returns a dict similar to the following:
        {
        "sheet1": {
            "formulas": {
                "C5": "_xlfn.CONCAT(C3, \" ist der hund\")"
            },
            "text": {
                "C2": "hunden",
                "C5": "blutenheim ist der hund"
            }
        },
        "sheet2": {
            "text": {
                "B3": "hundeblarten"
            }
        }
    }
    '''
    sheets = []
    shared_strings_file = None
    workbook_file = None
    for f in zf.filelist:
        zfname = f.filename
        if zfname.endswith('xl/sharedStrings.xml'):
            shared_strings_file = f
        elif re.search('xl/worksheets/(?:.*?)\.xml$', f.filename):
            sheets.append(f)
        elif zfname.endswith('xl/workbook.xml'):
            workbook_file = f
    shared_strings = []
    if shared_strings_file:
        with zf.open(shared_strings_file) as shf:
            text = str(shf.read(), encoding='utf-8')
            root = ET.fromstring(text)
            for element in root:
                if element.tag.endswith('si'):
                    t = [e for e in element if e.tag.endswith('t')]
                    if t:
                        shared_strings.append(t[0].text)
    results = {}
    logging.debug(f'{sheets = }')
    logging.debug(f'{shared_strings = }')
    sheetnames = []
    if workbook_file:
        with zf.open(workbook_file) as wbf:
            # contains a "sheets" element that has name data for each sheet
            text = str(wbf.read(), encoding='utf-8')
            root = ET.fromstring(text)
            sheetdata = None
            for element in root:
                if element.tag.endswith('sheets'):
                    sheetdata = element
                    break
            for element in sheetdata:
                sheetname = element.attrib['name']
                if re.search(sheet_name_regex, sheetname, re.I):
                    sheetnames.append(sheetname)
    logging.debug(f'{sheetnames = }')
    if sheet_names_only:
        return sheetnames
    for sheet, sheetname in zip(sheets, sheetnames):
        logging.debug(f'{sheetname = }')
        sheet_results = grep_in_one_sheet(zf, sheet, is_match, shared_strings)
        if sheet_results:
            logging.info(f'got {sheet_results = }')
            results[sheetname] = sheet_results
    logging.info(f'got overall {results = }')
    return results
                        

def grep_in_excel_files(text_pattern, dirname, regex, recurse, ignorecase, sheet_name_regex, sheet_names_only, fname_pattern):
    '''
    text_pattern: a string or regex to match
    dirname: the absolute name of a directory containing Excel files
    regex: whether to do regular expression matching
    recurse: whether to also search in subdirectories of dirname
    ignorecase: whether to ignore case when trying to match text patterns
    sheet_name_regex: a regex that sheet names must match to be considered
    sheet_names_only: if true, only get a list of sheets that match
        sheet_name_regex in the workbooks of interest
    fname_pattern: a glob for determining what filenames to search
    returns: see EXAMPLES
EXAMPLES:
>>> grep_in_excel_files('hund', "c:\\users\\mjols\documents\\example nested excel dirs", regex=False, recurse=True, ignorecase=True)
{
    "c:\\users\\mjols\\documents\\example nested excel dirs\\bar\\barfoo\\barfoo.xlsx": {
        "sheet1": {
            "text": {
                "D1": "HUND"
            }
        }
    },
    "c:\\users\\mjols\\documents\\example nested excel dirs\\foo\\foo.xlsx": {
        "sheet1": {
            "formulas": {
                "C5": "_xlfn.CONCAT(C3, \\" ist der hund\\")"
            },
            "text": {
                "C2": "hunden",
                "C5": "blutenheim ist der hund"
            }
        },
        "sheet2": {
            "text": {
                "B3": "hundeblarten"
            }
        }
    }
}
    '''
    results = {}
    if regex:
        if ignorecase:
            is_match = lambda s: re.search(text_pattern, s, re.I)
        else:
            is_match = lambda s: re.search(text_pattern, s)
    else:
        if ignorecase:
            is_match = lambda s: text_pattern.lower() in s.lower()
        else:
            is_match = lambda s: text_pattern in s
    if recurse:
        fname_pattern = '**/' + fname_pattern
    files = glob.iglob(fname_pattern, root_dir=dirname, recursive=recurse)
    for file in files:
        fname = os.path.join(dirname, file)
        logging.info(f'reading {fname = }')
        try:
            zf = ZipFile(fname)
            file_result = grep_in_one_file(zf, is_match, sheet_name_regex, sheet_names_only)
            if file_result:
                results[fname] = file_result
        except:
            logging.error(f'Error in file {fname}')
            logging.error(traceback.format_exc())
        finally:
            if zf:
                zf.close()
    return results
    
def plugin_actions():
    '''
    prompt the user for args to grep_in_excel files (see above),
    then create a new buffer in Notepad++ and display results.
    '''
    import json
    try:
        from Npp import notepad, editor
        curdir = os.path.dirname(notepad.getCurrentFilename())
    except:
        raise NotepadNotFound
    user_input = notepad.prompt(
        'Enter your choices below',
        f'Excel Grepping tool {__version__}',
        '\r\n'.join((
            'text to search for:',
            f'absolute directory path:{curdir}',
            'use regex? (Y/N): N',
            'recursive search? (Y/N): N',
            'ignore case? (Y/N): N',
            'sheet name regex (leave blank to match all):',
            'list sheet names only? (Y/N): N',
            'filename pattern: *.xlsx'))
    )
    if not user_input:
        return
    choices = [re.split(': ?', line, 1)[1] for line in user_input.split('\r\n')]
    if len(choices) < 8:
        print("write your choices after the colons, and don't erase any lines")
        return
    pattern, dirname, regex, recurse, ignorecase, sheet_name_regex, sheet_names_only, fname_pattern = choices
    if not pattern:
        notepad.messageBox('Must enter a pattern!', 'Enter a pattern!')
        return
    if not dirname:
        notepad.messageBox('Must enter a directory name!', 'Enter a directory!')
        return
    if not os.path.exists(dirname):
        notepad.messageBox(f'Directory {dirname} does not exist', 'Invalid directory')
        return
    result = grep_in_excel_files(
        pattern,
        dirname,
        regex.lower() == 'y',
        recurse.lower() == 'y',
        ignorecase.lower() == 'y',
        sheet_name_regex,
        sheet_names_only.lower() == 'y',
        fname_pattern
    )
    notepad.new()
    editor.insertText(0, json.dumps(result, indent=4))

try:
    plugin_actions()
except NotepadNotFound:
    logging.debug('could not find Notepad++, doing CLI instead')
    import argparse
    import json
    parser = argparse.ArgumentParser()
    parser.add_argument('text_pattern')
    parser.add_argument('sheet_regex', nargs='?', default='', help='sheet names must match this regex to be considered')
    parser.add_argument('filename_pattern', nargs='?', default='*.xlsx', help='workbook filenames must match this glob to be considered')
    parser.add_argument('dirname', nargs='?', default=os.getcwd())
    parser.add_argument('--regex', '-x', action='store_true')
    parser.add_argument('--recurse', '-r', action='store_true')
    parser.add_argument('--ignorecase', '-i', action='store_true')
    parser.add_argument('--sheets_only', '-s', action='store_true', help='whether to only show a list of sheet names matching sheet_regex in the files')
    args = parser.parse_args()
    results = grep_in_excel_files(
        args.text_pattern,
        args.dirname,
        regex=args.regex,
        recurse=args.recurse,
        ignorecase=args.ignorecase,
        sheet_name_regex=args.sheet_regex,
        sheet_names_only=args.sheets_only,
        fname_pattern=args.filename_pattern
    )
    print(json.dumps(results, indent=4))
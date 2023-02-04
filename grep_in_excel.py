import logging
import os
import re
import traceback
from xml.etree import ElementTree as ET
from zipfile import ZipFile
# yay, nothing outside standard library required!

logging.basicConfig(level=logging.ERROR)

class NotepadNotFound(Exception): pass

def grep_in_one_file(fname: str, is_match) -> dict[str, dict[str, dict[str, str]]]:
    '''
    fname: an absolute file name
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
    if not fname.endswith('.xlsx'):
        return
    with ZipFile(fname) as zf:
        shared_strings = []
        shared_strings_file = [f for f in zf.filelist if 'xl/sharedStrings' in f.filename]
        if shared_strings_file:
            with zf.open(shared_strings_file[0]) as shf:
                text = str(shf.read(), encoding='utf-8')
                root = ET.fromstring(text)
                for element in root:
                    if element.tag.endswith('si'):
                        t = [e for e in element if e.tag.endswith('t')]
                        if t:
                            shared_strings.append(t[0].text)
        sheets = [f for f in zf.filelist
                  if re.search('xl/worksheets/(?:.*?)\.xml$', f.filename)]
        results = {}
        logging.info(f'reading {fname = }')
        logging.debug(f'{shared_strings = }')
        for sheet in sheets:
            sheetname = re.findall('xl/worksheets/(.*?)\.xml$', sheet.filename)[0]
            logging.debug(f'{sheetname = }')
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
            if sheet_results:
                logging.info(f'got {sheet_results = }')
                results[sheetname] = sheet_results
    logging.info(f'got overall {results = }')
    return results
                        

def grep_in_excel_files(text_pattern, dirname, regex=False, recurse=False, ignorecase=False):
    '''
    text_pattern: a string or regex to match
    dirname: the absolute name of a directory containing Excel files
    regex: whether to do regular expression matching
    recurse: whether to also search in subdirectories of dirname
    ignorecase: whether to ignore case when trying to match text patterns
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
    def one_file_search(fname):
        try:
            return grep_in_one_file(fpath, is_match)
        except:
            logging.error(f'Error in file {fname}')
            logging.error(traceback.format_exc())
    if recurse:
        for root, dirs, files in os.walk(dirname):
            for file in files:
                if not file.endswith('.xlsx'):
                    continue
                fpath = os.path.join(dirname, root, file)
                file_result = one_file_search(fpath)
                if file_result:
                    results[fpath] = file_result
    else: # only in top-level directory
        for file in os.listdir(dirname):
            fpath = os.path.join(dirname, file)
            file_result = one_file_search(fpath)
            if file_result:
                results[fpath] = file_result
    return results
    
def plugin_actions():
    '''
    prompt the user for args to grep_in_excel files (see above),
    then create a new buffer in Notepad++ and display results.
    '''
    import json
    try:
        from Npp import notepad, editor
    except:
        raise NotepadNotFound
    user_input = notepad.prompt(
        'Enter your choices below',
        'Excel Grepping tool',
        ('text to search for:\r\n'
        'absolute directory path:\r\n'
        'use regex? (Y/N): N\r\n'
        'recursive search? (Y/N): N\r\n'
        'ignore case? (Y/N): N')
    )
    if not user_input:
        return
    choices = [re.split(': ?', line, 1)[1] for line in user_input.split('\r\n')]
    if len(choices) < 5:
        print("write your choices after the colons, and don't erase any lines")
        return
    pattern, dirname, regex, recurse, ignorecase = choices
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
        ignorecase.lower() == 'y'
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
    parser.add_argument('dirname', nargs='?', default=os.path.dirname(os.path.abspath(__file__)))
    parser.add_argument('--regex', action='store_true')
    parser.add_argument('--recurse', action='store_true')
    parser.add_argument('--ignorecase', action='store_true')
    args = parser.parse_args()
    results = grep_in_excel_files(
        args.text_pattern,
        args.dirname,
        regex=args.regex,
        recurse=args.recurse,
        ignorecase=args.ignorecase)
    print(json.dumps(results, indent=4))
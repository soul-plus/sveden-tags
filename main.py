from jinja2 import Environment, FileSystemLoader
from openpyxl import load_workbook


workbook_contingent = 'chisl.xlsx'
workbook_vacant = 'vacant.xlsx'
workbook_priem = 'priem.xlsx'
workbook_perevod = 'perevod.xlsx'
workbook_praktika = 'praktika.xlsx'

env = Environment(loader=FileSystemLoader('templates'))

keys_contingent = ['code', 'name', 'level', 'form', 'BF', 'BFF', 'BR', 'BRF', 'BM', 'BMF', 'P', 'PF', 'All']
keys_vacant = ['code', 'name', 'level', 'prof', 'course', 'form', 'BFVacant', 'BRVacant', 'BMVacant', 'PVacant']
keys_priem = ['code', 'name', 'level', 'form', 'BF', 'BR', 'BM', 'P', 'score']
keys_perevod = ['code', 'name', 'level', 'form', 'NO', 'NT', 'NR', 'NE']
keys_praktika = ['code', 'name', 'year', 'profile', 'form', 'subject', 'technology', 'PRU', 'PRP', 'PRD']


def return_worksheet(workbook_name):
    path = f"C:/Users/program.KAZGIK/Desktop/t/{workbook_name}"
    book = load_workbook(filename=path)
    worksheets = book.worksheets
    sheet_ = worksheets[0]
    return sheet_


def html_tbl(template_name, keys, sheet, start_from=3):
    set_ = []
    dictionary = {}
    current_template = env.get_template(template_name)
    for i in range(start_from, sheet.max_row + 1):
        for j in range(len(keys)):
            if not(sheet[i][j].value is None):
                dictionary[keys[j]] = sheet[i][j].value
            # need to check if cell has a hyperlink
            else:
                dictionary[keys[j]] = '–'
        set_.append(dictionary.copy())
    return current_template.render(sheet=set_)


# final_tbl = html_tbl('vacant_table.html', keys_vacant, return_worksheet(workbook_vacant))
# final_tbl = html_tbl('contingent_table.html', keys_contingent, return_worksheet(workbook_contingent))
# final_tbl = html_tbl('priem_table.html', keys_priem, return_worksheet(workbook_priem))
final_tbl = html_tbl('perevod_table.html', keys_perevod, return_worksheet(workbook_perevod), 2)
# final_tbl = html_tbl('praktika_table.html', keys_praktika, return_worksheet(workbook_praktika))

with open("resultresult.html", "w", encoding="utf-8") as file:
    file.write(final_tbl)

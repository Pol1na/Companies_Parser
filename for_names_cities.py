import openpyxl

companies = []
names = []
cities = []
fname = 'companies.xlsx'


def get_companies():
    wb = openpyxl.load_workbook(fname)
    sheet = wb.get_sheet_by_name('ПРИМЕР')
    row = str(input('Введите количество компаний в файле Excel: '))
    print(row)
    for rowOfCellObjects in sheet['E2':'E'+row]:
        for cellObj1 in rowOfCellObjects:
            names.append(cellObj1.value)
    for rowOfCellObjects in sheet['G2':'G'+row]:
        for cellObj2 in rowOfCellObjects:
            cities.append(cellObj2.value)
    i = 0
    while i < int(row)-1:
        companies.append(str(names[i]) + ':' + str(cities[i]))
        i += 1
    return names,companies
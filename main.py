import requests
from bs4 import BeautifulSoup
import fake_useragent
import for_names_cities
import openpyxl


user = fake_useragent.UserAgent().random
INN = []
headers = {
    'user-agent': user}
for_names_cities.get_companies()
companies = for_names_cities.companies
inns = []

print('Не отключайте программу до ее завершения, иначе данные не сохранятся!')

def parse():
    print(len(companies))
    for company in companies:
        try:
            url = 'https://checko.ru/search?query=' + company.replace(':', ' ') + '&active=true'
            response = requests.get(url, headers=headers).text
            soup = BeautifulSoup(response, 'html.parser')
            items = soup.find('table').findAll('tr')
            counter = 0
            for item in items:
                find_comp_name = item.find('a', class_='default-link').get_text(strip=True).replace('"', '')
                find_city_name = item.find('td',
                                           class_='uk-width-1 uk-width-1-4@m uk-width-1-5@l uk-width-1-6@xl uk-text-left uk-text-right@m').get_text(
                    strip=True)
                city = find_city_name.upper()
                if 'МОСКОВ' in city:
                    city = 'МОСКВА'
                elif 'ЛЕНИНГРАД' in city:
                    city = 'САНКТ-ПЕТЕРБУРГ'

                if company.split(':')[0].upper() in find_comp_name.upper() and company.split(':')[1].upper() in city:
                    url = 'https://checko.ru' + item.find('a', class_='default-link').get('href')
                    response = requests.get(url, headers=headers).text
                    soup = BeautifulSoup(response, 'html.parser')
                    inn = soup.find('strong', id='copy-inn').get_text(strip=True)
                    inns.append(inn)
                    break
                else:
                    counter += 1
                    if counter == len(items):
                        inns.append('Не найдено')
        except AttributeError:
            inns.append('Не найдено')
        print(f'Собрано {len(inns)} компаний')
    save(inns)


def save(inns):
    fname = 'companies.xlsx'
    wb = openpyxl.load_workbook(fname)
    sheet = wb.get_sheet_by_name('ПРИМЕР')
    i = 2
    for inn in inns:
        sheet['H' + str(i)].value = inn
        i += 1
    wb.save(fname)
    wb.close()
    print('Работа программы завершена')

parse()

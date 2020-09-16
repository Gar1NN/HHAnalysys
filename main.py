
import requests as res
import json
import html2text as html2text
import xlsxwriter
import re
import pandas as pd

def find_id(parent_area: list, id:str) -> dict:
    for area in parent_area:
        if area["id"] == id:
            return area
    return None

def print_selection_list(selectinon_list: list, count_in_line:int):
    line = ""
    counter = 0
    max_len_area = max(selectinon_list, key=lambda x: len(" " + x["id"] + " " + x["name"]))
    max_len = len(" " + max_len_area["id"] + " " + max_len_area["name"])
    for item in sorted(selectinon_list, key=lambda x: x["name"]):
        selection = " " + item["id"] + " \"" + item["name"] + "\""
        line += selection
        for i in range(max_len - len(selection)):
            line += " "
        counter += 1
        if counter == count_in_line:
            counter = 0
            print(line)
            line = ""

def get_area() -> dict:
    russia_areas = res.get("https://api.hh.ru/areas/113").json()
    toOpen = russia_areas
    while True:
        print("Введите номер региона, в котором проводить поиск:")
        print("(Для вывода подрегионов в регионе введите m <номер региона который хотите раскрыть>)")
        print("(Чтобы вернуться в начало введите c или 0 для полного поиска)\n")
        print(toOpen["id"], toOpen["name"], sep="  ")

        print_selection_list(toOpen["areas"], 3)
        answer = input().split()
        if len(answer) == 0:
            print("Синтаксическая ошибка!")
            continue
        if answer[0] == "c":
            toOpen = russia_areas
            continue
        elif answer[0] == "m":
            found_area = find_id(toOpen["areas"], answer[1])
            if found_area == None:
                print("Вы ввели неправильный номер")
            else:
                if len(found_area["areas"]) == 0:
                    print("Нет вложенных подрегионов")
                    continue
                else:
                    toOpen = found_area
            continue
        elif answer[0] == "0":
            return None
        else:
            found_area = find_id(toOpen["areas"], answer[0])
            if found_area != None:
                return found_area
                break
            elif toOpen["id"] == answer[0]:
                return toOpen
            else:
                print("Вы ввели неправильный номер")
                continue
    raise Exception()

def get_specialization() -> dict:
    specializations = res.get("https://api.hh.ru/specializations").json()
    selected = None
    selectSpecialization = specializations
    while True:
        print("Введите профобласть, по которой ввести поиск или расскройте профобласть до специальностей:")
        print("(Для вывода специальностей в профобласти введите m <номер профобласти которую хотите расскрыть>)")
        print("(Чтобы вернуться в начало введите c  или 0 для полного поиска)\n")
        print_selection_list(selectSpecialization, 3)
        answer = input().split()
        if len(answer) == 0:
            print("Синтаксическая ошибка!")
            continue
        if answer[0] == "c":
            selectSpecialization = specializations
            continue
        elif answer[0] == "m":
            found_specialization = find_id(selectSpecialization, answer[1])
            if found_specialization == None:
                print("Вы ввели неправильный номер")
            else:
                selectSpecialization = found_specialization["specializations"]
                selected = found_specialization
            continue
        elif answer[0] == "0":
            return None
        else:
            found_specialization = find_id(selectSpecialization, answer[0])
            if found_specialization != None:
                return found_specialization["id"]
                break
            elif answer[0] == selected["id"]:
                return selected["id"]
            else:
                print("Вы ввели неправильный номер")
                continue
        raise Exception()

def get_industry() -> dict:
    industries = res.get("https://api.hh.ru/industries").json()
    selected = None
    selectIndustries = industries
    while True:
        print("Введите индустрию, по которой ввести поиск или расскройте индустрию до подиндустрий:")
        print("(Для вывода подиндустрий в индустрии введите m <номер индустрии которую хотите расскрыть>)")
        print("(Чтобы вернуться в начало введите c или 0 для полного поиска)\n")
        print_selection_list(selectIndustries, 3)
        answer = input().split()
        if len(answer) == 0:
            print("Синтаксическая ошибка!")
            continue
        if answer[0] == "c":
            selectIndustries = industries
            continue
        elif answer[0] == "m":
            found_industry = find_id(selectIndustries, answer[1])
            if found_industry == None:
                print("Вы ввели неправильный номер")
            else:
                selectIndustries = found_industry["industries"]
                selected = found_industry
            continue
        elif answer[0] == "0":
            return None
        else:
            found_industry = find_id(selectIndustries, answer[0])
            if found_industry != None:
                return found_industry["id"]
                break
            elif answer[0] == selected["id"]:
                return  selected["id"]
            else:
                print("Вы ввели неправильный номер")
                continue
    raise Exception()

def get_vacancy_search_fields() -> dict:
    fields = res.get("https://api.hh.ru/dictionaries").json()["vacancy_search_fields"]
    selected = None

    while True:
        print("Введите поле, по которой ввести поиск или нажмите Enter для поиска по всем полям")
        print("(Введите 0 для полного поиска\n")
        print_selection_list(fields, 3)
        answer = input().split()
        if len(answer) == 0:
            print("Синтаксическая ошибка!")
            continue
        if answer[0] == "0":
            return None
        else:
            return  answer[0]
    raise Exception()

parameters = dict()
print("Введите параметры для поиска!")
parameters.update(text=input("Текст поиска: "))
parameters.update(area=get_area()["id"])
parameters.update(per_page=100)
specialization = get_specialization()
if specialization != None:
    parameters.update(specialization=specialization)
industry = get_industry()
if industry != None:
    parameters.update(industries=industry)
field = get_vacancy_search_fields()
if field != None:
    parameters.update(search_field=field)
parameters.update(page=0)
print("Параметры приняты!")


while True:
    count_pages = input("Введите количество страниц которое необходимо обработать: ")
    if count_pages.isdigit():
        count_pages = int(count_pages)
        break
    else:
        print("Это не число!")


pages = list()

for i in range(count_pages):
    parameters["page"] = i
    pages.append(res.get("https://api.hh.ru/vacancies", params=parameters).json())

vacancies = list()
print("Подождите, собираем информацию о всех вакансиях!")
counter = 0
for page in pages:
    counter += 1
    for item in page['items']:
        vacancies.append(res.get("https://api.hh.ru/vacancies/" + item['id']).json())
    print("Загрузили страницу №" + str(counter))

if input("Сохранить в json файл?(y\\n)") == "y":
    name = input("Введите имя файла")
    fp = open(name + ".json", 'w')
    json.dump(vacancies, fp)
skills_set = set()
skills_full_array = ""
for i in vacancies:
    for skill in i['key_skills']:
        skills_set.add(str.lower(skill['name']))
        skills_full_array += str.lower(skill['name']) + " "
skills_full_array = skills_full_array.split()

# Запись в Excel файл
workbook = xlsxwriter.Workbook('Vacancy_7.xlsx')
worksheet = workbook.add_worksheet()
# Добавим стили форматирования
bold = workbook.add_format({'bold': 1})
bold.set_align('center')
center_H_V = workbook.add_format()
center_H_V.set_align('center')
center_H_V.set_align('vcenter')
center_V = workbook.add_format()
center_V.set_align('vcenter')
cell_wrap = workbook.add_format()
cell_wrap.set_text_wrap()
red_words = workbook.add_format({'font_color': 'red'})

# Настройка ширины колонок
worksheet.set_column(0, 0, 35)  # A  https://xlsxwriter.readthedocs.io/worksheet.html#set_column
worksheet.set_column(1, 1, 135)  # B
worksheet.set_column(2, 2, 20)  # C
worksheet.set_column(3, 3, 40)  # D
worksheet.set_column(4, 4, 135)  # E
worksheet.set_column(5, 5, 45)  # F

worksheet.write('A1', 'Название вакансии', bold)
worksheet.write('B1', 'Компания', bold)
worksheet.write('C1', 'Требования к кандидату', bold)
worksheet.write('D1', 'Ключевые навыки', bold)
worksheet.write('E1', 'Ссылка', bold)

row = 1
col = 0
for vacancy in vacancies:
    worksheet.write_string(row, col, vacancy['name'], center_V)
    local_skills = set()
    words = []
    description = html2text.html2text(vacancy['description']).replace('**', "").replace('*', ' ')
    for word in description.split(' '):
        for skill in re.split(r'[\s,.!()/<>*]', word):
            if str.lower(skill) in skills_set:
                words.append(red_words)
                local_skills.add(skill)
                break
        words.append(word + " ")
    words.append(cell_wrap)
    worksheet.write_rich_string(row, col + 2, *words)
    #worksheet.write_string(row, col + 1, description, cell_wrap)
    skills_str = ""
    for i in local_skills:
        skills_str += i + ', '
    #worksheet.write_string(row, col + 2, skills_str, center_H_V)
    skills_str = ""
    for i in vacancy['key_skills']:
        skills_str += i['name'] + ', '
    worksheet.write_string(row, col + 3, skills_str, center_H_V)
    worksheet.write_string(row, col + 1, vacancy['employer']['name'], center_H_V)
    worksheet.write_string(row, col + 4, vacancy['alternate_url'], cell_wrap)
    row += 1
workbook.close()

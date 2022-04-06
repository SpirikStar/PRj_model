import openpyxl
import re

'''
Задача:
Оптимизировать исходный код. 
Поля db должны быть в точности как в словаре excelData, (разрешается нижний регистор). Последовательность такая-же, как и словаре. 
Нужно проверить. 

'''


excelData = {
    1: 'KeyUser',
    2: 'Фамилия',
    3: 'Имя',
    4: 'Отчество',
    5: 'Специальность',
    6: 'Год поступления',
    7: 'Уровень образования',
    8: 'Форма обучения',
    9: 'Логин',
    10: 'Пароль',
    11: 'Группа',
    12: 'Курс',
    13: 'Семестр',
    14: 'Подгруппа',
    15: 'Лицензия',
    16: 'Статус пользователя',
}


def run(path: str):
    book = openpyxl.open(path)
    sheet = book.active
    # Попробуй найти альтернативу этому длинному условию and
    '''
    sheet['B1'].value.lower() метод lower подсвечивается в том случае, если тип является стройкой. 
    В данном случае он переводит в нужный формат, но есть мелкие ошибки. Попробуй исправисть это и улучшить код 
    '''
    if str(sheet['A1'].value).lower() == excelData[1].lower() and sheet['B1'].value.lower() == excelData[2].lower() and sheet['C1'].value.lower() == excelData[3].lower() and sheet['D1'].value.lower() == excelData[4].lower() and sheet['E1'].value.lower() == excelData[5].lower() and sheet['F1'].value.lower() == excelData[6].lower() and sheet['G1'].value.lower() == excelData[7].lower() and sheet['H1'].value.lower() == excelData[8].lower() and sheet['I1'].value.lower() == excelData[9].lower() and sheet['J1'].value.lower() == excelData[10].lower() and sheet['K1'].value.lower() == excelData[11].lower() and sheet['L1'].value.lower() == excelData[12].lower() and sheet['M1'].value.lower() == excelData[13].lower() and sheet['N1'].value.lower() == excelData[14].lower() and sheet['O1'].value.lower() == excelData[15].lower() and sheet['P1'].value.lower() == excelData[16].lower():
        for Ikeyg in range(2, len(sheet['A'])+1): # явное указание Len, сколько строчек заполненно
            _KeyUser = sheet[f'A{str(Ikeyg)}'].value
            _Surname = sheet[f'B{str(Ikeyg)}'].value
            _Name = sheet[f'C{str(Ikeyg)}'].value
            _MiddleName = sheet[f'D{str(Ikeyg)}'].value
            _Specialization = sheet[f'E{str(Ikeyg)}'].value
            _Year_of_admission = sheet[f'F{str(Ikeyg)}'].value
            _Education_level = sheet[f'G{str(Ikeyg)}'].value
            _Form_of_training = sheet[f'H{str(Ikeyg)}'].value
            _Login = sheet[f'I{str(Ikeyg)}'].value
            _Pasw = sheet[f'J{str(Ikeyg)}'].value
            _Group = sheet[f'K{str(Ikeyg)}'].value
            _Course = sheet[f'L{str(Ikeyg)}'].value
            _Term = sheet[f'M{str(Ikeyg)}'].value
            _Subgroup = sheet[f'N{str(Ikeyg)}'].value
            _License = sheet[f'O{str(Ikeyg)}'].value
            _Status = sheet[f'P{str(Ikeyg)}'].value


if __name__ == "__main__":
    path = 'db.xlsx'
    if re.search('.xlsx', path):  # 1. Тебе нужно эту строчку поменять, и воспользоваться с помощью import os и т.д., проверить, что расширение данного файла .xlsx . (Не обязательно)
        run(path)

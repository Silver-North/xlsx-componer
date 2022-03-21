from genericpath import isfile
from os import listdir, path, walk
from xlrd import open_workbook
from xlwt import Workbook


filenames = next(walk('xls/'), (None, None, []))[2]
count = len([f for f in listdir('xls/') if isfile(path.join('xls/', f))])
writing_book = Workbook()
writing = writing_book.add_sheet('rezult')


for j in range(count):
    workbook = open_workbook(f'xls/{filenames[j]}')
    worksheet = workbook.sheet_by_index(0)
    lists = {
        '№': worksheet.cell(0,1).value,
        'Дата': worksheet.cell(1,1).value,
        'Склад': worksheet.cell(2,1).value,
        'Фамилия': worksheet.cell(3,1).value,
        'Имя': worksheet.cell(4,1).value,
        'Отчество': worksheet.cell(5,1).value,
        'Пол': worksheet.cell(6,1).value,
        'Телефон': worksheet.cell(7,1).value,
        'Контрагент': worksheet.cell(8,1).value,
        'Автор': worksheet.cell(9,1).value,
        'Комментарий': worksheet.cell(10,1).value,
        'Товары и услуги': worksheet.cell(13,1).value,
        'Наименование': worksheet.cell(14,1).value,
        'ОРТ-4. Оттиск альгинат': worksheet.cell(15,1).value,
        'ОРД-01. Консультация': worksheet.cell(16,1).value,
        'Количество': f"{worksheet.cell(15,2).value},{worksheet.cell(16,2).value}",
        'Цена': f"{worksheet.cell(15,3).value},{worksheet.cell(16,3).value}",
        'Сумма': f"{worksheet.cell(15,4).value},{worksheet.cell(16,4).value}"
    }

    row = writing.row(j)
    for i in enumerate(lists):
        if j == 0:
            row.write(i[0], i[1])
            rows = writing.row(1)
            rows.write(i[0], lists[i[1]])
            continue
        else:
            row = writing.row(j+1)
            row.write(i[0], lists[i[1]])

writing_book.save('rezult.xls')
print('complete work script!!!')

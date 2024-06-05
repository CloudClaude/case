from openpyxl import load_workbook
import PySimpleGUI as sg
from datetime import datetime

element_size = (25, 1)

layout = [
    [sg.Text('Наименование', size=element_size), sg.Input(key='name', size=element_size)],
    [sg.Text('Действующее вещество', size=element_size), sg.Input(key='chemicals', size=element_size)],
    [sg.Text('Форма выпуска', size=element_size), sg.Input(key='place', size=element_size)],
    [sg.Text('Дозировка', size=element_size), sg.Input(key='dose', size=element_size)],
    [sg.Text('Размер', size=element_size), sg.Input(key='size', size=element_size)],
    [sg.Text('Рецептурный препарат', size=element_size), sg.Input(key='receipt', size=element_size)],
    [sg.Text('Особое место хранения', size=element_size), sg.Input(key='splace', size=element_size)],
    [sg.Text('Прекурсор', size=element_size), sg.Input(key='precursor', size=element_size)],
    [sg.Text('Количество', size=element_size), sg.Input(key='count', size=element_size)],
    [sg.Button('Добавить'), sg.Button('Закрыть')]
]

window = sg.Window('Учет поступивших лекарственных средств', layout, element_justification='center')

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Закрыть':
        break
    if event == 'Добавить':
        try:
            wb = load_workbook('аптека.xlsx')
            sheet = wb['Лист1']
            ID = len(sheet['ID'])
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            data = [
                ID,
                values['name'],
                values['chemicals'],
                values['place'],
                values['dose'],
                values['size'],
                values['receipt'],
                values['splace'],
                values['precursor'],
                values['count'],
                time_stamp
            ]
            sheet.append(data)
            wb.save('аптека.xlsx')

            # Очистка полей ввода
            for key in values:
                window[key].update(value='')
            window['name'].set_focus()
            sg.popup('Данные сохранены')
        except PermissionError:
            sg.popup('Ошибка доступа', 'Файл используется другим пользователем.\nПопробуйте позже.')


window.close()
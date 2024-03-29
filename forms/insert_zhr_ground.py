﻿from openpyxl import load_workbook
from datetime import *
import os.path


def create_row(ws, ws_style, current_row, coord, height=13.50):
    ws.insert_rows(current_row)
    ws.row_dimensions[current_row].height = height
    for index in range(1, 15):
        ws.cell(current_row, index)._style = ws_style.cell(*coord)._style


def insert_zhr_ground(def_db, road_programm, road_db, tube, km_start, km_finish, dy_tube):
    wb_zhr = load_workbook(road_programm + road_to_excel)
    ws_l1, ws_l2, ws_l3 = wb_zhr['Tittle'], wb_zhr['Table'], wb_zhr['Styles']
    current_row = 4
    # -----------------------Title----------
    ws_l1['A1'] = def_db[0]['otv'][2]
    ws_l1['A3'] = 'Устранение дефектов на секциях %s, %s-%s км, Ду %s мм.' % (tube, km_start, km_finish, dy_tube)
    ws_l1['V11'] = km_start + ' км'
    ws_l1['AR11'] = km_finish + ' км'
    ws_l1['S17'] = '%s %s %s' % (def_db[0]['otv'][1], def_db[0]['otv'][2], def_db[0]['otv'][0])
    ws_l1['BO25'] = def_db[0]['date'][0].strftime('%d.%m.%Y') + ' года'
    ws_l1['BO27'] = def_db[-1]['date'][2].strftime('%d.%m.%Y') + ' года'
    # -----------------------Table----------
    day_s = []
    for remont in def_db:  # creat date list
        date_1 = remont['date'][0]
        date_2 = remont['date'][2]
        for date_num in (date_1, date_2):
            if date_num not in day_s:
                day_s.append(date_num)
    day_s = sorted(day_s)  # sort date list
    print(day_s)
    for i in day_s:
        for remont in def_db:
            # V = round(remont['rand_value'][2] * remont['rand_value'][3] * (remont['dist'][4] - remont['dist'][3]) * 1.4)
            if remont['date'][0] == i:
                temp_row = '5'  # Excavation
            elif remont['date'][2] == i:
                temp_row = '6'  # covering
            else:
                continue
            create_row(ws_l2, ws_l3, current_row, (1, 1), height=62)
            ws_l2['A' + str(current_row)] = i.strftime('%d.%m.%y') + ' г.'  # date
            ws_l2['B' + str(current_row)] = ws_l3['B' + temp_row].value
            ws_l2['C' + str(current_row)] = ws_l3['C' + temp_row].value
            ws_l2['D' + str(current_row)] = '%s км\n%s м\n-\n%s м' % (remont['km'],
                                                                      remont['dist'][3],
                                                                      remont['dist'][4])  # range start
            ws_l2['E' + str(current_row)] = ws_l3['E' + temp_row].value
            ws_l2['F' + str(current_row)] = ws_l3['F' + temp_row].value
            ws_l2['G' + str(current_row)] = ws_l3['D' + temp_row].value
            ws_l2['H' + str(current_row)] = ws_l3['H' + temp_row].value
            ws_l2['I' + str(current_row)] = ws_l3['I' + temp_row].value
            ws_l2['J' + str(current_row)] = ws_l3['J' + temp_row].value
            if 'grnd_maker' in remont:
                temp_man = 'grnd_maker'
            else:
                temp_man = 'otv'
            ws_l2['L' + str(current_row)] = '%s %s\n_______________\n%s' % (remont[temp_man][1],
                                                                            remont[temp_man][2],
                                                                            remont[temp_man][0])
            ws_l2['M' + str(current_row)] = '%s %s\n_______________\n%s' % (remont['sk'][1],
                                                                            remont['sk'][2],
                                                                            remont['sk'][0])
            current_row += 1

    # -----------------------Заключение
    ws_l2['M' + str(current_row + 1)] = 'Работы закончены. Журнал закрыт.'
    ws_l2['M' + str(current_row + 2)] = '%s %s %s _____________ %s г.' % (
        remont['otv'][1], remont['otv'][2], remont['otv'][0], remont['date'][2].strftime('%d.%m.%Y'))
    ws_l2.print_area = 'A1:N' + str(current_row + 2)
    del (wb_zhr['Styles'])
    road_db = os.path.normpath(road_db)
    road_db = os.path.join(road_db, 'Журналы')
    if not os.path.exists(road_db):
        os.makedirs(road_db)
    road_db = os.path.join(road_db, '6. Журнал земляных работ.xlsx')
    wb_zhr.save(road_db)


road_to_excel = '/Excell/zhr_ground.xlsx'
if __name__ == '__main__':
    road_to_excel = '../Excell/zhr_ground.xlsx'
    road_db = os.getcwd() + '\..\..\Тест'
    road_programm = ''
    tube, km_start, km_finish, dy_tube = 'МНПП "Рязань-Тула-Орел" отвод на новомосковскую НБ, ДТ', '12', '13', '530'
    def_db = [{'sec': '12', 'date': [date(2018, 12, 12), date(2018, 12, 13), date(2018, 12, 14)],
               'dist': [10.1, 9.11, 11.09, 5.44, 14.76], 'km': '1', 'dl_muft': 1.0,
               'rand_value': [0.48533824224793676, 3.6729518931973657, 2.861346564462221, 2.705419511460805],
               'defect': {'1': {'sec': '12', 'dist': '123', 'lab': 'Риска с ППОШ с расслоением', 'dl': '5', 'sh': '6',
                                'gl': '7', 'type': 'Муфта П1'},
                          '2': {'sec': '12', 'dist': '123', 'lab': 'Риска с ППОШ с расслоением и ВНП', 'dl': '234',
                                'sh': '234', 'gl': '65', 'type': 'Муфта П2'},
                          '3': {'sec': '12', 'dist': '534536', 'lab': 'Риска с ППОШ с расслоением и ВНП', 'dl': '45534',
                                'sh': '345345', 'gl': '3535', 'type': 'Шлифовка'}},
               'otv': ('Афанасьев А.Б.', 'Начальник РУ№1', 'ЦРС "Рязань"'),
               'contr': ('Клименко Д.А.', 'Начальник ЛАЭС №1', 'ППС "Плавск'),
               'sk': ('Мирошкин М.В.', 'Инженер СК', 'ООО "Сег"'),
               'lkk': ('Козин А.П.', 'Инженер-дефектоскопист', 'ЛККиД')},
              {'sec': '13', 'date': [date(2018, 12, 13), date(2018, 12, 14), date(2018, 12, 14)],
               'dist': [20.0, 18.54, 21.46, 15.55, 24.45], 'km': '2', 'dl_muft': 2.0,
               'rand_value': [0.4581174447339836, 2.989229650534175, 2.9019640191285303, 2.8115670668469814],
               'defect': {
                   '1': {'sec': '13', 'dist': '12313', 'lab': 'Потеря метала', 'dl': '12', 'sh': '132', 'gl': '123',
                         'type': 'Шлифовка'},
                   '2': {'sec': '13', 'dist': '2344234', 'lab': 'Потеря метала', 'dl': '34', 'sh': '4343', 'gl': '434',
                         'type': 'Шлифовка'}}, 'otv': ('Кулешов А.Б.', 'Начальник РУ№3', 'ЦРС "Рязань"'),
               'contr': ('Клименко Д.А.', 'Начальник ЛАЭС №1', 'ППС "Плавск'), 'sk': ('', '', ''),
               'lkk': ('Козин А.П.', 'Инженер-дефектоскопист', 'ЛККиД')},
              {'sec': '14', 'date': [date(2018, 12, 14), date(2018, 12, 15), date(2018, 12, 16)],
               'dist': [30.0, 27.94, 32.06, 27.74, 32.26], 'km': '3', 'dl_muft': 3.0,
               'rand_value': [0.5619183603415756, 3.19813104261683, 3.0060601832822327, 2.9168215194726352], 'defect': {
                  '1': {'sec': '14', 'dist': '1131213,32', 'lab': 'Риска', 'dl': '1233', 'sh': '12331', 'gl': '2133',
                        'type': 'Муфта П1'},
                  '2': {'sec': '14', 'dist': '321313,21', 'lab': 'Риска', 'dl': '23', 'sh': '234', 'gl': '2342',
                        'type': 'Муфта П1'}}, 'otv': ('Макаров М.А.', 'Начальник АРС', 'ЛПДС "Рязань"'),
               'contr': ('Клименко Д.А.', 'Начальник ЛАЭС №1', 'ППС "Плавск'),
               'sk': ('Мирошкин М.В.', 'Инженер СК', 'ООО "Сег'),
               'lkk': ('Козин А.П.', 'Инженер-дефектоскопист', 'ЛККиД')},
              {'sec': '15', 'date': [date(2018, 12, 14), date(2018, 12, 16), date(2018, 12, 16)],
               'dist': [160.0, 157.71, 162.29, 154.79, 165.21], 'km': '16', 'dl_muft': 4.0,
               'rand_value': [0.2872474356253941, 2.915977705201697, 2.9123558666117617, 2.703436382412381], 'defect': {
                  '1': {'sec': '15', 'dist': '123131,12', 'lab': 'Риска', 'dl': '123', 'sh': '123', 'gl': '13',
                        'type': 'Муфта П2'},
                  '2': {'sec': '15', 'dist': '1231', 'lab': 'Риска', 'dl': '3213', 'sh': '123', 'gl': '312',
                        'type': 'Муфта П2'},
                  '3': {'sec': '15', 'dist': '432423', 'lab': 'Риска', 'dl': '13', 'sh': '123', 'gl': '13',
                        'type': 'Муфта П2'}}, 'otv': ('Кулешов А.Б.', 'Начальник РУ№3', 'ЦРС "Рязань"'),
               'contr': ('Клименко Д.А.', 'Начальник ЛАЭС №1', 'ППС "Плавск'), 'sk': ('', '', ''),
               'lkk': ('Козин А.П.', 'Инженер-дефектоскопист', 'ЛККиД')}]
    insert_zhr_ground(def_db, road_programm, road_db, tube, km_start, km_finish, dy_tube)

﻿from openpyxl import load_workbook
from openpyxl.worksheet.copier import WorksheetCopy
import os

def insert_akt_vh(akt_vh,road_programm,road_db,tube,km_start,km_finish,dy_tube):
    wb = load_workbook(road_programm+road_to_excel)
    for i,x in enumerate(akt_vh):
        current_wb=wb.create_sheet('%s. %s'%(i+3,x['name']))
        copy_ws=WorksheetCopy(wb['Empty'],current_wb)
        copy_ws.copy_worksheet()
        copy_ws._copy_cells()
        copy_ws._copy_dimensions()
        current_wb.print_area = 'A1:AB121'
        current_wb.page_margins=wb['Empty'].page_margins
        #--------------начинаем заполнение текущего акта
        current_wb['G3']=x['otv'][2] #служба
        current_wb['T3']='Устранение дефектов на секциях %s, %s-%s км, Ду %s мм.'%(tube,km_start,km_finish,dy_tube) #заголовок
        current_wb['T6']='%s, %s-%s км.'%(tube,km_start,km_finish) #заголовок
        current_wb['W9']=x['date']+' г.' #дата
        current_wb['P9']=i+3                     #номер акта
        current_wb['B11']='%s, в кол-ве %s %s.'%(x['full_name'],x['kol'],x['marker'])            #номер акта
        current_wb['H14']=' '.join((x['otv'][1],x['otv'][2],x['otv'][0]))
        current_wb['H16']=' '.join((x['contr'][1],x['contr'][2],x['contr'][0]))
        current_wb['H18']=' '.join((x['kk'][1],x['kk'][2],x['kk'][0]))
        current_wb['H22']=' '.join((x['sk'][1],x['sk'][2],x['sk'][0]))
        current_wb['M35']=x['doc']
        current_wb['J39']=x['param']
        current_wb['H50']=x['TU']
        current_wb['L51']=x['otv'][0]
        current_wb['L54']=x['contr'][0]
        current_wb['L57']=x['kk'][0]
        current_wb['L60']=x['sk'][0]
    del(wb['Empty'])
    #---------------------электроды
    ws1 = wb['1. 3.2 электроды 53.70']
    ws1['U4'] = akt_vh[0]['date']+' г.'
    ws1['O6'] = akt_vh[0]['otv'][0]
    ws1['O7'] = akt_vh[0]['otv'][0]
    ws1['AE47'] = akt_vh[0]['otv'][0]
    ws1['AE49'] = akt_vh[0]['otv'][0]
    ws2 = wb['2. 4.0 электроды 53.70']
    ws2['U4'] = akt_vh[0]['date']+' г.'
    ws2['O6'] = akt_vh[0]['otv'][0]
    ws2['O7'] = akt_vh[0]['otv'][0]
    ws2['AE47'] = akt_vh[0]['otv'][0]
    ws2['AE49'] = akt_vh[0]['otv'][0]
    road=road_db.rpartition('/')[0]+'/Входной контроль'
    if not os.path.exists(road):
        os.makedirs(road) 
    wb.save(road+'/Акты входного контроля.xlsx')

road_to_excel='/Excell/akt_vh.xlsx'
if __name__=='__main__':
    #import os
    road_to_excel='../Excell/akt_vh.xlsx'
    road=os.getcwd()
    road_db=road.replace('\\','/')
    road_programm = ''
    tube,km_start,km_finish,dy_tube='МНПП "Рязань-Тула-Орел" отвод на новомосковскую НБ, ДТ','12','13','530'
    akt_vh=[{'name': 'Абразив', 'full_name': 'Абразим 123', 'marker': 'кг', 'TU': 'ТУ5', 'param': '-', 'date': '02.03.2019', 'doc': 'Паспорт 34234', 'kol': '3', 'otv': ('Кулешов А.Б.', 'Начальник РУ№3', 'ЦРС "Рязань"'), 'contr': ('Макаров М.А.', 'Начальник АРС', 'ЛПДС "Рязань"'), 'kk': ('Макаров М.А.', 'Начальник АРС', 'ЛПДС "Рязань"'), 'sk': ('Мирошкин М.В.', 'Инженер СК', 'ООО "Сег"')},
            {'name': 'Праймер', 'full_name': 'Праймер ПЛ', 'marker': 'рул', 'TU': 'ТУ3', 'param': 'длина= мм, ширина= мм', 'date': '01.12.2019', 'doc': 'Сертификат 14', 'kol': '1', 'otv': ('Афанасьев А.Б.', 'Начальник РУ№1', 'ЦРС "Рязань"'), 'contr': ('Макаров М.А.', 'Начальник АРС', 'ЛПДС "Рязань"'), 'kk': ('Макаров М.А.', 'Начальник АРС', 'ЛПДС "Рязань"'), 'sk': ('Мирошкин М.В.', 'Инженер СК', 'ООО "Сег"')},
            {'name': 'Муфта П1 530', 'full_name': 'Муфта КМТ 530', 'marker': 'шт', 'TU': 'ТУ8', 'param': 'диаметр= мм, длина= мм, толщина стенки= м', 'date': '05.12.2019', 'doc': 'Паспорт 12', 'kol': '3534', 'otv': ('Кулешов А.Б.', 'Начальник РУ№3', 'ЦРС "Рязань"'), 'contr': ('Клименко Д.А.', 'Начальник ЛАЭС №1', 'ППС "Плавск'), 'kk': ('', '', ''), 'sk': ('', '', '')}, {'name': 'Герметик', 'full_name': 'состав полмерный Герметик', 'marker': 'компл.', 'TU': 'ТУ6', 'param': '-', 'date': '07.12.2019', 'doc': '645', 'kol': '3', 'otv': ('Кулешов А.Б.', 'Начальник РУ№3', 'ЦРС "Рязань"'), 'contr': ('Клименко Д.А.', 'Начальник ЛАЭС №1', 'ППС "Плавск'), 'kk': ('Клименко Д.А.', 'Начальник ЛАЭС №1', 'ППС "Плавск'), 'sk': ('Мирошкин М.В.', 'Инженер СК', 'ООО "Сег"')}]
    insert_akt_vh(akt_vh,road_programm,road_db,tube,km_start,km_finish,dy_tube)
    

from tkinter import *
from tkinter.messagebox import showinfo,showerror
from tkinter import filedialog
import shelve,os
import traceback

import forms.def_main as def_main
from forms.insert_akt_all import insert_akt_all
from forms.insert_akt_vh import insert_akt_vh
from forms.insert_main import insert_main
from forms.insert_zhr_ground import insert_zhr_ground
from forms.insert_zhr_izol import insert_zhr_izol
from forms.insert_zhr_ozhr import insert_zhr_ozhr
from forms.insert_zhr_svarki import insert_zhr_svarki
from forms.insert_zhr_vh import insert_zhr_vh
from forms.insert_zhr_zam import insert_zhr_zam

#-------блок вывода ошибок в лог--------------------------#
def log_uncaught_exceptions(ex_cls, ex, tb):
    text = '{}: {}:\n'.format(ex_cls.__name__, ex)
    import traceback
    text += ''.join(traceback.format_tb(tb))
    with open('error.txt', 'w', encoding='utf-8') as f:
        f.write(text)
import sys
sys.excepthook = log_uncaught_exceptions
#--------------------------------------------------------#

def save_db():                 #Фукнция сохранить БД
    global road_db
    main_theme='Устранение дефектов методом выборчного ремонта на секциях %s, %s-%s км, Ду %s мм.'%tuple(data_tube[i].get() for i in data)
    if road_db=='':
        save_as_db()
    else:
        try:
            if road_db[-4:-3]=='.':
                road_db=road_db[:road_db.rfind('.')]            #если перезаписываем файл, то файл идет с расширением, убираем его.
            db=shelve.open(road_db)
            for i in data:
                db[i]=data_tube[i].get()
                db['def_db']=def_db
                db['akt_vh']=akt_vh
            showinfo(title='Сохранение!', message=main_theme)
        finally:
            db.close()
    
def save_as_db():               #Фукнция сохранить БД как...
    global road_db
    temp=filedialog.asksaveasfilename(filetypes=(("Data files","*.dat"),))
    if temp!='':                #если открыли окно сохранения и закрыли его не выбрав файл, то путь становится пустым. Исключаем это.
        road_db=temp
        save_db()

def open_db():                  #Фукнция открытия БД
    global road_db,def_db,akt_vh,last_index
    temp=filedialog.askopenfilename(filetypes=(("Data files","*.dat"),))[:-4] 
    if temp!='':                        #если открыли окно "открытия" и закрыли его не выбрав файл, то путь становится пустым. Исключаем это.  
        try: 
            road_db=temp                #запоминаем выбранный путь до файла БД
            db=shelve.open(road_db)     #открываем файл БД
            for i in data:              #Заполняем поля ввода в разделе №1
                data_tube[i].delete(0,END)
                data_tube[i].insert(0,db[i])
            def_db=db['def_db']         #считываем информацию по дефетам
            akt_vh=db['akt_vh']         #считываем информацию по актам входного контроля
            last_index=None                     #при открытии,создании новой бд возвращаем занчения по умолчанию в поля раздела №3
            for i in (date_vh,doc_vh,kol_vh):   #очищаем поля ввода
                i.delete(0,END)
            for i in (var_otv,var_contr,var_kk,var_sk): #очищаем выпадающие списки
                i.set('')
            re_num_vh()                       
        finally:
            db.close()
            

def add_akt_vh():                               #функция добавления акта входного контроля
    temp=materials[var_mat.get()]               #временная переменная для уменьшения кода
    akt_vh.append({'name':temp[0],              #краткое название материала
                   'full_name':temp[1],         #полное название материала
                   'marker':temp[2],            #еденица измерения
                   'TU':temp[3],                #ТУ
                   'param':temp[4]              #параметры материала
                   })
    lbox.insert(END,'%d. %s'%(len(akt_vh)+2,temp[0]))   #добавлениt акта в листбокс
    var_mat.set('')
    
def del_akt_vh():                       #функция удаления акта
    try:
        del(akt_vh[lbox.curselection()[0]]) #удаляем выбранный акт из временной БД
        re_num_vh()                         #обновления списка актов
    except IndexError:
        showerror('Индекс!','Ошибка индекса!')
        traceback.print_exc(file=open('error.txt', 'w', encoding='utf-8'))

def re_num_vh():
    lbox.delete(0,END)                  #очищение виджета lbox  перед заполнением
    for i,j in enumerate(akt_vh):       #обновление нумерации
        lbox.insert(END,'%s. %s'%(i+3,akt_vh[i]['name']))

def update_box(a):                                                  #функция обновления информации при щелчку по акту в листбоксе
    global last_index
    temp=lbox.curselection()[0]                                     #временная переменная запомнить индекс выбранного элекмента
    if last_index!=None:
        save_akt_vh()                                               #сохраняет введеные значения предыдущего акта при щелчке на следующем акте(зменяет кнопку сохранения)
        lbox.selection_set(temp)                                    #после сохранения пропадает выделение, восстанавливаем
    last_index=temp
    for (i,j) in zip(('date','doc','kol'),(date_vh,doc_vh,kol_vh)): 
        if i in akt_vh[temp]:                                       #если запись существует
            j.delete(0,END)                                         #то удаляем текущее значение от предыдущего выделения
            j.insert(0,akt_vh[temp][i])                             #и вставляем значение текущего выбранного элемента
        else:
            if j!=date_vh:                                          #иначе это новая запись
                j.delete(0,END)                                     #и удаляем поля кроме даты
            else:                                                   #если же это поле даты, то помечаем красным для понимания, что дата осталась с предыдущего акта
                j.config(fg='Red')
    for (i,j) in zip(('otv','contr','kk','sk'),(var_otv,var_contr,var_kk,var_sk)):
        if i in akt_vh[temp]:                       #если есть записи по ответсвенным, то
            j.set(akt_vh[temp][i][0])               #обновляем выпадающие списки с ответсвенными лицами

def save_akt_vh():                                  #функция скнопки сохранения полей акта
    try:
        temp=last_index                 #временная переменная запомнить индекс выбранного элекмента
        for (i,j) in zip(('date','doc','kol'),(date_vh,doc_vh,kol_vh)): #записываем поля ввода
            akt_vh[temp][i]=j.get()                                     
        for (i,j,n) in zip(('otv','contr','kk','sk'),(var_otv,var_contr,var_kk,var_sk),(all_otv,all_control,all_control,all_sk)):
            akt_vh[temp][i]=n[j.get()]              #записываем выпадающие списки                  
        date_vh.config(fg='Black')                  #дата сохранилась, по этому снимаем выделение красным и делаем стандартный цвет
    except IndexError:
        showerror('Индекс!','Ошибка индекса!')
        traceback.print_exc(file=open('error.txt', 'w', encoding='utf-8'))
    except KeyError:
        traceback.print_exc(file=open('error.txt', 'w', encoding='utf-8'))

def up_akt_vh():                                #поднимает на единицу акт входного в списке листбокс
    try:
        global last_index
        temp=lbox.curselection()[0]
        akt_vh.insert(temp-1,akt_vh.pop(temp))
        re_num_vh()
        lbox.selection_set(temp-1)
        last_index=temp-1
    except IndexError:
        showerror('Индекс!','Ошибка индекса!')
        traceback.print_exc(file=open('error.txt', 'w', encoding='utf-8'))
        
def down_akt_vh():                              #опускает на единицу акт входного в списке листбокс
    try:
        global last_index
        temp=lbox.curselection()[0]
        akt_vh.insert(temp+1,akt_vh.pop(temp))
        re_num_vh()
        lbox.selection_set(temp+1)
        last_index=temp+1
    except IndexError:
        showerror('Индекс!','Ошибка индекса!')
        traceback.print_exc(file=open('error.txt', 'w', encoding='utf-8'))

def insert_vh(akt_vh,road_programm,road_db,data):
    try:
        insert_zhr_vh(akt_vh,road_programm,os.path.split(road_db)[0],*data)
        insert_akt_vh(akt_vh,road_programm,road_db,*data)
        showinfo(title='Успех!', message='Готово!')
    except:
        showerror('Ошибка!','Что-то пошло не так! Не удалось создать excell-файл.')
        traceback.print_exc(file=open('error.txt', 'w', encoding='utf-8'))
        
def insert_zhr(def_db,akt_vh,road_programm,road_db,data):
    try:
        road_db = os.path.split(road_db)[0]
        insert_zhr_ground(def_db,road_programm,road_db,*data)
        insert_zhr_izol(def_db,road_programm,road_db,*data)
        insert_zhr_ozhr(akt_vh,def_db,road_programm,road_db,*data)
        insert_zhr_svarki(def_db,road_programm,road_db,*data)
        insert_zhr_zam(def_db,road_programm,road_db,*data)
        showinfo(title='Успех!', message='Готово!')
    except:
        showerror('Ошибка!','Что-то пошло не так! Не удалось создать excell-файл.')
        traceback.print_exc(file=open('error.txt', 'w', encoding='utf-8'))

def insert_akts(akt_vh,road_programm,road_db,data):
    try:
        insert_main(akt_vh,road_programm,road_db,*data)
        insert_akt_all(akt_vh,road_programm,road_db,*data)
        showinfo(title='Успех!', message='Готово!')
    except:
        showerror('Ошибка!','Что-то пошло не так! Не удалось создать excell-файл.')
        traceback.print_exc(file=open('error.txt', 'w', encoding='utf-8'))
        
'''if 'Dmitr062' not in getpass.getuser() and 'LazarevDS' not in getpass.getuser():
    raise SystemExit'''

akt_vh=[]   #список с актами входного контроля
def_db=[]   #данные посекциям с дефектами получаемые из формы def_form
road_db=''
road_programm = os.getcwd()
last_index=None
with open("changelog.txt") as file_handler:
    about_programm = file_handler.readlines()
#-------------------------------------считываем из текстовых файлов с данными для выпадающих списков
all_otv={'':('','','')}             #для выпадающего списка с ответственными + пустая запись(если отсутствует)
all_control={'':('','','')}         #для выпадающего списка с контролирующими лицами + пустая запись(если отсутствует)
all_sk={'':('','','')}              #для выпадающего списка с строительным контролем + пустая запись(если отсутствует)
all_lkk={'':('','')} 
materials={'':('','','','','')}        #материалы для входного контроля

for (i,j) in zip(('otv.txt','control.txt','sk.txt','lkk.txt','materials.txt'),(all_otv,all_control,all_sk,all_lkk,materials)):      #Чтение ответственных из файлов с людьми
    try:
        f=open(os.getcwd()+'\\db\\'+i)
        for x in f:
            x = x.rstrip()
            j[x.split('#')[0]]=tuple(x.split('#'))
    finally:
        f.close()
    
#-------------------------------------Основное окно
main_win=Tk()
main_win.title('Дефекты')
#--------------------------------------Меню в шапке главного окна
db_open_status=False            #флаг, открыта ли БД
mainmenu=Menu(main_win)         #создаем меню
main_win.config(menu=mainmenu)
filemenu=Menu(mainmenu, tearoff=0)
filemenu.add_command(label='Открыть...',command=open_db)
filemenu.add_command(label='Сохранить',command=save_db)
filemenu.add_command(label='Сохранить как...',command=save_as_db)
filemenu.add_separator()
filemenu.add_command(label='Выход',command=sys.exit)
export_menu=Menu(mainmenu, tearoff=0)
export_menu.add_command(label='1. Входной контроль',command=lambda:insert_vh(akt_vh,road_programm,road_db,tuple(data_tube[i].get() for i in data)))
export_menu.add_command(label='2. Журналы',command=lambda:insert_zhr(def_db,akt_vh,road_programm,road_db,tuple(data_tube[i].get() for i in data)))
export_menu.add_command(label='3. Акты',command=lambda:insert_akts(def_db,road_programm,road_db,tuple(data_tube[i].get() for i in data)))

mainmenu.add_cascade(label='Файл', menu=filemenu)
mainmenu.add_cascade(label='Экспорт', menu=export_menu)
mainmenu.add_command(label='Справка',command=lambda:showinfo(title='О программе', message=about_programm))

#---------------------------------------1. Поля ввода основных полей Frame_1
frame_1 = Frame(main_win, relief=RIDGE, borderwidth=4) #область ввода начальных значений 
frame_1.grid(row=10,column=0)

Label(frame_1, text='1. Основное параметры').grid(row=1,column=1,columnspan=10,sticky=W)    #надпись в теле формы
Label(frame_1, text='Труба:').grid(row=2,column=1)
Label(frame_1, text='Километры:').grid(row=3,column=1)
Label(frame_1, text='Диаметр, мм:').grid(row=4,column=1)

data=('tube','km_start','km_finish','dy_tube')
data_tube={}
for i in data:
        data_tube[i]=Entry(frame_1,width=20)

data_tube[data[0]]['width']=40
data_tube[data[0]].grid(row=2,column=2,columnspan=2)
data_tube[data[1]].grid(row=3,column=2)
data_tube[data[2]].grid(row=3,column=3)
data_tube[data[3]].grid(row=4,column=2)
#--------------------------------------2. Кнопка открытия окна дефектов Frame_2
frame_2 = Frame(main_win, relief=RIDGE, borderwidth=4)
frame_2.grid(row=20,column=0,sticky=EW,pady=5)
Label(frame_2, text='2. Форма дефектов').pack(side=LEFT)
Button(frame_2,text='Тыц', command=lambda:def_main.def_form(def_db,all_otv,all_control,all_sk,all_lkk)).pack(fill=X)
#--------------------------------------3. Раздел формирования актов входного контроля Frame_3
frame_3 = Frame(main_win,relief=RIDGE, borderwidth=4)
frame_3.grid(row=30,column=0,ipadx=5)
Label(frame_3, text='3. Входной контроль').grid(row=0,column=0,columnspan=2)
lbox=Listbox(frame_3,height=15,exportselection=0)    #листбокс сдобавленными актами входного контроля
lbox.grid(row=1,column=0,columnspan=2,rowspan=10)
scroll=Scrollbar(frame_3,command=lbox.yview)
scroll.grid(row=1,column=2,rowspan=10,sticky=NS)
lbox.config(yscrollcommand=scroll.set)              #активация заполениями уже забитых форм в поля по щелчку в листбоксе
lbox.bind('<<ListboxSelect>>',update_box)

Button(frame_3,text='+',width=5,command=add_akt_vh).grid(row=0,column=4)
Button(frame_3,text='-',width=5,command=del_akt_vh).grid(row=0,column=5)
Button(frame_3,text='вверх',width=5,command=up_akt_vh).grid(row=11,column=0,sticky=EW)
Button(frame_3,text='вниз',width=5,command=down_akt_vh).grid(row=11,column=1,sticky=EW)
 
#--------------------------------------выпадающий список материалов и остальные поля в Frame_3
Label(frame_3, text='').grid(row=1,column=4,columnspan=2,sticky=W)
Label(frame_3, text='Дата').grid(row=2,column=4,columnspan=2,sticky=W)
Label(frame_3, text='Документация').grid(row=3,column=4,columnspan=2,sticky=W)
Label(frame_3, text='Количество').grid(row=4,column=4,columnspan=2,sticky=W)
var_mat=StringVar(value='')
OptionMenu(frame_3,var_mat,*materials).grid(row=0,column=6)
date_vh=Entry(frame_3,width=20)
doc_vh=Entry(frame_3,width=20)
kol_vh=Entry(frame_3,width=20)
date_vh.grid(row=2,column=6)
doc_vh.grid(row=3,column=6)
kol_vh.grid(row=4,column=6)
#------------------Меню ответственных по работам в Frame_3
Label(frame_3, text='Производитель работ:').grid(row=5,column=4,columnspan=2,sticky=W)
Label(frame_3, text='Контролирующее лицо:').grid(row=6,column=4,columnspan=2,sticky=W)
Label(frame_3, text='Контроль качества:').grid(row=7,column=4,columnspan=2,sticky=W)
Label(frame_3, text='СК:').grid(row=8,column=4,columnspan=2,sticky=W)

var_otv=StringVar(value='')
var_contr=StringVar(value='')
var_kk=StringVar(value='')
var_sk=StringVar(value='')

optm1 = OptionMenu(frame_3,var_otv,*all_otv)
optm1.grid(row=5,column=6)
optm1.config(width=16)
optm2 = OptionMenu(frame_3,var_contr,*all_control)
optm2.grid(row=6,column=6)
optm2.config(width=16)
optm3 = OptionMenu(frame_3,var_kk,*all_control)
optm3.grid(row=7,column=6)
optm3.config(width=16)
optm4 = OptionMenu(frame_3,var_sk,*all_sk)
optm4.grid(row=8,column=6)
optm4.config(width=16)

main_win.mainloop()






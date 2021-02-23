from datetime import *
import random,os
from tkinter import *
from tkinter.messagebox import showerror
    
def def_form(db_from_start,all_otv,all_control,all_sk,all_lkk):     #основное окно
    def add_s(def_sec):                   #Добавление дефектной секции (в db_from_start)
        db_from_start.append({'sec':def_sec,
                       'date':[None,None,None],
                       'dist':[0,None,None,None,None],
                       'km':0,
                       'dl_muft':0,
                       'zakl':'',
                       'rand_value':[None,None,None,None],
                       'defect':{}
                       })
        l_box.insert(END,'%d. %s'%(len(db_from_start),def_sec))
        for i in range(len(db_from_start)):
            db_from_start[i]['rand_value'][0]=random.uniform(0.25,0.6)     #Добавление рандомных велечин для котлована и изоляции
            db_from_start[i]['rand_value'][1]=random.uniform(2.5,3.8)      #Добавление рандомных велечин для котлована и изоляции
            db_from_start[i]['rand_value'][2]=random.uniform(2.5,3.2)      #ширина котлована
            db_from_start[i]['rand_value'][3]=random.uniform(2.5,3.0)      #глубина котлована
        def_ent_sec.delete(0,END)                                   #очищает строку ввода номера секции после введения(+)                                 

    def del_s():                              #Удаление секции с дефектом (из db_from_start)
        del(db_from_start[l_box.curselection()[0]])
        re_num_s()

    def re_num_s():                               #обновление нумерации секций в списке
        l_box.delete(0,END)                     #очищение виджета l_box  перед заполнением
        for i,j in enumerate(db_from_start):           #обновление нумерации
            l_box.insert(END,'%s. %s'%(i,db_from_start[i]['sec']))

    def update_lbox(a):                         #обновление информации в frame03 по щелчку в listbox
        try:
            temp=l_box.curselection()[0]
            label01['text']='Номер секции: '+db_from_start[temp]['sec'] #заголовок frame03
        except IndexError:pass
        def_ent_date.config(fg='Red')
        if db_from_start[temp]['date'][0]!=None:           #Если новая запись, то остается значени предыдщей записи, если есть - присваивается
            def_ent_date.delete(0,END)
            def_ent_date.insert(0,db_from_start[temp]['date'][0].strftime('%d.%m.%Y'))
            def_ent_date.config(fg='Black')
        def_ent_km.delete(0,END)
        def_ent_km.insert(0,db_from_start[temp]['km'])
        def_ent_dist.delete(0,END)
        def_ent_dist.insert(0,db_from_start[temp]['dist'][0])
        def_ent_dl_muft.delete(0,END)
        def_ent_dl_muft.insert(0,db_from_start[temp]['dl_muft'])
        def_ent_zakl.delete(0,END)
        def_ent_zakl.insert(0,db_from_start[temp]['zakl'])
        if 'type_base' in db_from_start[temp]:                   #Если новая запись, то остается значени предыдщей записи, если есть - присваивается
            var_type_base.set(db_from_start[temp]['type_base'])
        if 'otv' in db_from_start[temp]:                   #Если новая запись, то остается значени предыдщей записи, если есть - присваивается
            var_otv.set(db_from_start[temp]['otv'][0])
        if 'contr' in db_from_start[temp]:                 #Если новая запись, то остается значени предыдщей записи, если есть - присваивается
            var_contr.set(db_from_start[temp]['contr'][0])
        if 'sk' in db_from_start[temp]:                    #Если новая запись, то остается значени предыдщей записи, если есть - присваивается
            var_sk.set(db_from_start[temp]['sk'][0])
        if 'lkk' in db_from_start[temp]:                    #Если новая запись, то остается значени предыдщей записи, если есть - присваивается
            var_lkk.set(db_from_start[temp]['lkk'][0])
        if 'grnd_maker' in db_from_start[temp]:                    #Если новая запись, то остается значени предыдщей записи, если есть - присваивается
            var_grnd_maker.set(db_from_start[temp]['grnd_maker'][0])
        if 'grnd_contr' in db_from_start[temp]:                    #Если новая запись, то остается значени предыдщей записи, если есть - присваивается
            var_grnd_contr.set(db_from_start[temp]['grnd_contr'][0])       
        load_lbox_2()

    def save_s():                                                             #сохранение во временной БД (db_from_start)
        temp=l_box.curselection()[0]
        if def_ent_date.get()!='':
            try:
                temp_date=list(int(i) for i in (def_ent_date.get().split('.')))     #интерпритация и запись даты
                db_from_start[temp]['date'][0]=date(temp_date[2],temp_date[1],temp_date[0])#интерпритация и запись даты
                db_from_start[temp]['date'][1]=db_from_start[temp]['date'][0]
                if 'Муфта П1' == var_type_base.get():    #флаг, муфта П1 или нет
                    db_from_start[temp]['date'][1]=db_from_start[temp]['date'][0]+dday
                db_from_start[temp]['date'][2]=db_from_start[temp]['date'][1]+dday
                def_ent_date.config(fg='Black')
            except ValueError:showerror('Не сохранено!','Правильный формат:01.01.2019')
        db_from_start[temp]['km']=def_ent_km.get()                                 #запись километра
        temp_dist = def_ent_dist.get()
        if temp_dist != '':
            temp_dist = temp_dist.replace(',','.')
            try:
                result=db_from_start[temp]['dl_muft']/2
                db_from_start[temp]['dist'][0]=float(temp_dist)               #запись дистанции      
                db_from_start[temp]['dist'][1]=round(db_from_start[temp]['dist'][0]-db_from_start[temp]['rand_value'][0]-result,2)      #начало изоляиии
                db_from_start[temp]['dist'][2]=round(db_from_start[temp]['dist'][0]+db_from_start[temp]['rand_value'][0]+result,2)      #конец изоляции
                db_from_start[temp]['dist'][3]=round(db_from_start[temp]['dist'][1]-db_from_start[temp]['rand_value'][1]-result,2)      #начало котлована
                db_from_start[temp]['dist'][4]=round(db_from_start[temp]['dist'][2]+db_from_start[temp]['rand_value'][1]+result,2)      #конец котлована
            except ValueError:showerror('Не сохранено!','Правильный формат:12345.64')
        try:
            temp_dl_muft = (def_ent_dl_muft.get()).replace(',','.')
            db_from_start[temp]['dl_muft']=float(temp_dl_muft)                      #запись длины устанавливаемой муфты
        except ValueError:showerror('Не сохранено!','Правильный формат:1.5')
        db_from_start[temp]['zakl']=def_ent_zakl.get()
        db_from_start[temp]['type_base']=var_type_base.get()
        db_from_start[temp]['otv']=all_otv[var_otv.get()]                          #запись людей
        db_from_start[temp]['contr']=all_control[var_contr.get()]                  #запись людей
        db_from_start[temp]['sk']=all_sk[var_sk.get()]                             #запись людей
        db_from_start[temp]['lkk']=all_lkk[var_lkk.get()]                           #запись людей
        db_from_start[temp]['grnd_maker']=all_otv[var_grnd_maker.get()]                             #запись людей
        db_from_start[temp]['grnd_contr']=all_control[var_grnd_contr.get()] 
        
        
    def up_s():                 #переместить секцию вверх по списку lbox
        temp=l_box.curselection()[0]
        db_from_start.insert(temp-1,db_from_start.pop(temp))
        re_num_s()
        
    def down_s():               #переместить секцию вниз по списку lbox
        temp=l_box.curselection()[0]
        db_from_start.insert(temp+1,db_from_start.pop(temp))
        re_num_s()

    def add_d():                #Добавляет/обновляет дефекты в frame04
        curent_s=l_box.curselection()[0]
        if def_shablon_d['#'].get() not in db_from_start[curent_s]['defect']:
            db_from_start[curent_s]['defect'][def_shablon_d['#'].get()]={}
            l_box_2.insert(END,def_shablon_d['#'].get())
        for field in def_fields[1:]:
            db_from_start[curent_s]['defect'][def_shablon_d['#'].get()][field]=def_shablon_d[field].get()   
        for field in ('#','dist','dl','sh','gl'):
            def_shablon_d[field].delete(0,END)    

    def del_d():                #удаляет дефекты в frame04
        temp=l_box_2.curselection()
        del(db_from_start[l_box.curselection()[0]]['defect'][l_box_2.get(temp)])
        l_box_2.delete(temp)

    def update_lbox_2(a):       #обновление информации в frame04 по щелчку в listbox_2
        temp=l_box_2.get(l_box_2.curselection())
        for field in def_fields:
            def_shablon_d[field].delete(0,END)
            if field=='#':
                def_shablon_d[field].insert(0,temp)
            else:
                def_shablon_d[field].insert(0,db_from_start[l_box.curselection()[0]]['defect'][temp][field])

    def load_lbox_2():          #внесение в listbox_2 списка дефектов из бд по щелчку в listbox(1)
        l_box_2.delete(0,END)
        for field in def_fields:
            def_shablon_d[field].delete(0,END)
        if db_from_start[l_box.curselection()[0]]['defect']!={}:
            for i in db_from_start[l_box.curselection()[0]]['defect']:
                l_box_2.insert(END,i)

    dday=timedelta(days=1)
    def_fields=['#','dist','lab','dl','sh','gl','type']   #поля параметров дефектов frame05
    def_shablon_d={}                    #шаблон на дефекты
        
    form_def=Toplevel()
    form_def.title('Форма дефектов')
    #form_def.geometry("780x600") 
    #------------------Меню дефектов
    frame01_def = Frame(form_def, relief=RIDGE, borderwidth=2)
    frame01_def.grid(row=0,column=0,columnspan=2,sticky=W)
    Button(frame01_def,text='+',width=5,command=lambda:add_s(def_sec=def_ent_sec.get())).grid(row=0,column=0)
    Button(frame01_def,text='-',width=5,command=del_s).grid(row=0,column=1)
    Label(frame01_def, text='№ секции').grid(row=0,column=2)
    def_ent_sec=Entry(frame01_def,width=20)
    def_ent_sec.grid(row=0,column=3)

    frame02_def = Frame(form_def, relief=RIDGE, borderwidth=2)
    frame02_def.grid(row=1,column=0,sticky=NW,rowspan=2)
    l_box=Listbox(frame02_def,height=34,exportselection=0)
    l_box.grid(row=0,column=0,columnspan=2,sticky=NS)

    scroll_1=Scrollbar(frame02_def,command=l_box.yview)
    scroll_1.grid(row=0,column=2,sticky=NS)
    l_box.config(yscrollcommand=scroll_1.set)
    l_box.bind('<<ListboxSelect>>',update_lbox)

    Button(frame02_def,text='вверх',width=5,command=up_s).grid(row=1,column=0,sticky=EW)
    Button(frame02_def,text='вниз',width=5,command=down_s).grid(row=1,column=1,sticky=EW)
    #!!!-----------------!Frame03!
    frame03_def=Frame(form_def,relief=RIDGE, borderwidth=3)
    frame03_def.grid(row=1,column=1,sticky=NW,ipadx=5,columnspan=2)
    label01=Label(frame03_def, text='Номер секции:')
    label01.grid(row=20,column=0,columnspan=2,sticky=W)
    #------------------Дата дефекта в Frame03
    Label(frame03_def, text='Дата:').grid(row=30,column=0)
    def_ent_date=Entry(frame03_def,width=30)
    def_ent_date.grid(row=30,column=1)
    #------------------Километр дефекта в Frame03
    Label(frame03_def, text='Км:').grid(row=35,column=0)
    def_ent_km=Entry(frame03_def,width=30)
    def_ent_km.grid(row=35,column=1)
    #------------------Дистанция главного дефекта в Frame03
    Label(frame03_def, text='Дистанция, м.:').grid(row=36,column=0)
    def_ent_dist=Entry(frame03_def,width=30)
    def_ent_dist.grid(row=36,column=1)
    #------------------Дистанция главного дефекта в Frame03
    Label(frame03_def, text='Длина муфты, м:').grid(row=37,column=0)
    def_ent_dl_muft=Entry(frame03_def,width=30)
    def_ent_dl_muft.grid(row=37,column=1)
    #------------------Номер заключения в Frame03 22.03.2020
    Label(frame03_def, text='Номер закл., ВИК:').grid(row=40,column=0)
    def_ent_zakl=Entry(frame03_def,width=30)
    def_ent_zakl.grid(row=40,column=1)
    #------------------Тип дефекта в Frame03 22.03.2020
    Label(frame03_def, text='Тип устранения:').grid(row=45,column=0)
    var_type_base=StringVar(value='')
    OptionMenu(frame03_def,var_type_base,'Муфта П1','Муфта П2','Муфта П3','Муфта П4','Муфта П5У','Муфта П6','Муфта П7','Муфта П8','Муфта П9','Муфта П10','Шлифовка').grid(row=45,column=1)
    #------------------Меню ответственных по работам в Frame03
    Label(frame03_def, text='Производитель работ:').grid(row=50,column=0)
    Label(frame03_def, text='Контролирующее лицо:').grid(row=60,column=0)
    Label(frame03_def, text='СК:').grid(row=70,column=0)
    Label(frame03_def, text='Лаборатория:').grid(row=80,column=0)
    Label(frame03_def, text='Земляные, производитель:').grid(row=90,column=0)
    Label(frame03_def, text='Земляные, контроль:').grid(row=100,column=0)
    var_otv=StringVar(value='')
    var_contr=StringVar(value='')
    var_sk=StringVar(value='')
    var_lkk=StringVar(value='')
    var_grnd_maker = StringVar(value='')
    var_grnd_contr = StringVar(value='')
    OptionMenu(frame03_def,var_otv,*all_otv).grid(row=50,column=1)
    OptionMenu(frame03_def,var_contr,*all_control).grid(row=60,column=1)
    OptionMenu(frame03_def,var_sk,*all_sk).grid(row=70,column=1)
    OptionMenu(frame03_def,var_lkk,*all_lkk).grid(row=80,column=1)
    OptionMenu(frame03_def,var_grnd_maker,*all_otv).grid(row=90,column=1)
    OptionMenu(frame03_def,var_grnd_contr,*all_control).grid(row=100,column=1)
    Button(frame03_def,text='Сохранить',command=save_s).grid(row=200,column=0)
    #!!!-----------------!Frame04!
    frame04_def=Frame(form_def,relief=RIDGE, borderwidth=3)
    frame04_def.grid(row=2,column=1,sticky=NW)
    #--------------------Лист бокс дефектов секции
    l_box_2=Listbox(frame04_def,height=12)
    l_box_2.pack(side=LEFT)
    scroll_2=Scrollbar(frame04_def,command=l_box_2.yview)
    scroll_2.pack(side=LEFT,fill=Y)
    l_box_2.config(yscrollcommand=scroll_2.set)
    l_box_2.bind('<<ListboxSelect>>',update_lbox_2)
    #!!!-----------------!Frame05!
    frame05_def=Frame(form_def,relief=RIDGE, borderwidth=3)
    frame05_def.grid(row=2,column=2,sticky=NW,ipadx=30,ipady=14)
    #------------------Параметры конкретного дефекта в Frame05
    Button(frame05_def,text='+',width=5,command=add_d).grid(row=0,column=0)
    Button(frame05_def,text='-',width=5,command=del_d).grid(row=0,column=1)
    Label(frame05_def, text='Номер дефекта:').grid(row=10,column=0,columnspan=2)
    Label(frame05_def, text='Дистанция:').grid(row=20,column=0,columnspan=2)
    Label(frame05_def, text='Описание дефекта:').grid(row=30,column=0,columnspan=2)
    Label(frame05_def, text='Длина:').grid(row=40,column=0,columnspan=2)
    Label(frame05_def, text='Ширина:').grid(row=50,column=0,columnspan=2)
    Label(frame05_def, text='Глубина:').grid(row=60,column=0,columnspan=2)
    Label(frame05_def, text='Метод устранения:').grid(row=70,column=0,columnspan=2)
    j=10
    for i in def_fields:
        def_shablon_d[i]=Entry(frame05_def,width=30)
        def_shablon_d[i].grid(row=j,column=3)
        j+=10
    re_num_s() #сразу запускаем функцию обновления списка секций



if __name__=='__main__':
    db_from_start=[]                    #главный список, в качестве БД
    all_otv={'':('','','')}             #для выпадающего списка с ответственными + пустая запись(если отсутствует)
    all_control={'':('','','')}         #для выпадающего списка с контролирующими лицами + пустая запись(если отсутствует)
    all_sk={'':('','','')}              #для выпадающего списка с строительным контролем + пустая запись(если отсутствует)
    all_lkk={'':('','','')}
    for (i,j) in zip(('otv.txt','control.txt','sk.txt','lkk.txt'),(all_otv,all_control,all_sk,all_lkk)):      #Чтение ответственных из файлов с людьми
        try:
            f=open(os.getcwd()[:-5]+'db\\'+i)
            for x in f:
                j[x.split('#')[0]]=x.split('#')
        finally:
            f.close()
    def_form(db_from_start,all_otv,all_control,all_sk,all_lkk)

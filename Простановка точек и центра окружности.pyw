#-------------------------------------------------------------------------------
# Name:        Простановка точек и центра окружности
# version:     v0.1.0.0
#
# Author:      dimak222
#
# Created:     18.08.2022
# Copyright:   (c) dimak222 2022
# Licence:     No
#-------------------------------------------------------------------------------

title = "Простановка точек и центра окружности"

def DoubleExe():# проверка на уже запущеное приложение, с отключённым консольным окном "CREATE_NO_WINDOW"

    import subprocess # модуль вывода запущеных процессов
    from sys import exit # для выхода из приложения без ошибки

    CREATE_NO_WINDOW = 0x08000000 # отключённое консольное окно
    processes = subprocess.Popen('tasklist', stdin=subprocess.PIPE, stderr=subprocess.PIPE, stdout=subprocess.PIPE, creationflags=CREATE_NO_WINDOW).communicate()[0] # список всех процессов
    processes = processes.decode('cp866') # декодировка списка

    if str(processes).count(title[0:25]) > 2: # если найдено название программы (два процесса) с ограничением в 25 символов
        Message("Приложение уже запущено!") # сообщение, поверх всех окон и с автоматическим закрытием
        exit() # выходим из програмы

def KompasAPI(): # подключение API компаса

    from win32com.client import Dispatch, gencache # библиотека API Windows
    from sys import exit # для выхода из приложения без ошибки

    try: # попытаться подключиться к КОМПАСу

        global KompasAPI5 # значение делаем глобальным
        global iKompasObject # значение делаем глобальным
        global KompasAPI7 # значение делаем глобальным
        global iApplication # значение делаем глобальным
        global iKompasDocument # значение делаем глобальным
        global iDocument2D # значение делаем глобальным
        global iKompasDocument2D1 # значение делаем глобальным
        global iViews # значение делаем глобальным

        KompasConst3D = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants # константа 3D документов
        KompasConst2D = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа 2D документов
        KompasConst = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа для скрытия вопросов перестроения

        KompasAPI5 = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0, 1, 0) # API5 КОМПАСа
        iKompasObject = Dispatch('Kompas.Application.5', None, KompasAPI5.KompasObject.CLSID) # интерфейс API КОМПАС

        KompasAPI7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0) # API7 КОМПАСа
        iApplication = Dispatch('Kompas.Application.7') # интерфейс приложения КОМПАС-3D.

        iKompasDocument = iApplication.ActiveDocument # делаем активный открытый документ

        document_type() # определение типа документа

        iDocument2D = iKompasObject.ActiveDocument2D() # указатель на интерфейс текущего графического документа
        iKompasDocument2D = KompasAPI7.IKompasDocument2D(iKompasDocument) # базовый класс графических документов КОМПАС
        iKompasDocument2D1 = KompasAPI7.IKompasDocument2D1(iKompasDocument) # дополнительный интерфейс IKompasDocument2D

        iViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager # менеджер видов и слоев документа
        iViews = iViewsAndLayersManager.Views # коллекция видов

        if iApplication.Visible == False: # если компас невидимый
            iApplication.Visible = True # сделать КОМПАС-3D видемым

    except: # если не получилось подключиться к КОМПАСу

        Message("КОМПАС-3D не найден!\nУстановите или переустановите КОМПАС-3D!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из програмы

def Message(text = "Ошибка!", counter = 4): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    from threading import Thread # библиотека потоков
    import time # модуль времени

    def Message_Thread(text, counter): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

        import tkinter.messagebox as mb # окно с сообщением
        import tkinter as tk # модуль окон

        if counter == 0: # время до закрытия окна (если 0)
            counter = 1 # закрытие через 1 сек
        window_msg = tk.Tk() # создание окна
        window_msg.iconbitmap(default = Resource_path("cat.ico")) # значёк программы
        window_msg.attributes("-topmost",True) # окно поверх всех окон
        window_msg.withdraw() # скрываем окно "невидимое"
        time = counter * 1000 # время в милисекундах
        window_msg.after(time, window_msg.destroy) # закрытие окна через n милисекунд
        if mb.showinfo(title, text, parent = window_msg) == "": # информационное окно закрытое по времени
            pass
        else:
            window_msg.destroy() # окно закрыто по кнопке
        window_msg.mainloop() # отображение окна

    msg_th = Thread(target = Message_Thread, args = (text, counter)) # запуск определения положения мышки в отдельном потоке
    msg_th.start() # запуск потока

    msg_th.join() # ждать завершения процесса, иначе может закрыться следующие окно

def Kompas_message(text): # сообщение в окне компаса если он открыт

    if iApplication.Visible == True: # если компас видимый
        iApplication.MessageBoxEx(text, 'Message:', 64)

def document_type(): # определение типа документа

    from sys import exit # для выхода из приложения без ошибки

    if iKompasDocument == None or iKompasDocument.DocumentType not in (1, 2): # если не открыт док. или не 2D док., выдать сообщение (1-чертёж; 2- фрагмент; 3-СП; 4-модель; 5-СБ; 6-текст. док.; 7-тех. СБ;)

        msg = "Откройте чертёж!"
        Kompas_message(msg) # сообщение в компасе
        exit() # завершаем макрос

def iSelectedObjects_processing(iSelectedObject): # обработка выбраных объектов

    if iSelectedObject.DrawingObjectType == 123: # если вид (тип графических объектов (123 - вид., см DrawingObjectTypeEnum))
        View_processing(iSelectedObject) # обработка всего вида

    else:
        R_X_Y_obg(iSelectedObject) # получение координат и указатель на объект

def View_processing(iView): # обработка всего вида

    if iView.Visible: # если вид видимый, используем его

        iDrawingContainer = KompasAPI7.IDrawingContainer(iView) # интерфейс контейнера объектов вида графического документа

        for obg in (2, 3, 36, 32, 34, 35): # перебор всех объектов
            iObjects = iDrawingContainer.Objects(obg) # получить массив объектов коллекции

            if iObjects: # если есть объекты
                for OBJ in iObjects: # обработать каждый объект
                    R_X_Y_obg(OBJ) # получение координат и указатель на объект

def R_X_Y_obg(iSelectedObject): # получение координат и указатель на объект

    iReference = iSelectedObject.Reference # указатель объекта

    iNumber = iDocument2D.ksGetViewNumber(iReference) # номер вида по выделеному объекту
    iView = iViews.ViewByNumber(iNumber) # вид, заданный по номеру
    Scale = iView.Scale # масштаб вида

    Type_obj = iSelectedObject.DrawingObjectType # тип графических объектов

    if Type_obj in (2, 3, 36): # Обработка объектов по типу объекта (2 - окружность; 3 - дуга; 36 - многоугольник, см. DrawingObjectTypeEnum)
        R = iSelectedObject.Radius # радиус окружности
        R = R * Scale # радиус окружности умноженый на масштаб
        X = iSelectedObject.Xc # координата центра окружности по оси Х
        Y = iSelectedObject.Yc # координата центра окружности по оси Х

        BD.append([R, X, Y, iSelectedObject, iView]) # добавляем запись координат в список (радиус/размер, положение X, положение Y, объект, вид)

    elif Type_obj in (32, 34): # Обработка объектов по типу объекта (32 - элипс; 34 - дуга элипса; 35 - прямоугольник, см. DrawingObjectTypeEnum)

        A = iSelectedObject.SemiAxisA # длина полуосей
        B = iSelectedObject.SemiAxisB # длина полуосей
        X = iSelectedObject.Xc # координата центра окружности по оси Х
        Y = iSelectedObject.Yc # координата центра окружности по оси Х

        BD.append([max(A, B), X, Y, iSelectedObject, iView]) # добавляем запись координат в список (радиус/размер, положение X, положение Y, объект, вид)

    elif Type_obj == 35: # Обработка объектов по типу объекта (35 - прямоугольник, см. DrawingObjectTypeEnum)

        A = iSelectedObject.Height # длина полуосей
        B = iSelectedObject.Width # длина полуосей
        X = iSelectedObject.X # координата центра окружности по оси Х
        Y = iSelectedObject.Y # координата центра окружности по оси Х

        BD.append([max(A, B), X, Y, iSelectedObject, iView]) # добавляем запись координат в список (радиус/размер, положение X, положение Y, объект, вид)

def list_processing(): # обработка списка (удаление лишнего и сортировка отв. от большего к меньшему)

    from sys import exit # для выхода из приложения без ошибки

    if len(BD) >= 1: # если в списке больше 1 объекта

        print(BD)
        BD.sort() # сортируем его по увеличению радиусов

        removing_duplicates(BD) # удаление дубликатов объектов по их положению X и Y (перебор списка)

    else: # если нет объекта в списке
        Kompas_message("Не найдены объекты для создания обозначений центра!") # сообщение в окне компаса если он открыт
        exit()

def removing_duplicates(BD): # удаление дубликатов объектов по их положению X и Y (перебор списка)

    for n in range(len(BD)): # перебор значений из диапазона количества объектов

        m = n + 1 # для проверки последующего объектов

        while m < len(BD): # цикл проверки последующих объектов

            if BD[n][1] == BD[m][1] and BD[n][2] == BD[m][2]: # если координата X и Y объекта n совпадает с координатами X и Y объекта m
                del BD[n] # удаляем первый объект (меньший диаметр)

            else: # если не совпало пробуем следующий элемент
                m = m + 1  # следующий элемент на обработку

def Centre(x, y, R, obj, iView): # простановка центров окружностей (радиус/размер, положение X, положение Y, объект, вид)

    iSymbols2DContainer = KompasAPI7.ISymbols2DContainer(iView) # интерфейс контейнера условных обозначений
    iCentreMarkers = iSymbols2DContainer.CentreMarkers # коллекция обозначений центра
    iCentreMarker = iCentreMarkers.Add() # создать новый элемент и добавить его в коллекцию

    iCentreMarker.X = x # координата точки x делённая на масштаб
    iCentreMarker.Y = y # координата точки y
##    iCentreMarker.Angle = 0.0 # угол наклона обозначения центра
    iCentreMarker.SignType = 2 # тип обозначения центра: -1 - неизвестный 0 - маленький крестик; 1 - одна ось; 2 - две оси
    iCentreMarker.SetSemiAxisLength(0, R) # длина 1-ой полуоси
    iCentreMarker.SetSemiAxisLength(1, R) # длина 2-ой полуоси
    iCentreMarker.SetSemiAxisLength(2, R) # длина 3-ей полуоси
    iCentreMarker.SetSemiAxisLength(3, R) # длина 4-ой полуоси

    iCentreMarker.BaseObject = obj # ассоциировать данный объект с другим объектом

    iAxisLineParam = KompasAPI7.IAxisLineParam(iCentreMarker) # параметры осевой линии или обозначения центра
    iAxisLineParam.JutLength = 5.0 # выступание осевой линии
    iAxisLineParam.DottedLength = 1.5 # длина пунктира
    iAxisLineParam.Interval = 1.5 # промежуток
    iAxisLineParam.AutoDetectedDash = True # ручное/автоматическое задание длины штриха
    iAxisLineParam.DashMaxLength = 105.0 # максимальная длина штриха
    iAxisLineParam.JutLengthModify = False # использовать параметр Выступание осевой из настроек документа
    iAxisLineParam.DottedLengthModify = False # использовать параметр Длина пунктира из настроек документа
    iAxisLineParam.IntervalModify = False # использовать параметр Промежуток из настроек документа
    iAxisLineParam.AutoDetectedDashModify = False # использовать параметр Задание длины штриха из настроек документа
    iAxisLineParam.DashMaxLengthModify = False # использовать параметр Максимальная длина штриха из настроек документа
    iCentreMarker.Update() # приминить свойства

def Point(x, y, obj, iView): # простановка точек в окружности (радиус/размер, положение X, положение Y, объект, вид)

    iDrawingContainer = KompasAPI7.IDrawingContainer(iView) # интерфейс контейнера объектов вида графического документа
    iPoints = iDrawingContainer.Points # коллекция точек
    iPoint = iPoints.Add() # создать новый элемент и добавить его в коллекцию
    iPoint.X = x # координата точки x
    iPoint.Y = y # координата точки y
    iPoint.Angle = 0.0 # угол наклона для точки со стрелкой
    iPoint.Style = 1 # аннотационные символы (см. ksAnnotationSymbolEnum и ksAnnotativeTerminatorSignEnum)
    iPoint.Update() # приминить свойства

    iDrawingObject1 = KompasAPI7.IDrawingObject1(iPoint) # дополнительный интерфейс для графических объектов
    iParametriticConstraint = iDrawingObject1.NewConstraint() # создатём новое ограничение
    iParametriticConstraint.ConstraintType = 22 # тип ограничения (см. ksConstraintTypeEnum)
    iParametriticConstraint.Index = 0 # индекс точки на объекте (начинается с 0, у дуги и окружности 0-центр)
    iParametriticConstraint.Partner = obj # второй объект или массив SAFEARRAY объектов для установки ограничения
    iParametriticConstraint.PartnerIndex = 0 # индекс точки на втором объекте (начинается с 0, у дуги и окружности 0-центр)
    iParametriticConstraint.Create() # cоздать ограничение в модели

def Creating(): # создание центров и точек

    list_processing() # обработка списка (удаление лишнего и сортировка отв. от большего к меньшему)

    iGroup = iDocument2D.ksNewGroup(1) # создаём обозначения центров во временной группе

    for k in range(len(BD)): # строим обозначения центров во все объекты со списка
        Centre(BD[k][1], BD[k][2], BD[k][0], BD[k][3], BD[k][4]) # простановка центров окружностей (радиус/размер, положение X, положение Y, объект, вид)
        Point(BD[k][1], BD[k][2], BD[k][3], BD[k][4]) # простановка точек в окружности (радиус/размер, положение X, положение Y, объект, вид)

    iDocument2D.ksEndGroup # завершить создание группы
    iDocument2D.ksStoreTmpGroup(iGroup) # вставляем группу в чертёж

#-------------------------------------------------------------------------------

BD = [] # список для сбора координат и радиусов окружностей

DoubleExe() # проверка на уже запущеное приложение, с отключённым консольным окном "CREATE_NO_WINDOW"

KompasAPI() # подключение API компаса

iSelectionManager = iKompasDocument2D1.SelectionManager # менеджер выделенных объектов
iSelectedObjects = iSelectionManager.SelectedObjects # массив выделенных объектов в виде SAFEARRAY | VT_DISPATCH

if iSelectedObjects != None: # если выделены объекты

    if isinstance(iSelectedObjects, tuple): # если выбрано несколько объектов (кортеж объектов)
        for iSelectedObject in iSelectedObjects: # перебор всех выделеных объектов
            iSelectedObjects_processing(iSelectedObject) # обработка выбраных объектов

    else: # выбран один объект
        iSelectedObjects_processing(iSelectedObjects) # обработка выбраных объектов

    Creating() # создание центров и точек

    Kompas_message("Обозначения центров окружностей проставлены") # сообщение в окне компаса если он открыт

else: # если ничего не выбрано, выводим сообщение
    Kompas_message("Нет выделенных объектов или вида!") # сообщение в окне компаса если он открыт

#-------------------------------------------------------------------------------
# Name:        Простановка точек и центра окружности
# version:     v0.6.0.0
#
# Author:      dimak222
#
# Created:     18.08.2022
# Copyright:   (c) dimak222 2022
# Licence:     No
#-------------------------------------------------------------------------------

title = "Простановка точек и центра окружности"

#------------------------------Настройки!---------------------------------------
All_iView = True # обрабатывать ли все видимые виды если ничего не выделоенно (True - да, False - нет)

Centre_var = True # ставить ли обозначения центров (True - да, False - нет)
Associate_Centre = True # параметризовать ли обозначения центров (True - да, False - нет)
Centre_repeat = False # повторять ли простановку уже выделеных обозначений центрров (True - да, False - нет)

Point_var = True # ставить ли точки (True - да, False - нет)
Associate_Point = True # параметризовать ли точки (True - да, False - нет)
Point_repeat = False # повторять ли простановку уже выделеных точек (True - да, False - нет)

Delta = 0.01 #(мм). максимально допустимая разница координат центров/точек (если разница координат больше этой величины, центры/точки будет ставиться на каждый объект (0 - ставит точки на каждый объект))
Circle = True # обрабатывать ли окружности (True - да, False - нет)
Arc = True # обрабатывать ли дуги окружностей (True - да, False - нет)
Ellipse = True # обрабатывать ли эллипсы (True - да, False - нет)
Arc_ellipse = True # обрабатывать ли дуги эллипсов (True - да, False - нет)
RegularPolygon = True # обрабатывать ли многоугольники (True - да, False - нет)
Rectangle = True # обрабатывать ли прямоугольники (True - да, False - нет)

Style = "1-25" # стили линий объектов, "1-25" - все стили; "1-5,7,9" - стиль с 1-го по 5-й, 7-ой, 9-й; "1" - только основная, см. ksCurveStyleEnum
# системные стили линии:
#1 - основная,
#2 - тонкая,
#3 - осевая,
#4 - штриховая,
#5 - для линии обрыва
#6 - вспомогательная,
#7 - утолщенная,
#8 - пунктир 2,
#9 - штриховая осн.
#10 - осевая осн.
#11 - тонкая линия, включаемая в штриховку
#12 - ISO 02 штриховая линия,
#13 - ISO 03 штриховая линия (дл. пробел),
#14 - ISO 04 штрихпунктирная линия (дл. штрих),
#15 - ISO 05 штрихпунктирная линия (дл. штрих 2 пунктира),
#16 - ISO 06 штрихпунктирная линия (дл. штрих 3 пунктира),
#17 - ISO 07 пунктирная линия,
#18 - ISO 08 штрихпунктирная линия (дл. и кор. штрихи),
#19 - ISO 09 штрихпунктирная линия (дл. и 2 кор. штриха),
#20 - ISO 10 штрихпунктирная линия,
#21 - ISO 11 штрихпунктирная линия (2 штриха),
#22 - ISO 12 штрихпунктирная линия (2 пунктира),
#23 - ISO 13 штрихпунктирная линия (3 пунктира),
#24 - ISO 14 штрихпунктирная линия (2 штриха 2 пунктира),
#25 - ISO 15 штрихпунктирная линия (2 штриха 3 пунктира).
#-------------------------------------------------------------------------------

def DoubleExe():# проверка на уже запущеное приложение, с отключённым консольным окном "CREATE_NO_WINDOW"

    import subprocess # модуль вывода запущеных процессов
    from sys import exit # для выхода из приложения без ошибки

    CREATE_NO_WINDOW = 0x08000000 # отключённое консольное окно
    processes = subprocess.Popen('tasklist', stdin=subprocess.PIPE, stderr=subprocess.PIPE, stdout=subprocess.PIPE, creationflags=CREATE_NO_WINDOW).communicate()[0] # список всех процессов
    processes = processes.decode('cp866') # декодировка списка

    if str(processes).count(title[0:25]) > 2: # если найдено название программы (два процесса) с ограничением в 25 символов
        Message("Приложение уже запущено!") # сообщение, поверх всех окон и с автоматическим закрытием
        exit() # выходим из програмы

def Сhecking_settings(): # проверка правильности введённых значений

    from sys import exit # для выхода из приложения без ошибки

    if Circle == Arc == Ellipse == Arc_ellipse == RegularPolygon == Rectangle == False: # если все объекты выключены
        Message("Не заданы типы объектов для обработки!") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
        exit() # выходим из програмы

    if Centre_var == Point_var == False:
        Message("Не задана хотя бы одна простановка (центра/точки)!") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
        exit() # выходим из програмы

def KompasAPI(): # подключение API компаса

    from win32com.client import Dispatch, gencache # библиотека API Windows
    from sys import exit # для выхода из приложения без ошибки

    def document_type(): # определение типа документа

        from sys import exit # для выхода из приложения без ошибки

        if iKompasDocument == None or iKompasDocument.DocumentType not in (1, 2): # если не открыт док. или не 2D док., выдать сообщение (1-чертёж; 2- фрагмент; 3-СП; 4-модель; 5-СБ; 6-текст. док.; 7-тех. СБ;)

            msg = "Откройте чертёж!"
            Kompas_message(msg) # сообщение в компасе
            exit() # завершаем макрос

    try: # попытаться подключиться к КОМПАСу

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

    def Resource_path(relative_path): # для сохранения картинки внутри exe файла

        import os # работа с файовой системой

        try: # попытаться определить путь к папке
            base_path = sys._MEIPASS # путь к временной папки PyInstaller

        except Exception: # если ошибка
            base_path = os.path.abspath(".") # абсолютный путь

        return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

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

    msg_th = Thread(target = Message_Thread, args = (text, counter)) # запуск окна в отдельном потоке
    msg_th.start() # запуск потока

    msg_th.join() # ждать завершения процесса, иначе может закрыться следующие окно

def Kompas_message(text): # сообщение в окне компаса если он открыт

    if iApplication.Visible == True: # если компас видимый
        iApplication.MessageBoxEx(text, 'Message:', 64)

def iSelectedObjects_processing(iSelectedObject): # обработка выбраных объектов

    if iSelectedObject.DrawingObjectType == 123: # если вид (тип графических объектов (123 - вид., см DrawingObjectTypeEnum))
        View_processing(iSelectedObject) # обработка всего вида

    else:
        R_X_Y_obg(iSelectedObject) # получение координат, указатель на объект и указатенль на вид

def View_processing(iView): # обработка всего вида

    if iView.Visible: # если вид видимый, используем его

        iDrawingContainer = KompasAPI7.IDrawingContainer(iView) # интерфейс контейнера объектов вида графического документа

        for obg in (2, 3, 36, 32, 34, 35, 42, 5): # перебор всех объектов (2 - окружность; 3 - дуга; 36 - многоугольник; 32 - элипс; 34 - дуга элипса; 35 - прямоугольник; 42 - обозначение центра; 5 - точка, см. DrawingObjectTypeEnum)
            iObjects = iDrawingContainer.Objects(obg) # получить массив объектов коллекции

            if iObjects: # если есть объекты
                for OBJ in iObjects: # обработать каждый объект
                    R_X_Y_obg(OBJ) # получение координат, указатель на объект и указатенль на вид

def R_X_Y_obg(iSelectedObject): # получение координат, указатель на объект и указатенль на вид

    import math # модуль математических фенкций

    def iteration(numbers): # обработка перечисленых цифр

        from sys import exit # для выхода из приложения без ошибки

        list_numbers = [] # список цифр

        numbers = numbers.split(",") # разделяем строку по ","

        for n in numbers: # обработка каждого элемента в списке

            if n.find("-") != -1: # если в элементе найден знак "-" обрапботать его
                n = n.split("-") # разделяем строку по "-"

                list_numbers = list_numbers + list(range(int(n[0]), int(n[1])+1)) # добавляем к списку цифр список целых числ из диапазона

            else: # без знака "-"
                if n is int and n != "0": # если значение число и не "0"
                    list_numbers.append(int(n)) # добавляем целое число к списку цифр

                else: # введены неправельные значения
                    Message("Введите правильное значение в Style!") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
                    exit() # выходим из програмы

        return list_numbers # выводим список цифр

    iReference = iSelectedObject.Reference # указатель объекта

    iNumber = iDocument2D.ksGetViewNumber(iReference) # номер вида по выделеному объекту
    iView = iViews.ViewByNumber(iNumber) # вид, заданный по номеру
    Scale = iView.Scale # масштаб вида

    Type_obj = iSelectedObject.DrawingObjectType # тип графических объектов

    if Type_obj == 2 and Circle and iSelectedObject.Style in iteration(Style): # обработка объектов по типу объекта и стилю (2 - окружность; 3 - дуга; 36 - многоугольник, см. DrawingObjectTypeEnum)
        R = iSelectedObject.Radius # радиус окружности
        R = R * Scale # радиус окружности умноженый на масштаб
        X = iSelectedObject.Xc # координата центра окружности по оси Х
        Y = iSelectedObject.Yc # координата центра окружности по оси Х

        list_Obg.append([R, X, Y, iSelectedObject, iView]) # добавляем запись координат в список

    elif Type_obj == 3 and Arc and iSelectedObject.Style in iteration(Style): # обработка объектов по типу объекта и стилю (2 - окружность; 3 - дуга; 36 - многоугольник, см. DrawingObjectTypeEnum)
        R = iSelectedObject.Radius # радиус окружности
        R = R * Scale # радиус окружности умноженый на масштаб
        X = iSelectedObject.Xc # координата центра окружности по оси Х
        Y = iSelectedObject.Yc # координата центра окружности по оси Х

        list_Obg.append([R, X, Y, iSelectedObject, iView]) # добавляем запись координат в список

    elif Type_obj == 36 and RegularPolygon and iSelectedObject.Style in iteration(Style): # обработка объектов по типу объекта и стилю (2 - окружность; 3 - дуга; 36 - многоугольник, см. DrawingObjectTypeEnum)
        R = iSelectedObject.Radius # радиус окружности
        R = R * Scale # радиус окружности умноженый на масштаб
        X = iSelectedObject.Xc # координата центра окружности по оси Х
        Y = iSelectedObject.Yc # координата центра окружности по оси Х

        list_Obg.append([R, X, Y, iSelectedObject, iView]) # добавляем запись координат в список

    elif Type_obj == 32 and Ellipse and iSelectedObject.Style in iteration(Style): # обработка объектов по типу объекта и стилю (32 - элипс; 34 - дуга элипса; 35 - прямоугольник, см. DrawingObjectTypeEnum)

        A = iSelectedObject.SemiAxisA # длина полуосей
        B = iSelectedObject.SemiAxisB # длина полуосей
        X = iSelectedObject.Xc # координата центра окружности по оси Х
        Y = iSelectedObject.Yc # координата центра окружности по оси Х

        list_Obg.append([max(A, B), X, Y, iSelectedObject, iView]) # добавляем запись координат в список (радиус/размер, положение X, положение Y, объект, вид)

    elif Type_obj == 34 and Arc_ellipse and iSelectedObject.Style in iteration(Style): # обработка объектов по типу объекта и стилю (32 - элипс; 34 - дуга элипса; 35 - прямоугольник, см. DrawingObjectTypeEnum)

        A = iSelectedObject.SemiAxisA # длина полуосей
        B = iSelectedObject.SemiAxisB # длина полуосей
        X = iSelectedObject.Xc # координата центра окружности по оси Х
        Y = iSelectedObject.Yc # координата центра окружности по оси Х

        list_Obg.append([max(A, B), X, Y, iSelectedObject, iView]) # добавляем запись координат в список (радиус/размер, положение X, положение Y, объект, вид)

    elif Type_obj == 35 and Rectangle and iSelectedObject.Style in iteration(Style): # обработка объектов по типу объекта и стилю (35 - прямоугольник, см. DrawingObjectTypeEnum)

        α = iSelectedObject.Angle # угол прямоугольника
        A = iSelectedObject.Height # длина полуосей
        B = iSelectedObject.Width # длина полуосей
        X = iSelectedObject.X # координата центра окружности по оси Х
        Y = iSelectedObject.Y # координата центра окружности по оси Х

        X4 = X + B*math.cos(α*math.pi/180) - A*math.sin(α*math.pi/180) # положение противоположной точки прямоугольника
        Xc = (X + X4)/2 # положение центра прямоугольника

        Y3 = Y + B*math.sin(α*math.pi/180) + A*math.cos(α*math.pi/180) # положение противоположной точки прямоугольника
        Yc = (Y + Y3)/2 # положение центра прямоугольника

        list_Obg.append([max(A, B), Xc, Yc, iSelectedObject, iView]) # добавляем запись координат в список (радиус/размер, положение X, положение Y, объект, вид)

    elif Type_obj == 42 and Centre_repeat == False: # обработка объектов по типу объекта и стилю (42 - обозначение центра, см. DrawingObjectTypeEnum)
        X = iSelectedObject.X # координата центра окружности по оси Х
        Y = iSelectedObject.Y # координата центра окружности по оси Х

        list_Centre.append([0, X, Y, iSelectedObject, iView]) # добавляем запись координат в список (радиус/размер, положение X, положение Y, объект, вид)

    elif Type_obj == 5 and Circle and Point_repeat == False: # обработка объектов по типу объекта и стилю (5 - точка, см. DrawingObjectTypeEnum)
        X = iSelectedObject.X # координата центра окружности по оси Х
        Y = iSelectedObject.Y # координата центра окружности по оси Х

        list_Point.append([0, X, Y, iSelectedObject, iView]) # добавляем запись координат в список (радиус/размер, положение X, положение Y, объект, вид)

def Creating(): # создание центров и точек

    def list_processing(): # обработка списка (удаление лишнего и сортировка отв. от большего к меньшему)

        from sys import exit # для выхода из приложения без ошибки

        def removing_duplicates(list_Obg): # удаление дубликатов объектов по их положению X и Y (перебор списка)

            for n in range(len(list_Obg)): # перебор значений из диапазона количества объектов

                m = n + 1 # для проверки последующего объектов

                while m < len(list_Obg): # цикл проверки последующих объектов

                    if Delta != 0: # если Delta не ноль, продолжать обработку
                        if abs(list_Obg[n][1] - list_Obg[m][1]) <= Delta and abs(list_Obg[n][2] - list_Obg[m][2]) <= Delta: # если координата X и Y объекта n совпадает с координатами X и Y объекта m
                            del list_Obg[n] # удаляем первый объект (меньший диаметр)

                        else: # если не совпало пробуем следующий элемент
                            m = m + 1  # следующий элемент на обработку

                    else: # если Delta ноль, ставим центра/точки на каждый объект
                        break # прервать цикл

        if len(list_Obg) >= 1: # если в списке больше 1 объекта

            list_Obg.sort(key = lambda x: x[0:3]) # сортируем его по увеличению размеров (по первому, второму и третьему значению)

            removing_duplicates(list_Obg) # удаление дубликатов объектов по их положению X и Y (перебор списка)

        else: # если нет объекта в списке
            Kompas_message("Не найдены объекты для создания обозначений центра!") # сообщение в окне компаса если он открыт
            exit() # выходим из програмы

    list_processing() # обработка списка (удаление лишнего и сортировка отв. от большего к меньшему)

    score_Centre = 0 # количество обработаных объектов
    score_Point = 0 # количество обработаных объектов

    iGroup = iDocument2D.ksNewGroup(1) # создаём обозначения центров во временной группе

    if Centre_var: # ставить ли обозначения центров

        if Centre_repeat == False: # если повторные значения не нужны
            list_Obg_for_Centre = Сhecking_duplicates(list_Obg, list_Centre) # проверка на повторные обозначения центров/точек (из перевого списка вычитаем объекты второго)
        else:
            list_Obg_for_Centre = list_Obg

        for k in range(len(list_Obg_for_Centre)): # строим обозначения центров во все объекты со списка
            Centre(list_Obg_for_Centre[k][1], list_Obg_for_Centre[k][2], list_Obg_for_Centre[k][0], list_Obg_for_Centre[k][3], list_Obg_for_Centre[k][4]) # простановка центров окружностей (радиус/размер, положение X, положение Y, объект, вид)
            score_Centre += 1 # подсчёт количества обработаных объектов

    if Point_var: # ставить ли точки

        if Point_repeat == False: # если повторные значения не нужны
            list_Obg_for_Point = Сhecking_duplicates(list_Obg, list_Point) # проверка на повторные обозначения центров/точек (из перевого списка вычитаем объекты второго)
        else:
            list_Obg_for_Point = list_Obg

        for k in range(len(list_Obg_for_Point)): # строим обозначения центров во все объекты со списка
            Point(list_Obg_for_Point[k][1], list_Obg_for_Point[k][2], list_Obg_for_Point[k][3], list_Obg_for_Point[k][4]) # простановка точек в окружности (положение X, положение Y, объект, вид)
            score_Point += 1 # подсчёт количества обработаных объектов

    iDocument2D.ksEndGroup # завершить создание группы
    iDocument2D.ksStoreTmpGroup(iGroup) # вставляем группу в чертёж

    End_msg(score_Centre, score_Point) # вывод сообщения о количестве обработаных файлов

def Сhecking_duplicates(list1, list2): # проверка на повторные обозначения центров/точек (из перевого списка вычитаем объекты второго)

    List_Temp = [] # временый список для расчёта

    for Obg1 in list1: # перебор объектов
        for Obg2 in list2: # перебор объектов
            if abs(Obg1[1] - Obg2[1]) <= Delta and abs(Obg1[2] - Obg2[2]) <= Delta: # если координата X и Y объекта Obg1 совпадает с координатами X и Y объекта Obg2
                List_Temp.append(Obg1) # добавляем запись координат в список

    for Obg in List_Temp: # перебор объектов
        try: # попутаться урать объект со списка
            list1.remove(Obg) # удаляем совпадающие объекты
        except ValueError: # если уже нет объкта
            pass # пропустить

    return list1 # возвращаем список

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

    if Associate_Centre: # параметризовать ли обозначения центров
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

def Point(x, y, obj, iView): # простановка точек в окружности (положение X, положение Y, объект, вид)

    iDrawingContainer = KompasAPI7.IDrawingContainer(iView) # интерфейс контейнера объектов вида графического документа
    iPoints = iDrawingContainer.Points # коллекция точек
    iPoint = iPoints.Add() # создать новый элемент и добавить его в коллекцию
    iPoint.X = x # координата точки x
    iPoint.Y = y # координата точки y
    iPoint.Angle = 0.0 # угол наклона для точки со стрелкой
    iPoint.Style = 1 # аннотационные символы (см. ksAnnotationSymbolEnum и ksAnnotativeTerminatorSignEnum)
    iPoint.Update() # приминить свойства

    if Associate_Point: # параметризовать ли точки
        iDrawingObject1 = KompasAPI7.IDrawingObject1(iPoint) # дополнительный интерфейс для графических объектов
        iParametriticConstraint = iDrawingObject1.NewConstraint() # создатём новое ограничение
        iParametriticConstraint.ConstraintType = 22 # тип ограничения (см. ksConstraintTypeEnum)
        iParametriticConstraint.Index = 0 # индекс точки на объекте (начинается с 0, у дуги и окружности 0-центр)
        iParametriticConstraint.Partner = obj # второй объект или массив SAFEARRAY объектов для установки ограничения
        iParametriticConstraint.PartnerIndex = 0 # индекс точки на втором объекте (начинается с 0, у дуги и окружности 0-центр)
        iParametriticConstraint.Create() # cоздать ограничение в модели

def End_msg(score_Centre, score_Point): # вывод сообщения о количестве обработаных файлов

    if score_Point != 0: # если проставлены обозначения центров

        if str(score_Centre)[-1] in ["0","5","6","7","8","9"]: # если значение заканчиваеться на "0","5","6","7","8","9"
            text = "обозначений центров" # окончание в зависимости от значения
        elif str(score_Centre)[-1] == "1": # если значение заканчиваеться на "1"
            text = "обозначение центра" # окончание в зависимости от значения
        else: # если другое
            text = "обозначения центров" # окончание в зависимости от значения

        text_Centre = "Проставлено: " + str(score_Centre) + " " + text # текст сообщения

    else: # если нет обозначений центров
        text_Centre = "Обозначения центров уже проставлены!" # текст сообщения

    if score_Point != 0: # если проставлены точеки

        if str(score_Point)[-1] in ["0","5","6","7","8","9"]: # если значение заканчиваеться на "0","5","6","7","8","9"
            text = "точек" # окончание в зависимости от значения
        elif str(score_Point)[-1] == "1": # если значение заканчиваеться на "1"
            text = "точка" # окончание в зависимости от значения
        else: # если другое
            text = "точки" # окончание в зависимости от значения

        text_Point = "Проставлено: " + str(score_Point) + " " + text # текст сообщения

    else: # если нет проставленых точек
        text_Point = "Точки уже проставлены!" # текст сообщения

    Kompas_message(text_Centre + "\n" + text_Point) # сообщение в окне компаса если он открыт

#-------------------------------------------------------------------------------

list_Obg = [] # список для сбора координат и радиусов окружностей
list_Centre = [] # список для сбора координат обозначения центров
list_Point = [] # список для сбора координат точек

DoubleExe() # проверка на уже запущеное приложение, с отключённым консольным окном "CREATE_NO_WINDOW"

Сhecking_settings() # проверка правильности введённых значений

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

else: # если ничего не выбрано, выводим сообщение

    if All_iView: # обрабатывать ли все видимые виды
        for n in range(iViews.Count): # получаем количество всех видов
            iView = iViews.View(n) # вид по номеру
            View_processing(iView) # обработка всего вида

        Creating() # создание центров и точек

    else: # не обрабатывать виды
        Kompas_message("Нет выделенных объектов или вида!") # сообщение в окне компаса если он открыт

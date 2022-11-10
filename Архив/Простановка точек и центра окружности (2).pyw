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

    import pythoncom
    from win32com.client import Dispatch, gencache
    from sys import exit # для выхода из приложения без ошибки

    try:
        global KompasAPI7
        global iApplication
        global iKompasObject
        global iKompasDocument
        global iDocument2D
        global iKompasDocument2D
        global iKompasDocument2D1

        KompasConst3D = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants
        KompasConst = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа для скрытия вопросов перестроения

        KompasAPI5 = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0, 1, 0)
        iKompasObject = Dispatch('Kompas.Application.5', None, KompasAPI5.KompasObject.CLSID)

        KompasAPI7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0)
        iApplication = Dispatch('Kompas.Application.7')

        document_type() # определение типа документа

        iKompasDocument = iApplication.ActiveDocument # делаем активный открытый документ
        iDocument2D = iKompasObject.ActiveDocument2D() # указатель на интерфейс текущего графического документа
        iKompasDocument2D = KompasAPI7.IKompasDocument2D(iKompasDocument) # базовый класс графических документов КОМПАС
        iKompasDocument2D1 = KompasAPI7.IKompasDocument2D1(iKompasDocument) # дополнительный интерфейс IKompasDocument2D

        if iApplication.Visible == False: # если компас невидимый
            iApplication.Visible = True # сделать КОМПАС-3D видемым

    except SystemExit: # обрабатываем ошибку, если был использован "exit()"
        exit() # остановить программу

    except:
        if iApplication.Visible == True: # если компас видимый
            iApplication.MessageBoxEx("КОМПАС-3D не найден! Установите или переустановите КОМПАС-3D!", 'Message:', 64) # выдать сообщение
        exit() # завершаем макрос

def Kompas_message(text): # сообщение в окне компаса если он открыт

    if iApplication.Visible == True: # если компас видимый
        iApplication.MessageBoxEx(text, 'Message:', 64)

def document_type(): # определение типа документа

    from sys import exit # для выхода из приложения без ошибки

    iKompasDocument = iApplication.ActiveDocument # делаем активный открытый документ

    if iKompasDocument == None or iKompasDocument.DocumentType not in (1, 2): # если не открыт док. или не 2D док., выдать сообщение (1-чертёж; 2- фрагмент; 3-СП; 4-модель; 5-СБ; 6-текст. док.; 7-тех. СБ;)

        msg = "Откройте чертёж!"
        Kompas_message(msg) # сообщение в компасе
        exit() # завершаем макрос

def R_X_Y_obg(iSelectedObject): # получение координат и указатель на объект

    iReference = iSelectedObject.Reference # указатель объекта
    Type_obj = iDocument2D.ksGetObjParam(iReference, None, -1) # получить параметры объекта (объект, структура для записи параметров, тип параметра объекта (все параметры))

    if Type_obj in (2, 3, 36): # Обработка объектов по типу объекта (2 - окружность; 3 - дуга; 36 - многоугольник, см. CIRCLE_OBJ)
        R = iSelectedObject.Radius # радиус окружности
        X = iSelectedObject.Xc # координата центра окружности по оси Х
        Y = iSelectedObject.Yc # координата центра окружности по оси Х

        BD.append([R, X, Y, iSelectedObject]) # добавляем запись координат в список

    elif Type_obj in (32, 34): # Обработка объектов по типу объекта (32 - элипс; 34 - дуга элипса; 35 - прямоугольник, см. CIRCLE_OBJ)

        A = iSelectedObject.SemiAxisA # длина полуосей
        B = iSelectedObject.SemiAxisB # длина полуосей
        X = iSelectedObject.Xc # координата центра окружности по оси Х
        Y = iSelectedObject.Yc # координата центра окружности по оси Х

        BD.append([max(A, B), X, Y, iSelectedObject]) # добавляем запись координат в список

    elif Type_obj == 35: # Обработка объектов по типу объекта (35 - прямоугольник, см. CIRCLE_OBJ)

        A = iSelectedObject.Height # длина полуосей
        B = iSelectedObject.Width # длина полуосей
        X = iSelectedObject.X # координата центра окружности по оси Х
        Y = iSelectedObject.Y # координата центра окружности по оси Х

        BD.append([max(A, B), X, Y, iSelectedObject]) # добавляем запись координат в список

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

def Centre(x, y, R, obj): # простановка центров окружностей

    R = R*Scale # радиус окружности умноженый на масштаб
    iSymbols2DContainer = KompasAPI7.ISymbols2DContainer(iView) # интерфейс контейнера условных обозначений
    iCentreMarkers = iSymbols2DContainer.CentreMarkers # коллекция обозначений центра
    iCentreMarker = iCentreMarkers.Add() # создать новый элемент и добавить его в коллекцию

    iCentreMarker.BaseObject = obj # ассоциировать данный объект с другим объектом

    iCentreMarker.X = x # координата точки x делённая на масштаб
    iCentreMarker.Y = y # координата точки y
##    iCentreMarker.Angle = 0.0 # угол наклона обозначения центра
    iCentreMarker.SignType = 2 # тип обозначения центра: -1 - неизвестный 0 - маленький крестик; 1 - одна ось; 2 - две оси
    iCentreMarker.SetSemiAxisLength(0, R) # длина 1-ой полуоси
    iCentreMarker.SetSemiAxisLength(1, R) # длина 2-ой полуоси
    iCentreMarker.SetSemiAxisLength(2, R) # длина 3-ей полуоси
    iCentreMarker.SetSemiAxisLength(3, R) # длина 4-ой полуоси
    iAxisLineParam = KompasAPI7.IAxisLineParam(iCentreMarker) # параметры осевой линии или обозначения центра
    iAxisLineParam.JutLength = 5.0 # выступание осевой линии
    iAxisLineParam.DottedLength = 1.5 # длина пунктира
    iAxisLineParam.Interval = 1.5 # промежуток
    iAxisLineParam.AutoDetectedDash = True # ручное/автоматическое задание длины штриха
    iAxisLineParam.DashMaxLength = 15.0 # максимальная длина штриха
    iAxisLineParam.JutLengthModify = False # использовать параметр Выступание осевой из настроек документа
    iAxisLineParam.DottedLengthModify = False # использовать параметр Длина пунктира из настроек документа
    iAxisLineParam.IntervalModify = False # использовать параметр Промежуток из настроек документа
    iAxisLineParam.AutoDetectedDashModify = False # использовать параметр Задание длины штриха из настроек документа
    iAxisLineParam.DashMaxLengthModify = False # использовать параметр Максимальная длина штриха из настроек документа
    iCentreMarker.Update() # приминить свойства

##    iDrawingObject1 = KompasAPI7.IDrawingObject1(iCentreMarker) # дополнительный интерфейс для графических объектов
##    iDrawingObject1.Associate() # ассоциировать данный объект с другими объектами (не работает)

def Point(x, y, obj): # простановка точек в окружности

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

#-------------------------------------------------------------------------------

Number_iView = {} # cловарь для сбора параметров по видам
BD = [] # список для сбора координат и радиусов окружностей

DoubleExe() # проверка на уже запущеное приложение, с отключённым консольным окном "CREATE_NO_WINDOW"

KompasAPI() # подключение API компаса

iViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager # менеджер видов и слоев документа
iViews = iViewsAndLayersManager.Views # коллекция видов
iView = iViews.ActiveView # активный вид
Scale = iView.Scale # масштаб вида

iSelectionManager = iKompasDocument2D1.SelectionManager # менеджер выделенных объектов
iSelectedObjects = iSelectionManager.SelectedObjects # массив выделенных объектов в виде SAFEARRAY | VT_DISPATCH

if iSelectedObjects != None: # если выделены объекты

    if isinstance(iSelectedObjects, tuple): # если выбрано несколько объектов (кортеж объектов)
        for iSelectedObject in iSelectedObjects: # перебор всех выделеных объектов
            R_X_Y_obg(iSelectedObject) # получение координат и указатель на объект

    else:
        R_X_Y_obg(iSelectedObjects) # получение координат и указатель на объект

    list_processing() # обработка списка (удаление лишнего и сортировка отв. от большего к меньшему)

    iGroup = iDocument2D.ksNewGroup(1) # создаём обозначения центров во временной группе

    for k in range(len(BD)): # строим обозначения центров во все объекты со списка
        Centre(BD[k][1], BD[k][2], BD[k][0], BD[k][3]) # простановка центров окружностей
        Point(BD[k][1], BD[k][2], BD[k][3]) # простановка точек в окружности

    iDocument2D.ksEndGroup # завершить создание группы
    iDocument2D.ksStoreTmpGroup(iGroup) # вставляем группу в чертёж

##    iDrawingObjects = KompasAPI7.IDrawingObjects(iView) # базовый интерфейс для коллекций графических объектов
##    iDrawingObjects.Update() # обновляем вид

    iApplication.MessageBoxEx("Обозначения центров окружностей проставлены", 'Message:', 64) # выдать сообщение

else: # если ничего не выбрано, выводим сообщение
    Kompas_message("Нет выделенных объектов или вида!") # сообщение в окне компаса если он открыт

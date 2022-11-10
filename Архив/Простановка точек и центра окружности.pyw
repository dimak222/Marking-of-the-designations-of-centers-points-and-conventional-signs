#-------------------------------------------------------------------------------
# Name:        Простановка точек и центра окружности
# Purpose:
#
# Author:      dimak222
#
# Created:     18.08.2022
# Copyright:   (c) dimak222 2022
# Licence:     No
#-------------------------------------------------------------------------------

def KompasAPI(): # подключение API компаса

    import pythoncom
    from win32com.client import Dispatch, gencache
    from sys import exit

    try:
        global KompasAPI5
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

        iKompasDocument = iApplication.ActiveDocument # делаем активный открытый документ
        iDocument2D = iKompasObject.ActiveDocument2D() # указатель на интерфейс текущего графического документа
        iKompasDocument2D = KompasAPI7.IKompasDocument2D(iKompasDocument) # базовый класс графических документов КОМПАС
        iKompasDocument2D1 = KompasAPI7.IKompasDocument2D1(iKompasDocument) # дополнительный интерфейс IKompasDocument2D

        if iApplication.Visible == False: # если компас невидимый
            iApplication.Visible = True # сделать КОМПАС-3D видемым

    except:
        if iApplication.Visible == True: # если компас видимый
            iApplication.MessageBoxEx("КОМПАС-3D не найден! Установите или переустановите КОМПАС-3D!", 'Message:', 64) # выдать сообщение
        exit() # завершаем макрос

def R_X_Y_obg(iSelectedObject): # получение координат и указатель на объект

    iReference = iSelectedObject.Reference # указатель объекта
    type_obj = iDocument2D.ksGetObjParam(iReference, None, 0) # получить параметры объекта

    if type_obj == 2 or type_obj == 3: # Обработка объектов по типу объекта (2 - окружность; 3 - дуга, см. CIRCLE_OBJ)
        R = iSelectedObject.Radius # радиус окружности
        X = iSelectedObject.Xc # координата центра окружности по оси Х
        Y = iSelectedObject.Yc # координата центра окружности по оси Х
        obj = iSelectedObject # указатель на объект
        BD.append([R, X, Y, obj]) # добавляем запись координат в список

def list_processing(): # обработка списка (удаление лишнего и сортировка отв. от большего к меньшему)

    if len(BD) > 1: # если в списке больше 1 элемента

        BD.sort() #сортируем его
        #Разворачиваем его, чтобы окружности расположились в порядке убывания радиусов
        BD.reverse()
        #Устанавливаем начальный предел для перебора
        G=len(BD)
        #Устанавливаем начальный номер окружности в списке, для которой будем искать концентричные ей окружности
        J=0
        #Начинаем искать концентричные окружности
        while J<G:
            #Устанавливаем начальный предел для перебора
            g=len(BD)
            #Устанавливаем начальный номер окружности, которая может оказаться концентричной
            j=J+1
            #Перебираем список
            while j<g:
                #Если координаты j-й окружности совпадают,
                if BD[J][1]==BD[j][1] and BD[J][2]==BD[j][2]:
                    # то удаляем её из списка
                    del BD [j]
                    #Т.к. список уменьшился на 1 элемент, то граница перебора тоже должна уменьшиться
                    g=g-1
                #Если координаты не совпали, то переходим к следующей окружности из списка
                else:
                    j=j+1
            #Сравнили первый элемент списка с остальными, концентричные с малым радиусом удалили,
            #теперь сократим грануцу перебора первого цикла, узнав количество оставшихся окружностей в списке
            G=len(BD)
            #Возьмём следующую окружность для поиска окружностей концентричных уже ей
            J=J+1

def Centre_OLD(x, y, R): # простановка центров окружностей (старое)

    R = R*Scale # простановка центров окружностей
    iCentreParam = KompasAPI5.ksCentreParam(iKompasObject.GetParamStruct(93)) # интерфейс параметров объекта "обозначение центра"
    iCentreParam.Init() # инициализировать параметры
    iCentreParam.angle = 0.0 # угол наклона обозначения центра
    iCentreParam.lenXmTail = R # длина 1-ой полуоси
    iCentreParam.lenXpTail = R # длина 2-ой полуоси
    iCentreParam.lenYmTail = R # длина 3-ей полуоси
    iCentreParam.lenYpTail = R # длина 4-ой полуоси
    iCentreParam.standXmTail = True # признак стандартного изображения полуоси в отрицательном направлении оси Х
    iCentreParam.standXpTail = True # признак стандартного изображения полуоси в положительном направлении оси Х
    iCentreParam.standYmTail = True # признак стандартного изображения полуоси в отрицательном направлении оси Y
    iCentreParam.standYpTail = True # признак стандартного изображения полуоси в отрицательном направлении оси Y
    iCentreParam.type = 2 # тип обозначения центра: 0 - маленький крестик; 1 - одна ось; 2 - две оси
    iCentreParam.x = x # координаты точки привязки
    iCentreParam.y = y # координаты точки привязки
    iDocument2D.ksCentreMarker(iCentreParam) # cоздать обозначение центра

def Centre(x, y, R): # простановка центров окружностей

    R = R*Scale # радиус окружности умноженый на масштаб
    iSymbols2DContainer = KompasAPI7.ISymbols2DContainer(iView) # интерфейс контейнера условных обозначений
    iCentreMarkers = iSymbols2DContainer.CentreMarkers # коллекция обозначений центра
    iCentreMarker = iCentreMarkers.Add() # создать новый элемент и добавить его в коллекцию
    iCentreMarker.X = x/Scale # координата точки x делённая на масштаб
    iCentreMarker.Y = y # координата точки y
    iCentreMarker.Angle = 0.0 # угол наклона обозначения центра
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

    iDrawingObject1 = KompasAPI7.IDrawingObject1(iCentreMarker) # дополнительный интерфейс для графических объектов
    iDrawingObject1.Associate() # ассоциировать данный объект с другими объектами (не работает)

def Point(x, y, obj): # простановка точек в окружности

    iDrawingContainer = KompasAPI7.IDrawingContainer(iView) # интерфейс контейнера объектов вида графического документа
    points = iDrawingContainer.Points # коллекция точек
    point = points.Add() # создать новый элемент и добавить его в коллекцию
    point.X = x # координата точки x
    point.Y = y # координата точки y
    point.Angle = 0.0 # угол наклона для точки со стрелкой
    point.Style = 1 # аннотационные символы (см. ksAnnotationSymbolEnum и ksAnnotativeTerminatorSignEnum)
    point.Update() # приминить свойства

    iDrawingObject1 = KompasAPI7.IDrawingObject1(point) # дополнительный интерфейс для графических объектов
    iParametriticConstraint = iDrawingObject1.NewConstraint() # создатём новое ограничение
    iParametriticConstraint.ConstraintType = 11 # тип ограничения (см. ksConstraintTypeEnum)
    iParametriticConstraint.Index = 0 # индекс точки на объекте (начинается с 0, у дуги и окружности 0-центр)

##    iSelectionManager = iKompasDocument2D1.SelectionManager # менеджер выделенных объектов
##    try: # попытаться обработать массив элементов
##        iSelectedObjects = iSelectionManager.SelectedObjects[k] # массив выделенных объектов в виде SAFEARRAY | VT_DISPATCH
##    except: # обработать один элемент
##        iSelectedObjects = iSelectionManager.SelectedObjects # массив выделенных объектов в виде SAFEARRAY | VT_DISPATCH

    iParametriticConstraint.Partner = obj # второй объект или массив SAFEARRAY объектов для установки ограничения

##    R = iSelectedObjects.Radius # радиус окружности
##    X = iSelectedObjects.Xc # координата центра окружности по оси Х
##    Y = iSelectedObjects.Yc # координата центра окружности по оси Х
##    obj = iSelectedObjects
##    print(R, X, Y, obj)

    iParametriticConstraint.PartnerIndex = 0 # индекс точки на втором объекте (начинается с 0, у дуги и окружности 0-центр)
    iParametriticConstraint.Create() # cоздать ограничение в модели

#-------------------------------------------------------------------------------

BD = [] # список для сбора координат и радиусов окружностей

KompasAPI() # подключение API компаса

iViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager # менеджер видов и слоев документа
iViews = iViewsAndLayersManager.Views # коллекция видов
iView = iViews.ActiveView # активный вид
Scale = iView.Scale # масштаб вида

iSelectionManager = iKompasDocument2D1.SelectionManager # менеджер выделенных объектов
iSelectedObjects = iSelectionManager.SelectedObjects # массив выделенных объектов в виде SAFEARRAY | VT_DISPATCH

if iSelectedObjects != None: # если выделены объекты

    try: # попытаться обработать массив объектов
        for iSelectedObject in iSelectedObjects: # перебор всех выделеных объектов
            R_X_Y_obg(iSelectedObject) # получение координат и указатель на объект

    except: # обработать один элемент
        R_X_Y_obg(iSelectedObjects) # получение координат и указатель на объект

    list_processing() # обработка списка (удаление лишнего и сортировка отв. от большего к меньшему)

    for k in range(len(BD)): # строим обозначения центров для окружностей оставшихся в списке
        Centre (BD[k][1], BD[k][2], BD[k][0]) # простановка центров окружностей
        Point(BD[k][1], BD[k][2], BD[k][3]) # простановка точек в окружности

    iApplication.MessageBoxEx("Обозначения центров окружностей проставлены", 'Message:', 64) # выдать сообщение

else: # если ничего не выбрано, выводим сообщение
    iApplication.MessageBoxEx("Нет выделенных объектов!", 'Message:', 64) # выдать сообщение
##    iKompasObject.ksMessage("Нет выделенных объектов!") # выдать сообщение





##exit()
##
###--------------------------
##iCircleParam = iKompasObject.GetParamStruct(20) # указатель на интерфейс структуры параметров объекта нужного типа (20 - ksCircleParam)
##
##ksIterator = iKompasObject.GetIterator() # указатель на интерфейс ksIterator для навигации по объектам
##ksIterator.ksCreateIterator(127, 0) # создать итератор для перемещения по объектам документа (группа селектирования, перемещения по документам; видам; группам; слоям)
##obj = ksIterator.ksMoveIterator("F") # переместить итератор (позиционироваться на объекте) (F - первый объект, N - следующий объект)
###--------------------------------
##
##if obj != 0: # если объекты выделены
##
##    while obj != 0: # пока есть выделеные объекти продолжать
##
##        type_obj = iDocument2D.ksGetObjParam(obj, None, 0) # получить параметры объекта
##
##        if type_obj == 2 or type_obj == 3: # Тип объекта (2 - окружность; 3 - дуга, см. CIRCLE_OBJ)
##
##            iDocument2D.ksGetObjParam(obj, iCircleParam, -1) #  (см. ALLPARAM)
##            R=iCircleParam.rad
##            x=iCircleParam.xc
##            y=iCircleParam.yc
##            BD.append([R, x, y])
##
##        obj = ksIterator.ksMoveIterator("N")
##
##
##    for k in range(len(BD)): # строим обозначения центров для окружностей оставшихся в списке
##        Centre (BD[k][1], BD[k][2], BD[k][0]) # простановка центров окружностей
##        Point(BD[k][1], BD[k][2], k) # простановка точек в окружности
##
##    iApplication.MessageBoxEx("Обозначения центров окружностей проставлены", 'Message:', 64) # выдать сообщение
##
##else: # если ничего не выбрано, выводим сообщение
##    iKompasObject.ksMessage("Нет выделенных объектов!") # выдать сообщение
##
##ksIterator.ksDeleteIterator()

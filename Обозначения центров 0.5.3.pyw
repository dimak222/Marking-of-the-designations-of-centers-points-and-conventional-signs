# -*- coding: cp1251 -*-
# Простановка обозначений центров, окружностей, эллипсов, дуг эллипсов, прямоугольников и многоугольников в выделении и выделенных видах
# Обсуждение Вконтакте  https://vk.com/topic-152266438_36152984
# Обсуждение на форуме Аскон http://forum.ascon.ru/index.php/topic,29314.msg243699.html#msg243699

title = "ОЦ v0.5.4"

###########   Настройки   ###############

delta = 0.000001 #(мм). Максимально допустимая разница координат центров концентрическтих окружностей.
#Если  разница координат больше этой величины, окружности не считаются концентрическими.
RegularPolygon = True # Обрабатывать ли многоугольники (True - да, False - нет)
Rectangle = True # Обрабатывать ли прямоугольники (True - да, False - нет)
circle = True # Обрабатывать ли окружности (True - да, False - нет)
arc = True # Обрабатывать ли дуги окружностей (True - да, False - нет)
ellipse = True # Обрабатывать ли эллипсы (True - да, False - нет)
arc_ellipse = True # Обрабатывать ли дуги эллипсов (True - да, False - нет)
associate = True # Параметризовать ли обозначения центров (True - да, False - нет)
concentricity = True # Учитывать ли концентричность  (True - да, False - нет)
Style = range(1, 26) # Стиль линий объектов, которые нужно обрабатывать. Перечислить номера через запятую, напимер: style = (2,) или style = (1, 3, 5, 6)
# Для всех стилей: style = range(1, 26)
##Cистемные стили линии:
##1 - основная,
##2 - тонкая,
##3 - осевая,
##4 - штриховая,
##5 - для линии обрыва
##6 - вспомогательная,
##7 - утолщенная,
##8 - пунктир 2,
##9 - штриховая осн.
##10 - осевая осн.
##11 - тонкая линия, включаемая в штриховку
##12 - ISO 02 штриховая линия,
##13 - ISO 03 штриховая линия (дл. пробел),
##14 - ISO 04 штрихпунктирная линия (дл. штрих),
##15 - ISO 05 штрихпунктирная линия (дл. штрих 2 пунктира),
##16 - ISO 06 штрихпунктирная линия (дл. штрих 3 пунктира),
##17 - ISO 07 пунктирная линия,
##18 - ISO 08 штрихпунктирная линия (дл. и кор. штрихи),
##19 - ISO 09 штрихпунктирная линия (дл. и 2 кор. штриха),
##20 - ISO 10 штрихпунктирная линия,
##21 - ISO 11 штрихпунктирная линия (2 штриха),
##22 - ISO 12 штрихпунктирная линия (2 пунктира),
##23 - ISO 13 штрихпунктирная линия (3 пунктира),
##24 - ISO 14 штрихпунктирная линия (2 штриха 2 пунктира),
##25 - ISO 15 штрихпунктирная линия (2 штриха 3 пунктира).

#########################################

import pythoncom
from win32com.client import Dispatch, gencache

def Ob_centr (x, y, R, Circle):
    """
    Проставляет центры выделенных окружностей
    """
    global OC
    iCentreParam.Init()

    if associate:
        iCentreParam.baseCurve = Circle.Reference

    iCentreParam.angle = 360-Angle
    iCentreParam.lenXmTail = R
    iCentreParam.lenXpTail = R
    iCentreParam.lenYmTail = R
    iCentreParam.lenYpTail = R
    iCentreParam.standXmTail = True
    iCentreParam.standXpTail = True
    iCentreParam.standYmTail = True
    iCentreParam.standYpTail = True
    iCentreParam.type = 2
    iCentreParam.x = x
    iCentreParam.y = y
    iDocument2D.ksCentreMarker(iCentreParam)
    OC += 1


def For_View(iView):
    """
    Обрабатывает вид
    """
    if not iView.Visible: # если вид невидимый, пропускаем его
        return False

    iView_DO = kompas_api7_module.IDrawingObject(iView)

    if not iView.Current: # если не текущий
        iView.Current = True # делаем вид текущим
        iView_DO.Update()

    global Angle
    Scale = iView.Scale # масштаб
    Angle = iView.Angle # угол поворота
    key = iView.Number # номер вида

    iDrawingContainer = kompas_api7_module.IDrawingContainer(iView) # контейнер графических объектов
    List = [] # список для сбора объектов

    type_obj = [] # типы обрабатываемых объектов
    if circle:
        type_obj.append(2)
    if RegularPolygon:
        type_obj.append(3)
    if RegularPolygon:
        type_obj.append(32)
    if RegularPolygon:
        type_obj.append(34)
    if Rectangle:
        type_obj.append(35)
    if RegularPolygon:
        type_obj.append(36)

    if len(type_obj)==0:
        application.MessageBoxEx( "Не заданы типы объектов для обработки!", title, 48)
        exit()

    for n in type_obj:
        Objects = iDrawingContainer.Objects( n ) # получить массив объектов коллекции

        if Objects:
            for OBJ in Objects:
                List.append(OBJ) # добавить объект в список

    if List != []:
        BD={iView.Number:[]} # Словарь для сбора параметров

        for Circle in List:
            iLayer = iView.Layers.Layer (iDocument2D.ksGetLayerNumber(Circle.Reference))

            if iLayer.Visible:
                Sbor_info (BD, Circle )

        if len(BD[key]) > 1 and concentricity: #Если в списке больше 1 элемента
            BD[key] = Sort_BD(BD[iView.Number])

        #Строим обозначения центров для окружностей оставшихся в списке
        Group = iDocument2D.ksNewGroup(1) # создаём обозначения центров во временной группе

        for k in range(len(BD[key])):
            Ob_centr (BD[key][k][1], BD[key][k][2], BD[key][k][0]*Scale, BD[key][k][3])

        iDocument2D.ksEndGroup
        iDocument2D.ksStoreTmpGroup(Group) # вставляем группу в чертёж
        return True


def Sort_BD(BD):
    """
    Сортирует и фильтрует список
    """
    #сортируем
    BD.sort()
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
            if abs(BD[J][1] - BD[j][1]) <=  delta and abs(BD[J][2] - BD[j][2]) <=  delta :
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
    return BD


def Sbor_info (BD, obj):
    """
    Собирает параметры объектов в список
    """
    global s
    View = iViews.ViewByNumber( iDocument2D.ksGetViewNumber(obj.Reference) ).Number
    if obj.Type == 13075 and circle: # если окружность
        iDocument2D.ksGetObjParam( obj.Reference, iCircleParam, -1) #ALLPARAM

        if iCircleParam.style in Style:
            R = iCircleParam.rad
            x = iCircleParam.xc
            y = iCircleParam.yc
            BD[View].append([ R, x, y, obj ])
            s += 1

    elif obj.Type == 13027 and arc: # если дуга
        iDocument2D.ksGetObjParam( obj.Reference, iArcByAngleParam, -1) #ALLPARAM

        if iArcByAngleParam.style in Style:
            R = iArcByAngleParam.rad
            x = iArcByAngleParam.xc
            y = iArcByAngleParam.yc
            BD[View].append([ R, x, y, obj ])
            s += 1

    elif obj.Type == 13093 and ellipse: # если эллипс
        iDocument2D.ksGetObjParam( obj.Reference, iEllipseParam, -1) #ALLPARAM

        if iEllipseParam.style in Style:
            a = iEllipseParam.A
            b = iEllipseParam.B
            x = iEllipseParam.xc
            y = iEllipseParam.yc
            BD[View].append([ max(a, b), x, y, obj ])
            s += 1

    elif obj.Type == 13095 and arc_ellipse: # если дуга эллипса
        iDocument2D.ksGetObjParam( obj.Reference, iEllipseArcParam, -1) #ALLPARAM

        if iEllipseArcParam.style in Style:
            a = iEllipseArcParam.A
            b = iEllipseArcParam.B
            x = iEllipseArcParam.xc
            y = iEllipseArcParam.yc
            BD[View].append([ max(a, b), x, y, obj ])
            s += 1

    elif obj.Type == 13097 and Rectangle: # если прямоугольник
        iDocument2D.ksGetObjParam( obj.Reference, iRectangleParam, -1) #ALLPARAM

        if iRectangleParam.style in Style:
            a = iRectangleParam.height #Высота прямоугольника.
            b = iRectangleParam.width  #Ширина прямоугольника
            x = iRectangleParam.x
            y = iRectangleParam.y
            BD[View].append([ max(a, b), x, y, obj ])
            s += 1

    elif obj.Type == 13099 and RegularPolygon: # если многоугольник
        iDocument2D.ksGetObjParam( obj.Reference, iRegularPolygonParam, -1) #ALLPARAM

        if iRegularPolygonParam.style in Style:
            R = iRegularPolygonParam.radius  #Радиус вписанной или описанной окружности
            x = iRegularPolygonParam.xc
            y = iRegularPolygonParam.yc
            BD[View].append([ R, x, y, obj ])
            s += 1

    return s


def Text(n): # вывод сообщения о количестве обработаных файлов

    if str(n)[-1] in ["0","5","6","7","8","9"]:
        text = "\nобозначений центра"
    elif str(n)[-1] == "1":
        text = "\nобозначение центра"
    else:
        text = "\nобозначения центра"

    return "Создано " + str(n) + text

#-------------------------------------------------------------------------------

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
iKompasObject = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))

iDocument2D = iKompasObject.ActiveDocument2D()
iDocument = application.ActiveDocument

iKompasDocument2D = kompas_api7_module.IKompasDocument2D(iDocument)
iKompasDocument2D1 = kompas_api7_module.IKompasDocument2D1(iKompasDocument2D)

iViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager
iViews = iViewsAndLayersManager.Views
ActiveView = iViews.ActiveView # запомним активный вид


iSelectionManager = iKompasDocument2D1.SelectionManager # интерфейс выделения
SelectedObjects = iSelectionManager.SelectedObjects

iCentreParam = kompas6_api5_module.ksCentreParam(iKompasObject.GetParamStruct(93)) # ko_CentreParam
iCircleParam = iKompasObject.GetParamStruct(20) # ko_CircleParam
iArcByAngleParam = iKompasObject.GetParamStruct(12) # ko_ArcByAngleParam
iEllipseParam = iKompasObject.GetParamStruct(22) # ko_EllipseParam
iEllipseArcParam = iKompasObject.GetParamStruct(23) # ko_EllipsArcParam
# интерфейс прямоугольник
iRectangleParam =  iKompasObject.GetParamStruct(91) # ko_RectangleParam
# интерфейс многоугольника
iRegularPolygonParam =  iKompasObject.GetParamStruct(92) # ko_RegularPolygonParam

OC, s = 0, 0

if SelectedObjects: # если есть выделение
    BD = {} #Словарь для сбора параметров по видам
    if isinstance(SelectedObjects, tuple): # если несколько объектов

        for i in range(iViews.Count):
            number =  iViews.View (i).Number
            BD[number] = []

        for obj in SelectedObjects:

            if obj.Type in [10031, 10032]: # если вид
                For_View(obj)

            else:
                s = Sbor_info (BD, obj)

        if s:

            for key in BD.keys():

                if len(BD[key]) > 1 and concentricity: #Если в списке больше 1 элемента
                    BD[key] = Sort_BD(BD[key])

                if not iViews.ViewByNumber ( key ).Current:
                    iViews.ViewByNumber ( key ).Current = True # возвращаем активность виду
                    kompas_api7_module.IDrawingObject(iViews.ViewByNumber ( key )).Update()

                Scale = iViews.ViewByNumber ( key ).Scale # масштаб
                Angle = iViews.ViewByNumber ( key ).Angle # угол поворота

                #Строим обозначения центров для окружностей оставшихся в списке
                Group = iDocument2D.ksNewGroup(1) # создаём обозначения центров во временной группе

                for k in range(len(BD[key])):
                    Ob_centr (BD[key][k][1], BD[key][k][2], BD[key][k][0]*Scale, BD[key][k][3])

                iDocument2D.ksEndGroup
                iDocument2D.ksStoreTmpGroup(Group) # вставляем группу в чертёж

    else:

        if SelectedObjects.Type in [10031, 10032]: # если вид
            For_View(SelectedObjects)

        else:
            View = iDocument2D.ksGetViewNumber(SelectedObjects.Reference)
            BD[View] = []
            s = Sbor_info (BD, SelectedObjects)

            if s:
                for key in BD.keys():
                   if not iViews.ViewByNumber ( key ).Current:
                        iViews.ViewByNumber ( key ).Current = True # возвращаем активность виду
                        kompas_api7_module.IDrawingObject(iViews.ViewByNumber ( key )).Update()

                   Scale = iViews.ViewByNumber ( key ).Scale # масштаб
                   Angle = iViews.ViewByNumber ( key ).Angle # угол поворота
                   Ob_centr (BD[key][0][1], BD[key][0][2], BD[key][0][0]*Scale, BD[key][0][3])

else:
    for i in range(iViews.Count):
        iView = iViews.View ( i )
        For_View(iView)

if not ActiveView.Current:
    ActiveView.Current = True # возвращаем активность виду
    kompas_api7_module.IDrawingObject(ActiveView).Update()

if OC:
    application.MessageBoxEx( Text(OC), title, 64)
else:
    application.MessageBoxEx( "Не найдены объекты для создания обозначений центра!", title, 48)




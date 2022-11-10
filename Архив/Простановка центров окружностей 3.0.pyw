# -*- coding: cp1251 -*-
# Простановка обозначений центров окружностей в выделении активного вида

import pythoncom
from win32com.client import Dispatch, gencache

#Список для сбора координат и радиусов окружностей
BD=[]

#Проставляет центры выделенных окружностей
def Ob_centr (x, y, R):
    R = R*Scale
    iCentreParam = kompas6_api5_module.ksCentreParam(iKompasObject.GetParamStruct(93))#LDefin2D.ko_CentreParam))
    iCentreParam.Init()
    iCentreParam.angle = 0.0
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


#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
iKompasObject = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))
#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))

iDocument2D = iKompasObject.ActiveDocument2D()
iDocument = application.ActiveDocument
iKompasDocument2D = kompas_api7_module.IKompasDocument2D(iDocument)

iViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager
iViews = iViewsAndLayersManager.Views
iView = iViews.ActiveView
Scale = iView.Scale


iCircleParam = iKompasObject.GetParamStruct(20) #ko_CircleParam

ksIterator = iKompasObject.GetIterator()
ksIterator.ksCreateIterator( 127, 0 )
obj = ksIterator.ksMoveIterator( "F" )

if obj != 0:

    while obj != 0:
        type_obj = iDocument2D.ksGetObjParam( obj, None, 0)

        if type_obj == 2: #CIRCLE_OBJ
            iDocument2D.ksGetObjParam( obj, iCircleParam, -1) #ALLPARAM
            R=iCircleParam.rad
            x=iCircleParam.xc
            y=iCircleParam.yc
            BD.append([ R, x, y ])

        obj = ksIterator.ksMoveIterator( "N" )

    #Если в списке больше 1 элемента
    if len(BD) > 1:
        #сортируем его
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

    #Строим обозначения центров для окружностей оставшихся в списке
    for k in range(len(BD)):
            Ob_centr (BD[k][1],BD[k][2],BD[k][0])

    iKompasObject.ksMessage("Обозначения центров окружностей проставлены")
#Если ничего не выбрано, выводим сообщение
else:
    iKompasObject.ksMessage("Нет выделенных объектов!")
ksIterator.ksDeleteIterator()

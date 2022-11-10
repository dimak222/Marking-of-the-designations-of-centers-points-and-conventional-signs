# -*- coding: cp1251 -*-
# ����������� ����������� �������, �����������, ��������, ��� ��������, ��������������� � ��������������� � ��������� � ���������� �����
# ���������� ���������  https://vk.com/topic-152266438_36152984
# ���������� �� ������ ����� http://forum.ascon.ru/index.php/topic,29314.msg243699.html#msg243699

title = "�� v0.5.4"

###########   ���������   ###############

delta = 0.000001 #(��). ����������� ���������� ������� ��������� ������� ���������������� �����������.
#����  ������� ��������� ������ ���� ��������, ���������� �� ��������� ����������������.
RegularPolygon = True # ������������ �� �������������� (True - ��, False - ���)
Rectangle = True # ������������ �� �������������� (True - ��, False - ���)
circle = True # ������������ �� ���������� (True - ��, False - ���)
arc = True # ������������ �� ���� ����������� (True - ��, False - ���)
ellipse = True # ������������ �� ������� (True - ��, False - ���)
arc_ellipse = True # ������������ �� ���� �������� (True - ��, False - ���)
associate = True # ��������������� �� ����������� ������� (True - ��, False - ���)
concentricity = True # ��������� �� ���������������  (True - ��, False - ���)
Style = range(1, 26) # ����� ����� ��������, ������� ����� ������������. ����������� ������ ����� �������, �������: style = (2,) ��� style = (1, 3, 5, 6)
# ��� ���� ������: style = range(1, 26)
##C�������� ����� �����:
##1 - ��������,
##2 - ������,
##3 - ������,
##4 - ���������,
##5 - ��� ����� ������
##6 - ���������������,
##7 - ����������,
##8 - ������� 2,
##9 - ��������� ���.
##10 - ������ ���.
##11 - ������ �����, ���������� � ���������
##12 - ISO 02 ��������� �����,
##13 - ISO 03 ��������� ����� (��. ������),
##14 - ISO 04 ��������������� ����� (��. �����),
##15 - ISO 05 ��������������� ����� (��. ����� 2 ��������),
##16 - ISO 06 ��������������� ����� (��. ����� 3 ��������),
##17 - ISO 07 ���������� �����,
##18 - ISO 08 ��������������� ����� (��. � ���. ������),
##19 - ISO 09 ��������������� ����� (��. � 2 ���. ������),
##20 - ISO 10 ��������������� �����,
##21 - ISO 11 ��������������� ����� (2 ������),
##22 - ISO 12 ��������������� ����� (2 ��������),
##23 - ISO 13 ��������������� ����� (3 ��������),
##24 - ISO 14 ��������������� ����� (2 ������ 2 ��������),
##25 - ISO 15 ��������������� ����� (2 ������ 3 ��������).

#########################################

import pythoncom
from win32com.client import Dispatch, gencache

def Ob_centr (x, y, R, Circle):
    """
    ����������� ������ ���������� �����������
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
    ������������ ���
    """
    if not iView.Visible: # ���� ��� ���������, ���������� ���
        return False

    iView_DO = kompas_api7_module.IDrawingObject(iView)

    if not iView.Current: # ���� �� �������
        iView.Current = True # ������ ��� �������
        iView_DO.Update()

    global Angle
    Scale = iView.Scale # �������
    Angle = iView.Angle # ���� ��������
    key = iView.Number # ����� ����

    iDrawingContainer = kompas_api7_module.IDrawingContainer(iView) # ��������� ����������� ��������
    List = [] # ������ ��� ����� ��������

    type_obj = [] # ���� �������������� ��������
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
        application.MessageBoxEx( "�� ������ ���� �������� ��� ���������!", title, 48)
        exit()

    for n in type_obj:
        Objects = iDrawingContainer.Objects( n ) # �������� ������ �������� ���������

        if Objects:
            for OBJ in Objects:
                List.append(OBJ) # �������� ������ � ������

    if List != []:
        BD={iView.Number:[]} # ������� ��� ����� ����������

        for Circle in List:
            iLayer = iView.Layers.Layer (iDocument2D.ksGetLayerNumber(Circle.Reference))

            if iLayer.Visible:
                Sbor_info (BD, Circle )

        if len(BD[key]) > 1 and concentricity: #���� � ������ ������ 1 ��������
            BD[key] = Sort_BD(BD[iView.Number])

        #������ ����������� ������� ��� ����������� ���������� � ������
        Group = iDocument2D.ksNewGroup(1) # ������ ����������� ������� �� ��������� ������

        for k in range(len(BD[key])):
            Ob_centr (BD[key][k][1], BD[key][k][2], BD[key][k][0]*Scale, BD[key][k][3])

        iDocument2D.ksEndGroup
        iDocument2D.ksStoreTmpGroup(Group) # ��������� ������ � �����
        return True


def Sort_BD(BD):
    """
    ��������� � ��������� ������
    """
    #���������
    BD.sort()
    #������������� ���, ����� ���������� ������������� � ������� �������� ��������
    BD.reverse()
    #������������� ��������� ������ ��� ��������
    G=len(BD)
    #������������� ��������� ����� ���������� � ������, ��� ������� ����� ������ ������������� �� ����������
    J=0
    #�������� ������ ������������� ����������
    while J<G:
        #������������� ��������� ������ ��� ��������
        g=len(BD)
        #������������� ��������� ����� ����������, ������� ����� ��������� �������������
        j=J+1
        #���������� ������
        while j<g:
            #���� ���������� j-� ���������� ���������,
            if abs(BD[J][1] - BD[j][1]) <=  delta and abs(BD[J][2] - BD[j][2]) <=  delta :
                # �� ������� � �� ������
                del BD [j]
                #�.�. ������ ���������� �� 1 �������, �� ������� �������� ���� ������ �����������
                g=g-1
            #���� ���������� �� �������, �� ��������� � ��������� ���������� �� ������
            else:
                j=j+1
        #�������� ������ ������� ������ � ����������, ������������� � ����� �������� �������,
        #������ �������� ������� �������� ������� �����, ����� ���������� ���������� ����������� � ������
        G=len(BD)
        #������ ��������� ���������� ��� ������ ����������� ������������� ��� ��
        J=J+1
    return BD


def Sbor_info (BD, obj):
    """
    �������� ��������� �������� � ������
    """
    global s
    View = iViews.ViewByNumber( iDocument2D.ksGetViewNumber(obj.Reference) ).Number
    if obj.Type == 13075 and circle: # ���� ����������
        iDocument2D.ksGetObjParam( obj.Reference, iCircleParam, -1) #ALLPARAM

        if iCircleParam.style in Style:
            R = iCircleParam.rad
            x = iCircleParam.xc
            y = iCircleParam.yc
            BD[View].append([ R, x, y, obj ])
            s += 1

    elif obj.Type == 13027 and arc: # ���� ����
        iDocument2D.ksGetObjParam( obj.Reference, iArcByAngleParam, -1) #ALLPARAM

        if iArcByAngleParam.style in Style:
            R = iArcByAngleParam.rad
            x = iArcByAngleParam.xc
            y = iArcByAngleParam.yc
            BD[View].append([ R, x, y, obj ])
            s += 1

    elif obj.Type == 13093 and ellipse: # ���� ������
        iDocument2D.ksGetObjParam( obj.Reference, iEllipseParam, -1) #ALLPARAM

        if iEllipseParam.style in Style:
            a = iEllipseParam.A
            b = iEllipseParam.B
            x = iEllipseParam.xc
            y = iEllipseParam.yc
            BD[View].append([ max(a, b), x, y, obj ])
            s += 1

    elif obj.Type == 13095 and arc_ellipse: # ���� ���� �������
        iDocument2D.ksGetObjParam( obj.Reference, iEllipseArcParam, -1) #ALLPARAM

        if iEllipseArcParam.style in Style:
            a = iEllipseArcParam.A
            b = iEllipseArcParam.B
            x = iEllipseArcParam.xc
            y = iEllipseArcParam.yc
            BD[View].append([ max(a, b), x, y, obj ])
            s += 1

    elif obj.Type == 13097 and Rectangle: # ���� �������������
        iDocument2D.ksGetObjParam( obj.Reference, iRectangleParam, -1) #ALLPARAM

        if iRectangleParam.style in Style:
            a = iRectangleParam.height #������ ��������������.
            b = iRectangleParam.width  #������ ��������������
            x = iRectangleParam.x
            y = iRectangleParam.y
            BD[View].append([ max(a, b), x, y, obj ])
            s += 1

    elif obj.Type == 13099 and RegularPolygon: # ���� �������������
        iDocument2D.ksGetObjParam( obj.Reference, iRegularPolygonParam, -1) #ALLPARAM

        if iRegularPolygonParam.style in Style:
            R = iRegularPolygonParam.radius  #������ ��������� ��� ��������� ����������
            x = iRegularPolygonParam.xc
            y = iRegularPolygonParam.yc
            BD[View].append([ R, x, y, obj ])
            s += 1

    return s


def Text(n): # ����� ��������� � ���������� ����������� ������

    if str(n)[-1] in ["0","5","6","7","8","9"]:
        text = "\n����������� ������"
    elif str(n)[-1] == "1":
        text = "\n����������� ������"
    else:
        text = "\n����������� ������"

    return "������� " + str(n) + text

#-------------------------------------------------------------------------------

#  ��������� �������� ����������� API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
iKompasObject = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))

#  ��������� �������� ����������� API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))

iDocument2D = iKompasObject.ActiveDocument2D()
iDocument = application.ActiveDocument

iKompasDocument2D = kompas_api7_module.IKompasDocument2D(iDocument)
iKompasDocument2D1 = kompas_api7_module.IKompasDocument2D1(iKompasDocument2D)

iViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager
iViews = iViewsAndLayersManager.Views
ActiveView = iViews.ActiveView # �������� �������� ���


iSelectionManager = iKompasDocument2D1.SelectionManager # ��������� ���������
SelectedObjects = iSelectionManager.SelectedObjects

iCentreParam = kompas6_api5_module.ksCentreParam(iKompasObject.GetParamStruct(93)) # ko_CentreParam
iCircleParam = iKompasObject.GetParamStruct(20) # ko_CircleParam
iArcByAngleParam = iKompasObject.GetParamStruct(12) # ko_ArcByAngleParam
iEllipseParam = iKompasObject.GetParamStruct(22) # ko_EllipseParam
iEllipseArcParam = iKompasObject.GetParamStruct(23) # ko_EllipsArcParam
# ��������� �������������
iRectangleParam =  iKompasObject.GetParamStruct(91) # ko_RectangleParam
# ��������� ��������������
iRegularPolygonParam =  iKompasObject.GetParamStruct(92) # ko_RegularPolygonParam

OC, s = 0, 0

if SelectedObjects: # ���� ���� ���������
    BD = {} #������� ��� ����� ���������� �� �����
    if isinstance(SelectedObjects, tuple): # ���� ��������� ��������

        for i in range(iViews.Count):
            number =  iViews.View (i).Number
            BD[number] = []

        for obj in SelectedObjects:

            if obj.Type in [10031, 10032]: # ���� ���
                For_View(obj)

            else:
                s = Sbor_info (BD, obj)

        if s:

            for key in BD.keys():

                if len(BD[key]) > 1 and concentricity: #���� � ������ ������ 1 ��������
                    BD[key] = Sort_BD(BD[key])

                if not iViews.ViewByNumber ( key ).Current:
                    iViews.ViewByNumber ( key ).Current = True # ���������� ���������� ����
                    kompas_api7_module.IDrawingObject(iViews.ViewByNumber ( key )).Update()

                Scale = iViews.ViewByNumber ( key ).Scale # �������
                Angle = iViews.ViewByNumber ( key ).Angle # ���� ��������

                #������ ����������� ������� ��� ����������� ���������� � ������
                Group = iDocument2D.ksNewGroup(1) # ������ ����������� ������� �� ��������� ������

                for k in range(len(BD[key])):
                    Ob_centr (BD[key][k][1], BD[key][k][2], BD[key][k][0]*Scale, BD[key][k][3])

                iDocument2D.ksEndGroup
                iDocument2D.ksStoreTmpGroup(Group) # ��������� ������ � �����

    else:

        if SelectedObjects.Type in [10031, 10032]: # ���� ���
            For_View(SelectedObjects)

        else:
            View = iDocument2D.ksGetViewNumber(SelectedObjects.Reference)
            BD[View] = []
            s = Sbor_info (BD, SelectedObjects)

            if s:
                for key in BD.keys():
                   if not iViews.ViewByNumber ( key ).Current:
                        iViews.ViewByNumber ( key ).Current = True # ���������� ���������� ����
                        kompas_api7_module.IDrawingObject(iViews.ViewByNumber ( key )).Update()

                   Scale = iViews.ViewByNumber ( key ).Scale # �������
                   Angle = iViews.ViewByNumber ( key ).Angle # ���� ��������
                   Ob_centr (BD[key][0][1], BD[key][0][2], BD[key][0][0]*Scale, BD[key][0][3])

else:
    for i in range(iViews.Count):
        iView = iViews.View ( i )
        For_View(iView)

if not ActiveView.Current:
    ActiveView.Current = True # ���������� ���������� ����
    kompas_api7_module.IDrawingObject(ActiveView).Update()

if OC:
    application.MessageBoxEx( Text(OC), title, 64)
else:
    application.MessageBoxEx( "�� ������� ������� ��� �������� ����������� ������!", title, 48)




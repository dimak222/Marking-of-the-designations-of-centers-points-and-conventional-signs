# -*- coding: cp1251 -*-
# ����������� ����������� ������� ����������� � ��������� ��������� ����

import pythoncom
from win32com.client import Dispatch, gencache

#������ ��� ����� ��������� � �������� �����������
BD=[]

#����������� ������ ���������� �����������
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


#  ��������� �������� ����������� API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
iKompasObject = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))
#  ��������� �������� ����������� API7
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

    #���� � ������ ������ 1 ��������
    if len(BD) > 1:
        #��������� ���
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
                if BD[J][1]==BD[j][1] and BD[J][2]==BD[j][2]:
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

    #������ ����������� ������� ��� ����������� ���������� � ������
    for k in range(len(BD)):
            Ob_centr (BD[k][1],BD[k][2],BD[k][0])

    iKompasObject.ksMessage("����������� ������� ����������� �����������")
#���� ������ �� �������, ������� ���������
else:
    iKompasObject.ksMessage("��� ���������� ��������!")
ksIterator.ksDeleteIterator()

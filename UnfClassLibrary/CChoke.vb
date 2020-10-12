'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
'
' ����� ��� ������� ������������� �������
' ������������ ��� ����������� ����������� ����� '������' �������������� �������
' ����������� ������� - �� ��� ��� ���������� ������������ ������ ����� ������ (�������� ������ �� ��������� �����)
' �������� �� �������� � ������������ ��������� ��������� ��������� ������� �� ����� - �� ���� ��������� ������
' � �������� ����� �������� ����� ��������������� ��������� �������� �������� ����� (��������)


'==============  Cchoke  ==============
' ����� ��� ������� ������������ ������ � ��������� ������������� - �������
Option Explicit On

Public Class CChoke
    ' �������������� ��������� �������
    Public d_up_m As Double
    Public d_down_m As Double
    Public d_choke_m As Double

    Public t_choke_C As Double

    '����� ����������� ����� ������
    Public fluid As New CPVT

    Public c_calibr_fr As Double
    Private c_degrad_choke_ As Double                             ' choke correction factor

    ' ������ ��� ������� ������������� �������
    ' �������� ��� ������� ���������� �������
    Public curve As New CCurves

    Private q_liqmax_m3day_ As Double  ' ������������ ����� ��� �������� �������� �� ����� � �� ������ ����� ������
    Private t_choke_throat_C_ As Double ' ����������� � �������
    Private t_choke_av_C_
    Public sonic_vel_msec As Double

    ' ����� ���������� ��� ������� ��� �������� ��������� ������
    'Private p_pbuf_atma As Double
    'Private p_plin_atma As Double

    ' internal vars
    ' ��������� ������ �������
    Private K As Double '  = 0.826,'K - Discharge coefficient (optional, default  is 0.826)
    Private f As Double ' = 1.25,'F - Ratio of gas spec. heat capacity at constant pressure to that at constant volume (optional, default  is 1.4)
    Private c_vw As Double ' = 4176'Cvw - water specific heat capacity (J/kg K)(optional, default  is 4176)

    Private a_u As Double 'upstream area
    Private a_c As Double 'choke throat area
    Private a_r As Double 'area ratio

    Private P_r As Double  ' critical pressure for output
    Private v_s As Double  ' sonic velosity
    Private q_m As Double  ' mass rate

    Private p_dcr As Double ' recovered downstream pressure at critical pressure ratio
End Class
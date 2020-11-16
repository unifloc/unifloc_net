﻿'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'                                                                      good (11/21/2019)
'=======================================================================================
' class describes reservoir properties and IPR
' allows to work with IPR data based on production test data
' use Vogel's correction for IPR with watercut and composite IPR
'
' reference
' 1. Brown, Kermit (1984). The Technology of Artificial Lift Methods. Volume 4.
'    Production Optimization of Oil and Gas Wells by Nodal System Analysis.
'    Tulsa, Oklahoma: PennWellBookss.
' 2. Vogel, J.V. 1968. Inflow Performance Relationships for Solution-Gas Drive Wells.
'    J Pet Technol 20 (1): 83–92. SPE 1476-PA. http://dx.doi.org/10.2118/1476-PA

Option Explicit On
Imports System.Math

Public Class CReservoirVogel
    'Implements IReservoir

    Public pi_sm3dayatm As Double
    Public p_res_atma As Double
    Public fluid As CPVT    ' take bubble point pressure and watercut from fluid
    Private IPR_curve_ As CInterpolation


    Public ReadOnly Property pb_atma() As Double
        Get
            pb_atma = fluid.pb_atma
        End Get
    End Property

    Public ReadOnly Property fw_perc() As Double
        Get
            fw_perc = fluid.Fw_perc
        End Get
    End Property

    ' IPR curve
    ' must be generated before access with proper sub
    Public ReadOnly Property IPRCurve() As CInterpolation
        Get
            IPRCurve = IPR_curve_
        End Get
    End Property

    ' ==================================================================
    ' main calculation functions and subroutines
    ' ==================================================================

    ' initialisation sub - set IPR properties from minimal data set
    Public Sub InitProp(Optional ByVal p_res_atma As Double = 0,
                        Optional ByVal pb_atma As Double = 0,
                        Optional ByVal fw_perc As Double = 0)
        ' p_res_atma - reservoir pressure
        ' pb_atma   - bubble point pressure
        ' fw_perc   - fraction of water in flow (watercut)

        Me.p_res_atma = p_res_atma
        fluid = New CPVT
        fluid.Class_Initialize()
        fluid.pb_atma = pb_atma
        fluid.Fw_perc = fw_perc
    End Sub

    ' calculate liquid rate from BHP with IPR given
    Public Function calc_qliq_sm3day(ByVal pwf_atma As Double) As Double
        ' Pwf_atma - bottom hole pressure

        calc_qliq_sm3day = calc_Q_IPR_m3Day(pwf_atma, p_res_atma, pb_atma, pi_sm3dayatm, fw_perc)
    End Function

    ' calculate BHP  from liquid rate with IPR given
    Public Function calc_pwf_atma(ByVal qtest As Double) As Double
        ' qtest - liquid rate
        calc_pwf_atma = calc_pwf_IPres_atma(qtest, p_res_atma, pb_atma, pi_sm3dayatm, fw_perc)
    End Function

    ' calculate productivity index by test rate and BHP
    Public Function calc_pi_sm3dayatm(ByVal qtest As Double, ByVal Ptest As Double) As Double
        ' qtest  - test liquid rate
        ' Ptest  - test bottom hole pressure
        pi_sm3dayatm = calc_pi_IPR_m3DayAtm(qtest, Ptest, p_res_atma, pb_atma, fw_perc)
        calc_pi_sm3dayatm = pi_sm3dayatm
    End Function

    ' generate IPR curve as CInterpolation object
    Public Function Build_IPRcurve() As CInterpolation
        Dim i As Integer
        Dim Qstep As Double
        Dim p_wf As Double
        'Dim Qliq_reserv As Double
        Dim maxQ As Double
        IPR_curve_ = New CInterpolation
        Const IPRNumPoints = 30

        maxQ = calc_qliq_sm3day(0)
        Qstep = maxQ / IPRNumPoints
        For i = 0 To IPRNumPoints
            p_wf = calc_pwf_atma(i * Qstep)
            IPR_curve_.AddPoint(i * Qstep, p_wf)
        Next i
        Build_IPRcurve = IPR_curve_
    End Function

    ' ==============================================================================
    ' interface properties and functions implementation
    ' ==============================================================================

    Private Function IReservoir_CalcPI(ByVal qtest As Double, ByVal Ptest As Double) As Double
        IReservoir_CalcPI = calc_pi_sm3dayatm(qtest, Ptest)
    End Function

    Private Function IReservoir_CalcPwf(ByVal qtest As Double) As Double
        IReservoir_CalcPwf = calc_pwf_atma(qtest)
    End Function


    Private Function IReservoir_CalcQliq(ByVal Ptest_atma As Double) As Double
        IReservoir_CalcQliq = calc_qliq_sm3day(Ptest_atma)
    End Function

    Private Property IReservoir_pi() As Double
        Get
            IReservoir_pi = pi_sm3dayatm
        End Get
        Set(RHS As Double)
            pi_sm3dayatm = RHS
        End Set
    End Property

    '    Private Property Get IReservoir_pi() As Double
    '    IReservoir_pi = pi_sm3dayatm
    'End Property

    Private Property IReservoir_Pres() As Double
        Get
            IReservoir_Pres = p_res_atma
        End Get
        Set(RHS As Double)
            p_res_atma = RHS
        End Set
    End Property

    '    Private Property Get IReservoir_Pres() As Double
    '    IReservoir_Pres = p_res_atma
    'End Property


    ' =======================================================================
    ' private vogel's correlation functions
    ' =======================================================================

    Private Function calc_Q_IPR_m3Day(ByVal Ptest As Double, ByVal Pr As Double,
                          ByVal pb As Double, ByVal Pi As Double, ByVal Wc As Double, Optional calc_method As Integer = 1) As Double
        If Ptest >= Pr Then
            ' addLogMsg "Предупреждение. функция calc_Q_IPR_m3Day. тестовое забойное давление " & Ptest & " больше или равно пластового " & Pr & ". дебит 0"
            calc_Q_IPR_m3Day = 0
            Exit Function
        End If

        Select Case calc_method
            Case 1
                calc_Q_IPR_m3Day = CDbl(calc_QliqVogel_m3Day(Ptest, Pr, pb, Pi, Wc))
        End Select
    End Function

    Private Function calc_pwf_IPres_atma(ByVal qtest As Double, ByVal Pr As Double,
                          ByVal pb As Double, ByVal Pi As Double, ByVal Wc As Double, Optional calc_method As Integer = 1) As Double

        Select Case calc_method
            Case 1
                calc_pwf_IPres_atma = CDbl(calc_p_wfVogel_atma(qtest, Pr, pb, Pi, Wc))
        End Select
    End Function

    Private Function calc_pi_IPR_m3DayAtm(ByVal qtest As Double, ByVal Ptest As Double,
                                ByVal Pr As Double, ByVal pb As Double, ByVal Wc As Double, Optional calc_method As Integer = 1) As Double

        If Ptest >= Pr Then
            ' addLogMsg "Ошибка. функция calc_pi_IPR_m3DayAtm. тестовое забойное давление " & Ptest & " больше чем пластовое давление " & Pr & ". расчет продуктивности невозможен"
            calc_pi_IPR_m3DayAtm = -1
            Exit Function
        End If

        Select Case calc_method
            Case 1
                calc_pi_IPR_m3DayAtm = calc_PIVogel_m3DayAtm(qtest, Ptest, Pr, pb, Wc)
        End Select
    End Function

    'Расчет забойного давления по Вогелю с учетом поправки на обводненность
    Private Function calc_QliqVogel_m3Day(ByVal P_test As Double, ByVal Pr As Double,
                          ByVal pb As Double, ByVal Pi As Double, ByVal Wc As Double) As Object
        '
        ' Q_test    - дебит жидкости при котором надо определить заб. давл. м3/сут
        ' Pr        - пластовое давление, атм
        ' Pb        - давление насыщения, атм
        ' pi - коэффициент продуктивности, м3/сут/атм
        ' wc        - обводненность, %

        Dim qb As Double
        Dim qo_max As Double
        Dim p_wfg As Double
        Dim CG As Double
        Dim cd As Double
        Dim fw As Double
        Dim fo As Double

        If P_test < 0 Then
            calc_QliqVogel_m3Day = "P_test<0!"
            Exit Function
        End If
        If Pr < 0 Then
            calc_QliqVogel_m3Day = "Pr<0!"
            Exit Function
        End If
        If pb < 0 Then
            calc_QliqVogel_m3Day = "Pb<0!"
            Exit Function
        End If
        If Pi < 0 Then
            calc_QliqVogel_m3Day = "PI<0!"
            Exit Function
        End If
        If Pr < pb Then
            pb = Pr
        End If

        ' вычисляем дебит при давлении равном давлению насыщения.
        qb = Pi * (Pr - pb)
        If Wc > 100 Then
            Wc = 100
        End If
        If Wc < 0 Then
            Wc = 0
        End If

        If (Wc = 100) Or (P_test >= pb) Then

            calc_QliqVogel_m3Day = Pi * (Pr - P_test)

        Else
            fw = Wc / 100
            fo = 1 - fw
            ' максимальный дебит чистой нефти
            qo_max = qb + (Pi * pb) / 1.8
            '  Dim pwf_g As Double
            p_wfg = fw * (Pr - qo_max / Pi)

            If P_test > p_wfg Then
                Dim a As Double : Dim b As Double : Dim c As Double : Dim d As Double
                a = 1 + (P_test - (fw * Pr)) / (0.125 * fo * pb)
                b = fw / (0.125 * fo * pb * Pi)
                c = (2 * a * b) + 80 / (qo_max - qb)
                d = (a ^ 2) - (80 * qb / (qo_max - qb)) - 81
                If b = 0 Then
                    calc_QliqVogel_m3Day = Abs(d / c)
                Else
                    calc_QliqVogel_m3Day = (-c + ((c * c - 4 * b * b * d) ^ 0.5)) / (2 * b ^ 2)
                End If

            Else
                CG = 0.001 * qo_max
                cd = fw * (CG / Pi) +
              fo * 0.125 * pb * (-1 + (1 + 80 * ((0.001 * qo_max) / (qo_max - qb))) ^ 0.5)
                calc_QliqVogel_m3Day = (p_wfg - P_test) / (cd / CG) + qo_max
            End If

        End If

    End Function

    ' Расчет забойного давления по Вогелю с учетом поправки на обводненность
    Private Function calc_p_wfVogel_atma(ByVal Q_test As Double, ByVal Pr As Double,
                          ByVal pb As Double, ByVal Pi As Double, ByVal Wc As Double) As Double
        '
        ' Q_test    - дебит жидкости при котором надо определить заб. давл. м3/сут
        ' Pr        - пластовое давление, атм
        ' Pb  - давление насыщения, атм
        ' pi - коэффициент продуктивности, м3/сут/атм
        ' wc        - обводненность, %

        Dim qb As Double
        Dim qo_max As Double
        'Dim p_wfg As Double
        Dim CG As Double
        Dim cd As Double
        Dim fw As Double
        Dim fo As Double

        'проверка  данных

        If Pr < pb Then
            pb = Pr
        End If

        If Q_test < 0 Then
            calc_p_wfVogel_atma = "Q<0!"
            Exit Function
        End If
        If Pr <= 0 Then
            calc_p_wfVogel_atma = "Pr<=0!"
            Exit Function
        End If
        If pb < 0 Then
            calc_p_wfVogel_atma = "Pb<0!"
            Exit Function
        End If
        If Pi <= 0 Then
            calc_p_wfVogel_atma = "PI<=0!"
            Exit Function
        End If

        ' вычисляем дебит при давлении равном давлению насыщения.
        qb = Pi * (Pr - pb)
        If Wc > 100 Then
            Wc = 100
        End If
        If Wc < 0 Then
            Wc = 0
        End If
        If (Wc = 100) Or (Q_test <= qb) Or (pb = 0) Then

            calc_p_wfVogel_atma = (Pr - Q_test / Pi)

        Else
            fw = Wc / 100
            fo = 1 - fw
            ' максимальный дебит чистой нефти
            qo_max = qb + (Pi * pb) / 1.8

            If Q_test < qo_max Then

                calc_p_wfVogel_atma = fw * (Pr - Q_test / Pi) +
                     fo * 0.125 * pb * (-1 + (1 - 80 * ((Q_test - qo_max) / (qo_max - qb))) ^ 0.5)

            Else
                CG = 0.001 * qo_max
                cd = fw * (CG / Pi) +
               fo * 0.125 * pb * (-1 + (1 + 80 * ((0.001 * qo_max) / (qo_max - qb))) ^ 0.5)
                calc_p_wfVogel_atma = fw * (Pr - qo_max / Pi) - (Q_test - qo_max) * (cd / CG)
            End If
        End If
        If calc_p_wfVogel_atma < 0 Then
            calc_p_wfVogel_atma = 0
        End If
    End Function

    ' определение продуктивности по Вогелю с коррекцией на обводненность
    Private Function calc_PIVogel_m3DayAtm(ByVal Q_test As Double, ByVal P_test As Double,
                                ByVal pres As Double, ByVal pb As Double, ByVal Wc As Double) As Double

        Dim j As Double
        Dim Q_calibr As Double

        If P_test < 0 Then
            P_test = 0
            calc_PIVogel_m3DayAtm = 0
            Exit Function
        End If

        If pres < pb Then
            pb = pres
        End If

        If Q_test <= 0 Then
            calc_PIVogel_m3DayAtm = 0 ' "Q<=0!"
            '  addLogMsg "calc_PIVogel_m3DayAtm  ошибка Q<=0!"
            Exit Function
        End If
        If P_test <= 0 Then
            calc_PIVogel_m3DayAtm = 0
            ' addLogMsg "calc_PIVogel_m3DayAtm  ошибка P_test<=0!"
            Exit Function
        End If
        If pb < 0 Then
            calc_PIVogel_m3DayAtm = 0
            '  addLogMsg "calc_PIVogel_m3DayAtm  ошибка Pb<0!!"
            Exit Function
        End If
        If pres <= 0 Then
            calc_PIVogel_m3DayAtm = 0
            '  addLogMsg "calc_PIVogel_m3DayAtm  ошибка Pres<=0!"
            Exit Function
        End If
        ' первое приближение для Кпр считаем
        j = Q_test / (pres - P_test)
        ' оцениваем его
        Q_calibr = CDbl(calc_QliqVogel_m3Day(P_test, pres, pb, j, Wc))
        ' кривая нелинейная надо подбирать значение
        j = j / ((Q_calibr) / Q_test)
        Q_calibr = CDbl(calc_QliqVogel_m3Day(P_test, pres, pb, j, Wc))
        If Abs(Q_calibr - Q_test) > 0.001 Then Debug.Assert(False)   ' если сработало, то алгоритм сломался
        calc_PIVogel_m3DayAtm = j

    End Function

End Class

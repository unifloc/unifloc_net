﻿'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
' Gaslift valve and gas line valve

Option Explicit On
Public Class CGLvalve

    Public h_mes_m As Double            ' глубина установки газлифтного клапана
    Public p_bellow_sc_atma As Double   ' давление зарядки сильфона на поверхности на поверхности

    Public p_in_atma As Double          ' casing pressure at gas valve
    Public p_v_atma As Double           ' pressure between inlet and outlet
    Public p_out_atma As Double         ' tubing pressure at gas valve

    Public t_C As Double             ' casing temperature at gas valve

    Public q_gas_inj_scm3day As Double  ' gas rate through valve
    Public crit_flow As Double
    Public p_crit_atma As Double
    Public q_gas_max_sm3day As Double

    ' IPO valve data
    ' for now supposed Wheatherford R-1 data
    Public r As Double
    Public dext_mm As Double
    Public Ab_mm2 As Double
    Public Ap_mm2 As Double
    Public PREF As Double
    Public IPO As Boolean

    Private d_vkr_mm_() As Double
    Private d_vkr_eff_mm_ As Double
    Private d_port_mm_ As Double             ' диаметр порта


    Public fluid As New CPVT                 ' gas properties

    Public Sub Class_init(Optional ByVal h_mes_m_ As Double = 0,
                          Optional ByVal p_bellow_sc_atma_ As Double = 0,
                          Optional ByVal p_in_atma_ As Double = 0,
                          Optional ByVal p_v_atma_ As Double = 0,
                          Optional ByVal p_out_atma_ As Double = 0,
                          Optional ByVal t_C_ As Double = 0,
                          Optional ByVal q_gas_inj_scm3day_ As Double = 0,
                          Optional ByVal crit_flow_ As Double = 0,
                          Optional ByVal p_crit_atma_ As Double = 0,
                          Optional ByVal q_gas_max_sm3day_ As Double = 0,
                          Optional ByVal r_ As Double = 0,
                          Optional ByVal dext_mm_ As Double = 0,
                          Optional ByVal Ab_mm2_ As Double = 0,
                          Optional ByVal Ap_mm2_ As Double = 0,
                          Optional ByVal PREF_ As Double = 0,
                          Optional ByVal IPO_ As Boolean = False,
                          Optional ByVal d_vkr_mm() As Double = Nothing,
                          Optional ByVal _vkr_eff_mm As Double = 0,
                          Optional ByVal d_port_mm As Double = 0)

        h_mes_m = h_mes_m_
        p_bellow_sc_atma = p_bellow_sc_atma_
        p_in_atma = p_in_atma_
        p_v_atma = p_v_atma_
        p_out_atma = p_out_atma_
        t_C = t_C_
        q_gas_inj_scm3day = q_gas_inj_scm3day_
        crit_flow = crit_flow_
        p_crit_atma = p_crit_atma_
        q_gas_max_sm3day = q_gas_max_sm3day_
        r = r_
        dext_mm = dext_mm_
        Ab_mm2 = Ab_mm2_
        Ap_mm2 = Ap_mm2_
        PREF = PREF_
        IPO = IPO_
        d_vkr_mm_ = d_vkr_mm
        d_vkr_eff_mm_ = _vkr_eff_mm
        d_port_mm_ = d_port_mm

    End Sub

    ' задаем характеристики клапана R-1 с учетом наличия вкруток
    Public Sub set_GLV_R1(Optional ByVal IPO As Boolean = False,
                      Optional port_mm As Double = GLV_R1_PORT_SIZE.R1_port_1_8,
                      Optional d_vkr1_mm As Double = -1,
                      Optional d_vkr2_mm As Double = -1,
                      Optional d_vkr3_mm As Double = -1,
                      Optional d_vkr4_mm As Double = -1)
        ' set R-1 valve data
        Dim Ap_Ab As Double
        Dim i As Integer

        Me.IPO = IPO

        d_port_mm_ = port_mm
        dext_mm = 25.4
        Ab_mm2 = 200
        Select Case port_mm
            Case GLV_R1_PORT_SIZE.R1_port_1_4
                Ap_mm2 = 8.4
            Case GLV_R1_PORT_SIZE.R1_port_5_32
                Ap_mm2 = 13.55
            Case GLV_R1_PORT_SIZE.R1_port_3_16
                Ap_mm2 = 18.71
            Case GLV_R1_PORT_SIZE.R1_port_1_8
                Ap_mm2 = 33.55
            Case GLV_R1_PORT_SIZE.R1_port_5_16
                Ap_mm2 = 51.61
            Case Else
                Ap_mm2 = (port_mm ^ 2) * (1 - 0.16)  ' approximation based on R-1 table
        End Select
        Ap_Ab = Ap_mm2 / Ab_mm2
        r = Ap_Ab
        PREF = Ap_Ab / (1 - Ap_Ab)
        ' estimate effective diameter with "vkrutka"
        Dim num_vkr As Integer
        num_vkr = 0
        If d_vkr1_mm > 0 Then
            num_vkr = num_vkr + 1
            ReDim Preserve d_vkr_mm_(0 To num_vkr)
            d_vkr_mm_(num_vkr) = d_vkr1_mm
        End If
        If d_vkr2_mm > 0 Then
            num_vkr = num_vkr + 1
            ReDim Preserve d_vkr_mm_(0 To num_vkr)
            d_vkr_mm_(num_vkr) = d_vkr2_mm
        End If
        If d_vkr3_mm > 0 Then
            num_vkr = num_vkr + 1
            ReDim Preserve d_vkr_mm_(0 To num_vkr)
            d_vkr_mm_(num_vkr) = d_vkr3_mm
        End If
        If d_vkr4_mm > 0 Then
            num_vkr = num_vkr + 1
            ReDim Preserve d_vkr_mm_(0 To num_vkr)
            d_vkr_mm_(num_vkr) = d_vkr4_mm
        End If
        d_vkr_eff_mm_ = 0
        If num_vkr > 0 Then
            For i = d_vkr_mm_.GetLowerBound(0) To d_vkr_mm_.GetUpperBound(0)
                d_vkr_eff_mm_ = d_vkr_eff_mm_ + (d_vkr_mm_(i)) ^ 2
            Next i
        End If
        d_vkr_eff_mm_ = d_vkr_eff_mm_ ^ 0.5

    End Sub

    Public Property d_port_mm() As Double
        Get
            Return d_port_mm_
        End Get
        Set(val As Double)
            d_port_mm_ = val
        End Set
    End Property

    Public ReadOnly Property p_open_atma() As Double
        Get
            If IPO Then
                ' for opening assume p_v_atma = p_out_atma
                p_open_atma = p_bellow_rc_atma / (1 - r) - p_out_atma * r / (1 - r)
            Else
                p_open_atma = 1
            End If
        End Get
    End Property

    Public ReadOnly Property p_bellow_rc_atma() As Double
        Get
            If IPO Then
                p_bellow_rc_atma = GLV_p_close_atma(p_bellow_sc_atma, t_C)
            Else
                p_bellow_rc_atma = 1
            End If
        End Get
    End Property
    ' функция расчета расхода газа через клапан
    Public Function calc_q_gas_sm3day(Optional p_intake_atma As Double = -1,
                                 Optional p_out_atma As Double = -1,
                                 Optional t_in_C As Double = -1) As Double
        Dim rslt() As Object

        If p_intake_atma > 0 Then
            p_in_atma = p_intake_atma
        Else
            p_intake_atma = p_in_atma
        End If
        If p_out_atma > 0 Then
            Me.p_out_atma = p_out_atma
        Else
            p_out_atma = Me.p_out_atma
        End If
        If p_intake_atma < 0 Then p_intake_atma = Me.p_in_atma
        If p_out_atma < 0 Then p_out_atma = Me.p_out_atma
        If t_in_C < 0 Then t_in_C = Me.t_C

        If (p_out_atma < p_intake_atma) And (d_port_mm_ > 0) And (d_vkr_eff_mm_ = 0) Then
            rslt = GLV_q_gas_sm3day(d_port_mm_, p_intake_atma, p_out_atma, fluid.gamma_g, t_in_C)
            calc_q_gas_sm3day = rslt(0) ' (0)(0)
            p_crit_atma = rslt(1) ' (0)(1)
            crit_flow = rslt(2) ' (0)(2)
            '  p_v_atma = p_out_atma
        ElseIf (p_out_atma < p_intake_atma) And (d_port_mm_ > 0) And (d_vkr_eff_mm_ > 0) Then
            rslt = GLV_q_gas_vkr_sm3day(d_port_mm_, d_vkr_eff_mm_, p_intake_atma, p_out_atma, fluid.gamma_g, t_in_C)
            calc_q_gas_sm3day = rslt(0) ' (0)(0)
            p_v_atma = rslt(2) ' (0)(2)
        End If

        If IPO Then
            ' need check open condition
            Dim pdif As Double
            If p_intake_atma < p_open_atma + 2 Then
                pdif = p_open_atma + 2 - p_intake_atma
                If pdif < 0 Then pdif = 0
                calc_q_gas_sm3day = calc_q_gas_sm3day * (pdif) / 2
            End If
        End If

    End Function

    Public Function calc_p_out_atma(p_intake_atma As Double, q_gas_scm3day As Double)
        Me.p_in_atma = p_intake_atma
        Dim res() As Object
        res = GLV_p_vkr_atma(d_port_mm, d_vkr_eff_mm_, p_intake_atma, q_gas_scm3day, fluid.gamma_g, t_C, True)

        '    ' ищем давление внутри клапана
        '    If d_vkr_eff_mm_ > 0 Then
        '        p_v_atma = GLV_p_atma(d_vkr_eff_mm_, p_intake_atma, q_gas_scm3day, fluid.gamma_g, t_C, True)
        '    Else
        '        p_v_atma = p_out_atma
        '    End If
        '    p_out_atma = GLV_p_atma(d_port_mm_, p_v_atma, q_gas_scm3day, fluid.gamma_g, t_C, True)
        '    calc_p_out_atma = p_out_atma
        '
        p_v_atma = res(2) ' (0)(2)
        p_in_atma = res(0) ' (0)(0)

        If IPO Then
            ' need check open condition
            Dim pdif As Double
            If p_in_atma < p_open_atma + 2 Then
                pdif = p_open_atma + 2 - p_in_atma
                If pdif < 0 Then pdif = 0
                calc_p_out_atma = calc_q_gas_sm3day() * (pdif) / 2
            End If
        End If


        calc_p_out_atma = p_in_atma

    End Function

    ' расчет давления на входе в клапан (затрубное давление)
    ' по давлению в НКТ
    Public Function calc_p_in_atma(p_out_atma As Double, q_gas_scm3day As Double)
        Me.p_out_atma = p_out_atma
        ' ищем давление внутри клапана
        Dim res() As Object
        res = GLV_p_vkr_atma(d_port_mm, d_vkr_eff_mm_, p_out_atma, q_gas_scm3day, fluid.gamma_g, t_C, False)
        '    If d_vkr_eff_mm_ > 0 Then
        '        p_v_atma = GLV_p_atma(d_vkr_eff_mm_, p_out_atma, q_gas_scm3day, fluid.gamma_g, t_C, False)
        '    Else
        '        p_v_atma = p_out_atma
        '    End If
        p_v_atma = res(2) '(0)(2)
        p_in_atma = res(0) '(0)(0)
        calc_p_in_atma = p_in_atma
    End Function


    Public Function table_pin(ByVal p_in_atma As Double, ByVal t_in_C As Double) As CInterpolation
        ' calculate valve characteristics table for given d
        Dim i As Integer
        Dim q As Double
        Dim Tbl As New CInterpolation
        Dim qeps As Double
        Dim p As Double
        qeps = 0.001
        With Tbl
            q_gas_max_sm3day = Me.calc_q_gas_sm3day(p_in_atma, 1, t_in_C)
            .AddPoint(1, q_gas_max_sm3day)
            .AddPoint(p_crit_atma, q_gas_max_sm3day - qeps)
            .AddPoint(p_in_atma, 0)
            Dim N As Integer
            N = 10
            For i = 1 To N - 1
                p = p_in_atma - (p_in_atma - p_crit_atma) / N * i
                q = Me.calc_q_gas_sm3day(p_in_atma, p, t_in_C)
                .AddPoint(p, q)
            Next i
        End With
        table_pin = Tbl
    End Function
End Class

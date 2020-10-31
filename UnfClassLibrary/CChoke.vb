'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
'
' класс для расчета характеристик штуцера
' потребовался для организации корректного учета 'крутой' характеристики штуцера
' особенность штуцера - то что при достижении критического потока через штуцер (движения потока со скоростью звука)
' давление за штуцером в определенном диапазоне перестает оказывать влияние на поток - то есть заданному дебиту
' и давление перед штуцером может соответствовать несколько значений давлений после (линейных)


'==============  Cchoke  ==============
' класс для расчета многофазного потока в локальном сопротивлении - штуцере
Option Explicit On
Imports System.Math
Public Class CChoke
    ' геометрические параметры штуцера
    Public d_up_m As Double
    Public d_down_m As Double
    Public d_choke_m As Double

    Public t_choke_C As Double

    'флюид протекающий через штуцер
    Public fluid As New CPVT

    Public c_calibr_fr As Double
    Private c_degrad_choke_ As Double                             ' choke correction factor

    ' кривые для текущих характеристик штуцера
    ' строятся для текущих параметров штуцера
    Public curve As New CCurves

    Private q_liqmax_m3day_ As Double  ' максимальный дебит для заданных давлений на входе и на выходе через штуцер
    Private t_choke_throat_C_ As Double ' температура в штуцере
    Private t_choke_av_C_
    Public sonic_vel_msec As Double

    ' набор параметров для которых был проведен последний расчет
    'Private p_pbuf_atma As Double
    'Private p_plin_atma As Double

    ' internal vars
    ' параметры модели штуцера
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

    Public Sub Class_Initialize(Optional ByVal K_ As Double = 0.826,
                                Optional ByVal f_ As Double = 1.25,
                                Optional ByVal c_vw_ As Double = 1,
                                Optional ByVal c_calibr_fr_ As Double = 1,
                                Optional ByVal c_degrad_choke As Double = 0,
                                Optional ByVal d_up_m_ As Double = 0.1,
                                Optional ByVal d_down_m_ As Double = 0.1,
                                Optional ByVal d_choke_m_ As Double = 0.01,
                                Optional ByVal t_choke_C_ As Double = 30,
                                Optional ByVal q_liqmax_m3day As Double = 0,
                                Optional ByVal t_choke_throat_C As Double = 0,
                                Optional ByVal t_choke_av_C As Double = 0,
                                Optional ByVal sonic_vel_msec_ As Double = 0,
                                Optional ByVal a_u_ As Double = 0,
                                Optional ByVal a_c_ As Double = 0,
                                Optional ByVal a_r_ As Double = 0,
                                Optional ByVal P_r_ As Double = 0,
                                Optional ByVal v_s_ As Double = 0,
                                Optional ByVal q_m_ As Double = 0,
                                Optional ByVal p_dcr_ As Double = 0)

        K = K_  'K - Discharge coefficient (optional, default  is 0.826)
        f = f_
        c_vw = c_vw_
        c_calibr_fr = c_calibr_fr_
        c_degrad_choke_ = c_degrad_choke

        'параметры по умолчанию
        d_up_m = d_up_m_
        d_down_m = d_down_m_
        d_choke_m = d_choke_m_
        t_choke_C = t_choke_C_

        q_liqmax_m3day_ = q_liqmax_m3day
        t_choke_throat_C_ = t_choke_throat_C
        t_choke_av_C_ = t_choke_av_C
        sonic_vel_msec = sonic_vel_msec_

        a_u = a_u_
        a_c = a_c_
        a_r = a_r_
        P_r = P_r_
        v_s = v_s_
        q_m = q_m_
        p_dcr = p_dcr_
    End Sub

    Public ReadOnly Property D_choke_mm() As Double
        Get
            D_choke_mm = d_choke_m * 1000
        End Get
    End Property

    Public ReadOnly Property Fw_fr() As Double
        Get
            Fw_fr = fluid.Fw_fr ' fw_perc / 100
        End Get
    End Property

    Public ReadOnly Property Qlmax_m3day()
        Get
            Qlmax_m3day = q_liqmax_m3day_
        End Get
    End Property

    Public ReadOnly Property TchokeThroat_C()
        Get
            TchokeThroat_C = t_choke_throat_C_
        End Get
    End Property

    Public ReadOnly Property TchokeAv_C()
        Get
            TchokeAv_C = t_choke_av_C_
        End Get
    End Property

    Public ReadOnly Property PratioCrit()
        Get
            PratioCrit = P_r
        End Get
    End Property

    Public ReadOnly Property VelSonic_msec()
        Get
            VelSonic_msec = v_s
        End Get
    End Property

    Public ReadOnly Property Qm_kgsec()
        Get
            Qm_kgsec = q_m
        End Get
    End Property

    Public ReadOnly Property PdownCrit_atma()
        Get
            PdownCrit_atma = p_dcr
        End Get
    End Property

    Public Function calc_choke_calibration(
                                        ByVal p_intake_atma As Double,
                                        ByVal p_out_atma As Double,
                                        t_C As Double) As Double
        Dim qtest As Double
        t_choke_C = t_C
        If (p_intake_atma > p_out_atma) And d_choke_m > 0 Then
            qtest = calc_choke_qliq_sm3day(p_intake_atma, p_out_atma, t_choke_C)
            c_calibr_fr = fluid.qliq_sm3day / qtest
        Else
            c_calibr_fr = 1
        End If
    End Function

    Public Function calc_choke_qliq_sm3day(
                                          ByVal p_u As Double,
                                          ByVal p_d As Double,
                                          ByVal t_u As Double) As Double
        'Function calculates oil flow rate through choke given downstream and upstream pressures using Perkins correlation
        'Return ((sm3/day))
        'Arguments
        'p_u - Upstream pressure ( (atma))
        'p_d - Downstream pressure ( (atma))
        'T_u - Upstream temperature ( (C))

        Dim p_co As Double = 0
        Dim min_p_d As Double = 0
        Dim counter As Double = 0
        Dim w_i As Double = 0
        Dim wi_der1 As Double = 0
        Dim d_pr As Double = 0
        Const max_iters As Integer = 10
        Dim eps As Double = 0
        Dim p_ri As Double = 0
        '   Dim v_si As Double
        Dim p_dcr As Double = 0
        Dim p_c As Double = 0
        Dim p_ra As Double = 0
        Dim w As Double = 0
        Const p_r_inc As Double = 0.001

        Try
            ' calc areas
            Call init_params()

            Call fluid.Calc_PVT(p_u, t_u) ' calc PVT with upstream pressure and temperature
            With fluid
                'Calculate trial output choke pressure
                p_co = p_u - (p_u - p_d) / (1 - (d_choke_m / d_down_m) ^ 1.85)
                'Solve for critical pressure ratio
                counter = 0
                If (.Fm_gas_fr > 0.0000000000001) Then 'free gas present
                    'Calculate specific value of error at p_ri = 0.99
                    w_i = wi_calc(0.99, p_u, t_u, eps)
                    eps = Abs(eps * 0.01)
                    'Assume pressure ratio
                    p_ri = 0.5
                    Do
                        'Evaluate derivative for two points to find second derivative for Newton-Raphson iteration
                        w_i = wi_calc(p_ri, p_u, t_u, wi_der1, p_r_inc, d_pr)
                        'limit p_ri increment to prevent crossing [0,1] boundary
                        d_pr = Max(-p_ri / 2, Min(d_pr, (1 - p_ri) / 2))
                        p_ri = p_ri + d_pr
                        counter = counter + 1
                    Loop Until (Abs(wi_der1) < eps) Or (counter > max_iters)
                    If counter > max_iters Then
                        AddLogMsg("Cchoke.calc_choke_qliq_sm3day: iterations not converged. iterations number  = " & counter & "  error wi_der1 " & wi_der1 & " < " & eps)
                    End If

                    'Calculate sonic velocity of multiphase mixture (used for output)
                    sonic_vel_msec = w_i / a_c * (.Fm_oil_fr / .Rho_oil_rc_kgm3 + .Fm_wat_fr / .Rho_wat_rc_kgm3 + .Fm_gas_fr / .Rho_gas_rc_kgm3 * p_ri ^ (-1 / .Polytropic_exponent)) / 86400
                Else 'liquid flow
                    p_ri = 0
                    sonic_vel_msec = 5000
                End If

                ' calc PVT with upstream pressure and temperature
                Call fluid.Calc_PVT(p_u, t_u)

                q_liqmax_m3day_ = K * w_i * .Fm_oil_fr / .Rho_oil_sckgm3 + K * w_i * .Fm_wat_fr / .Rho_wat_sckgm3
                q_liqmax_m3day_ = q_liqmax_m3day_ * c_calibr_fr

                'Calculate recovered downstream pressure at critical pressure ratio
                p_dcr = p_u * (p_ri * (1 - (d_choke_m / d_down_m) ^ 1.85) + (d_choke_m / d_down_m) ^ 1.85)
                'Compare trial pressure ratio with critical and assign actual pressure ratio
                'Auxilary properties
                p_c = p_ri * p_u
                p_ra = Max(p_ri, p_co / p_u)
                w_i = wi_calc(p_ra, p_u, t_u, wi_der1)
                'Calculate isentropic mass flow rate
                w = K * w_i * c_calibr_fr
                ' calc PVT with upstream pressure and temperature
                Call fluid.Calc_PVT(p_u, t_u)

                calc_choke_qliq_sm3day = w * .Fm_oil_fr / .Rho_oil_sckgm3 + w * .Fm_wat_fr / .Rho_wat_sckgm3
                'Asign mass flow rate
                q_m = w / 86400 '/ c_m(Units)
                'Assign output critical pressure ratio (recovered critical pressure ratio)
                P_r = p_dcr / p_u
                'convert sonic velocity
                v_s = sonic_vel_msec '/ c_l(Units)
                Exit Function
            End With

        Catch ex As Exception
            Dim errmsg As String
            errmsg = "CChoke.calc_choke_qliq_sm3day:" & ex.Message
            AddLogMsg(errmsg)
            Throw New ApplicationException(errmsg)
        End Try
    End Function

    Private Sub init_params()
        a_u = const_Pi * d_up_m ^ 2 / 4      'upstream area
        a_c = const_Pi * d_choke_m ^ 2 / 4   'choke throat area
        a_r = a_c / a_u                     'area ratio
    End Sub

    Private Function wi_calc(P_r As Double,
                         p_u As Double,
                         t_u As Double,
          Optional ByRef wi_deriv As Double = -1,
          Optional p_r_inc As Double = -1,
          Optional ByRef d_pr As Double = -1)
        'Auxilary properties
        Dim t_C As Double
        Dim p_av As Double
        Dim t_av As Double
        'PVT properties
        Dim N As Double
        Dim wi_deriv2 As Double
        Dim wi_2_deriv As Double

        'Calculate specific heat capacities
        With fluid
            Call .Calc_PVT(p_u, t_u)

            'Calculate choke throat temperature
            t_C = (t_u + 273) * P_r ^ (1 - 1 / .Polytropic_exponent) - 273
            t_choke_throat_C_ = t_C
            'Calculate average pressure and temperature
            'p_av = (p_u + P_r * p_u) / 2
            t_av = (t_u + t_C) / 2
            t_choke_av_C_ = t_av
            wi_calc = wi_calc_(P_r, p_u, t_av, wi_deriv)
            If p_r_inc > 0 Then
                P_r = P_r + p_r_inc
                Call wi_calc_(P_r, p_u, t_av, wi_deriv2)
                wi_2_deriv = (wi_deriv2 - wi_deriv) / p_r_inc
                d_pr = -wi_deriv / wi_2_deriv
            End If
        End With
    End Function

    Private Function wi_calc_(P_r As Double,
                          p_u As Double,
                          t_av As Double,
          Optional ByRef wi_deriv As Double = -1)

        Dim alpha As Double
        Dim lambda As Double
        Dim betta As Double
        Dim GAMMA As Double
        Dim Delta As Double
        Dim p_av As Double
        'Calculate average pressure and temperature
        p_av = (p_u + P_r * p_u) / 2
        With fluid
            Call .Calc_PVT(p_av, t_av)
            If P_r = 0 Then
                P_r = 0.000001
            End If

            alpha = .Rho_gas_rc_kgm3 * (.Fm_oil_fr / .Rho_oil_rc_kgm3 + .Fm_wat_fr / .Rho_wat_rc_kgm3)
            'Calculate auxilary values
            lambda = (.Fm_gas_fr + (.Fm_gas_fr * .Cv_gas_JkgC + .Fm_oil_fr * .Cv_oil_JkgC + .Fm_wat_fr * .Cv_wat_JkgC) / (.Cv_gas_JkgC * (.Heat_capacity_ratio_gas - 1)))
            betta = .Fm_gas_fr / .Polytropic_exponent * P_r ^ (-1 - 1 / .Polytropic_exponent)
            GAMMA = .Fm_gas_fr + alpha
            Delta = .Fm_gas_fr * P_r ^ (-1 / .Polytropic_exponent) + alpha

            P_r = Min(P_r, 1)

            wi_calc_ = 27500000.0# * a_c * (2 * p_u * .Rho_gas_rc_kgm3 / Delta ^ 2 * (lambda * (1 - P_r ^ (1 - 1 / .Polytropic_exponent)) + alpha * (1 - P_r)) _
                       / (1 - (a_r * GAMMA / Delta) ^ 2)) ^ (1 / 2)

            'Calculate rate derivative
            wi_deriv = (2 * lambda * (1 - P_r ^ (1 - 1 / .Polytropic_exponent)) + 2 * alpha * (1 - P_r)) * betta -
                Delta * (1 - (a_r * GAMMA / Delta) ^ 2) * (lambda * (1 - 1 / .Polytropic_exponent) * P_r ^ (-1 / .Polytropic_exponent) + alpha)
        End With
    End Function


    'Function calculates downstream node pressure for choke
    Public Function calc_choke_p_lin(PTbuf As PTtype) As PTtype
        'PTbuf - well head pressure (upstream) ( (atma)) and temperature ( (C))
        'Return downstream pressure and temperature

        ' если расчет не возможен (решение не существует), возвращает 0, так как
        ' потенциально может возникнуть ситуация, что при заданном дебите, диаметре штуцера и
        ' получившемся давлении на входе - решения по давлению на выходе не будет существовать
        'PTbuf - well head pressure and  temperature Upstream

        Try
            Dim eps As Double
            Dim eps_q As Double
            eps = 0.001
            eps_q = 0.1
            If (d_choke_m > d_up_m - 2 * eps) Or (d_choke_m < 0.001) Or (fluid.qliq_sm3day < eps_q) Then
                calc_choke_p_lin = PTbuf
                Exit Function
            End If
            ' Если при расчете линейного давления возникла ошибка, то скорее всего для дебита нет соотвествия для линейного давления
            calc_choke_p_lin = calc_choke_p(PTbuf, calc_p_down:=1)
            Exit Function
        Catch ex As Exception
            calc_choke_p_lin = Set_PT(0, 0)
            Dim errmsg As String
            errmsg = "Cchoke.calc_choke_plin_atma: error. set calc_choke_plin_atma = 0 : pbuf_atma  = " & PTbuf.p_atma & "  t_choke_C = " & PTbuf.t_C
            AddLogMsg(errmsg)
            Throw New ApplicationException(errmsg)
        End Try

    End Function

    Public Function calc_choke_p(pt As PTtype, Optional calc_p_down As Integer = 0) As PTtype
        'Function calculates end node pressure for choke (weather upstream or downstream)
        Dim p_sn As Double, t_u As Double
        Dim P_en As Double
        Dim counter As Double
        Dim eps As Double
        Dim eps_p As Double
        Const max_iters As Integer = 25
        Dim void As Double
        Dim q_l As Double
        Dim P_en_min As Double
        Dim P_en_max As Double
        Dim i As Integer

        Dim q_good As Boolean
        Dim p_good As Boolean

        Try
            p_sn = pt.p_atma
            t_u = pt.t_C
            counter = 0

            eps = fluid.qliq_sm3day * 0.0001 'set precision equal to 0.01%
            eps_p = const_pressure_tolerance

            If (calc_p_down = 0) Then 'Calculate upstream pressure given downstream
                'Solve for upstream pressure
                i = 1
                counter = 0
                Do
                    ' ищем давление на входе заведомо превышающее необходимое для обеспечения заданного потока
                    counter = counter + 1
                    i = 2 * i
                    P_en_max = p_sn * i
                    q_l = calc_choke_qliq_sm3day(P_en_max, p_sn, t_u)
                Loop Until q_l > fluid.qliq_sm3day Or counter > max_iters

                If q_l <= fluid.qliq_sm3day Then   ' значит поиск дебита не увенчался успехом
                    AddLogMsg("calc_choke_P(calc_p_down = 0): no solution found for rate = " & String.Format("{0:0.##}", fluid.qliq_sm3day))
                End If

                ' определим нижнюю границу поиска давления
                P_en_min = i * p_sn / 2
                counter = 0
                Do
                    ' ищем точное значение давления на входе обеспечивающего дебит
                    ' потенциально можно ускорить если не делить отрезок пополам а использовать линейное приближение (характеристика должна быть довольно гладкой)
                    counter = counter + 1
                    P_en = (P_en_min + P_en_max) / 2
                    q_l = calc_choke_qliq_sm3day(P_en, p_sn, t_u)
                    If q_l > fluid.qliq_sm3day Then
                        P_en_max = P_en
                    Else
                        P_en_min = P_en
                    End If
                    q_good = Abs(fluid.qliq_sm3day - q_l) < eps
                    p_good = Abs(P_en_min - P_en_max) < eps_p
                Loop Until (q_good And p_good) Or counter > max_iters

                If (counter > max_iters) And (Abs(fluid.qliq_sm3day - q_l) > eps * 100) Then ' значит поиск дебита не увенчался успехом
                    AddLogMsg("calc_choke_P(calc_p_down = 0): number of iterations too much, no solution found for rate = " & String.Format("{0:0.##}", fluid.qliq_sm3day))
                End If
            End If
            Dim p_cr As Double
            If (calc_p_down = 1) Then 'Calculate downstream pressure given upstream
                'Solve for upstream pressure
                'Calculate critical oil rate
                q_l = calc_choke_qliq_sm3day(p_sn, 0, t_u)
                If (fluid.qliq_sm3day - q_l) > 0.0001 Then 'Given oil rate can't be archieved
                    P_en = -1
                Else
                    If (q_l - fluid.qliq_sm3day) < 0.0001 Then
                        calc_choke_p = Set_PT(0, 0)
                        P_en = 0
                    Else
                        i = 1
                        counter = 0
                        Do
                            i = 2 * i
                            P_en_min = p_sn / i
                            q_l = calc_choke_qliq_sm3day(p_sn, P_en_min, t_u)
                        Loop Until q_l > fluid.qliq_sm3day Or counter > max_iters

                        If q_l <= fluid.qliq_sm3day Then   ' значит поиск дебита не увенчался успехом
                            AddLogMsg("calc_choke_P(calc_p_down = 1):no solution found for rate = " & String.Format("{0:0.##}", fluid.qliq_sm3day))
                        End If
                        P_en_max = 2 * p_sn / i
                        counter = 0
                        Do
                            counter = counter + 1
                            P_en = (P_en_min + P_en_max) / 2
                            q_l = calc_choke_qliq_sm3day(p_sn, P_en, t_u)
                            If q_l > fluid.qliq_sm3day Then
                                P_en_min = P_en
                            Else
                                P_en_max = P_en
                            End If
                        Loop Until Abs(fluid.qliq_sm3day - q_l) < eps Or counter > max_iters
                        If counter > max_iters Then   ' значит поиск дебита не увенчался успехом
                            AddLogMsg("calc_choke_P(calc_p_down = 1): number of iterations exeeded, no solution found for rate = " & String.Format("{0:0.##}", fluid.qliq_sm3day))
                        End If
                    End If
                End If
            End If
            calc_choke_p.p_atma = P_en
            calc_choke_p.t_C = t_u    ' пока предполагаем для штуцера температура не меняется

            Exit Function
        Catch ex As Exception
            Dim errmsg As String
            errmsg = "CChoke.calc_choke_P: error"
            AddLogMsg(errmsg)
            Throw New ApplicationException(errmsg)
        End Try

    End Function

    'Function calculates upstream node pressure for choke
    Public Function calc_choke_p_buf(PTLine As PTtype) As PTtype
        'Arguments
        'PTline_atma - line pressure (downstream) ( (atma)) and temperature ( (C))
        'Return upstream pressure and temperature

        Dim eps As Double
        Dim eps_q As Double
        eps = 0.001
        eps_q = 0.1
        Try
            If (d_choke_m > d_up_m - 2 * eps) Or (d_choke_m < 0.001) Or (fluid.qliq_sm3day < eps_q) Then
                calc_choke_p_buf = PTLine
                Exit Function
            End If
            calc_choke_p_buf = calc_choke_p(PTLine, 0)
            Exit Function
        Catch ex As Exception
            calc_choke_p_buf = Set_PT(0, 0)
            Dim errmsg As String
            errmsg = "Cchoke.calc_choke_p_buf: error. set calc_choke_p_buf = 0 : p_line_atma  = " & PTLine.p_atma & "  t_choke_C = " & PTLine.t_C
            AddLogMsg(errmsg)
            Throw New ApplicationException(errmsg)
        End Try

    End Function
End Class
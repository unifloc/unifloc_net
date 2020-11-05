﻿'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
' функции расчета скважины для проведения расчетов из интерфейса Excel

Option Explicit On
Imports System.Math
Public Module u7_Excel_functions_GL

    ' ==============  функции для расчета скважины ==========================
    ' =====================================================================

    'Private Function wellGL_InitData(q_m3day As Double, _
    '                                 fw_perc As Double, _
    '                        Optional p_cas_atma As Double = 10, _
    '                        Optional str_well As String = WELL_GL_DEFAULT, _
    '                        Optional str_PVT As String = PVT_DEFAULT, _
    '                        Optional hydr_corr As H_CORRELATION = 0 _
    '                                 ) As CWellGL
    '    ' функция для шаблонного чтения данных по скважине в интерфейсных функциях
    '    '
    '    ' на входе данные по конструкции скважины, PVT, ГЛ берутся из строк с закодированными параметрами
    '    '    следует использовать функции Encode для передачи параметров
    '    ' на выходе объект скважина с загруженными данными
    '    Dim well As New CWellGL
    '    Dim PVT As New CPVT
    '
    '    Set PVT = PVT_decode_string(str_PVT)
    '    PVT.qliq_sm3day = q_m3day
    '    PVT.fw_fr = fw_perc / 100
    '
    '    Set well = wellGL_decode_string(str_well)
    '    Set well.fluid = PVT
    '  '  well.p_cas_atma = p_cas_atma
    '    well.hydraulic_correlation = hydr_corr
    '
    '    Set wellGL_InitData = well
    'End Function
    '
    'Private Function wellGL_out_arr(well As CWellGL, Optional FirsrCol As Integer = 0)
    '    Dim ar1(), ar2()
    '    Dim vlv As CGLvalve
    '    With well
    '        ' подготовим данные для вывода
    '        ' данные выводятся в одну линию, чтобы, хотя бы теоретически, можно было вывести данные по скважине в таблице
    '        ' во второй строке выводятся подписи параметров, если необходимо
    '
    '        Dim i As Integer, j As Integer
    '
    '        ' первый параметр настраиваемый
    '        i = 1: ReDim Preserve ar1(i): ReDim Preserve ar2(i)
    '        ar1(0) = "":  ar2(0) = ""
    '        ' блок параметров по давлениям
    '        i = i: ReDim Preserve ar1(i): ReDim Preserve ar2(i)
    '        ar1(i) = .p_line_atma:  ar2(i) = "p_line_atma"
    '        i = i + 1: ReDim Preserve ar1(i): ReDim Preserve ar2(i)
    '        ar1(i) = .pbuf_atma:  ar2(i) = "pbuf_atma"
    '        i = i + 1: ReDim Preserve ar1(i): ReDim Preserve ar2(i)
    '        ar1(i) = .p_cas_atma:  ar2(i) = "p_cas_atma"
    '        i = i + 1: ReDim Preserve ar1(i): ReDim Preserve ar2(i)
    '        ar1(i) = .p_gas_inj_atma:   ar2(i) = "Pgas_inj_atma"
    '        i = i + 1: ReDim Preserve ar1(i): ReDim Preserve ar2(i)
    '        ar1(i) = .pwf_atma:  ar2(i) = "pwf_atma"
    '
    '        For j = 1 To .valves.Count
    '            Set vlv = .valves.valves(j)
    '            i = i + 1: ReDim Preserve ar1(i): ReDim Preserve ar2(i)
    '            ar1(i) = vlv.p_in_atma:   ar2(i) = "GLV" + CStr(j) + ".p_in_atma"
    '            i = i + 1: ReDim Preserve ar1(i): ReDim Preserve ar2(i)
    '            ar1(i) = vlv.p_out_atma:   ar2(i) = "GLV" + CStr(j) + ".p_out_atma"
    '            i = i + 1: ReDim Preserve ar1(i): ReDim Preserve ar2(i)
    '            ar1(i) = vlv.h_mes_m:    ar2(i) = "GLV" + CStr(j) + ".h_mes_m"
    '            i = i + 1: ReDim Preserve ar1(i): ReDim Preserve ar2(i)
    '            ar1(i) = vlv.q_gas_inj_scm3day:     ar2(i) = "GLV" + CStr(j) + ".q_gas_inj_scm3day"
    '
    '        Next j
    '        ' параметры температуры
    '        i = i + 1: ReDim Preserve ar1(i): ReDim Preserve ar2(i)
    '        ar1(i) = .t_buf_C:  ar2(i) = "t_buf_C"
    '        i = i + 1: ReDim Preserve ar1(i): ReDim Preserve ar2(i)
    '        ar1(i) = .t_bh_C:  ar2(i) = "t_bh_C"
    '
    '        ar1(0) = ar1(FirsrCol): ar2(0) = ar2(FirsrCol)
    '
    '        wellGL_out_arr = Array(ar1, ar2)
    '    ' можно еще добавить сюда вывод кривых распределения давления и температуры по стволу и еще 4 параметров  (потом)
    '    End With
    'End Function

    'Public Function wellGL_plin_pwf_atma(ByVal pwf_atma As Double, _
    '                                     ByVal q_m3day As Double, _
    '                                     ByVal fw_perc As Double, _
    '                                     Optional ByVal p_cas_atma As Double = 10, _
    '                                     Optional qgas_inj_scm3day As Double = -1, _
    '                                     Optional str_well As String = WELL_GL_DEFAULT, _
    '                                     Optional str_PVT As String = PVT_DEFAULT, _
    '                                     Optional ByVal hydr_corr As H_CORRELATION = 0)
    '' функция расчета линейного давления скважины по забойному
    '    Dim well As CWellGL
    '    Set well = wellGL_InitData(q_m3day, fw_perc, p_cas_atma, _
    '                               str_well, str_PVT, hydr_corr)
    '    Call well.set_qgas_inj(p_cas_atma, qgas_inj_scm3day)
    '
    '    Call well.calc_plin_pwf_atma(pwf_atma)           ' проведем расчет
    '    ' в качестве результата выведем ряд расчитанных значений давления
    '    wellGL_plin_pwf_atma = wellGL_out_arr(well, 1)
    '
    'End Function
    '
    'Public Function wellGL_pwf_plin_atma(ByVal plin_atma As Double, _
    '                                     ByVal q_m3day As Double, _
    '                                     ByVal fw_perc As Double, _
    '                                     Optional ByVal p_cas_atma As Double = 10, _
    '                                     Optional qgas_inj_scm3day As Double = -1, _
    '                                     Optional str_well As String = WELL_GL_DEFAULT, _
    '                                     Optional str_PVT As String = PVT_DEFAULT, _
    '                                     Optional ByVal hydr_corr As H_CORRELATION = 0)
    '' функция расчета линейного давления скважины по забойному
    '    Dim well As CWellGL
    '    Set well = wellGL_InitData(q_m3day, fw_perc, p_cas_atma, _
    '                               str_well, str_PVT, hydr_corr)
    '    Call well.set_qgas_inj(p_cas_atma, qgas_inj_scm3day)
    '    Call well.calc_pwf_plin_atma(plin_atma, well.t_bh_C)            ' проведем расчет
    '    ' в качестве результата выведем ряд расчитанных значений давления
    '    wellGL_pwf_plin_atma = wellGL_out_arr(well, 5)
    '
    'End Function

    ' function to calculated gas passage trough orifice or gas valve
    ' link in K Brawn AL 2A - Craft, Holden, Graves (p.111)
    ' also found in Mischenko book

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' функция расчета расхода газа через газлифтный клапан/штуцер
    ' результат массив значений и подписей
    Public Function GLV_q_gas_sm3day(ByVal d_mm As Double,
                                 ByVal p_in_atma As Double,
                                 ByVal p_out_atma As Double,
                                 ByVal gamma_g As Double,
                                 ByVal t_C As Double,
                        Optional ByVal c_calibr As Double = 1)
        ' d_mm        - диаметр основного порта клапана, мм
        ' p_in_atma   - давление на входе в клапан (затруб), атма
        ' p_out_atma  - давление на выходе клапана (НКТ), атма
        ' gamma_g     - удельная плотность газа
        ' t_C         - температура клапана, С
        'description_end

        Try
            Dim K As Double
            Dim d_in As Double
            Dim Pu_psi As Double
            Dim Pd_psi As Double
            Dim Tu_F As Double
            Dim Pd_Pu_crit As Double
            Dim cd As Double  ' discharge coefficient
            Dim g As Double
            Dim C0 As Double, C1 As Double, C2 As Double
            Dim a As Double
            Dim Qg_crit As Double
            Dim Qg As Double
            Dim Pd_Pu As Double
            Dim crit As Boolean
            Dim p_crit_out_atma As Double

            crit = False
            Pd_Pu = p_out_atma / p_in_atma

            If Pd_Pu >= 1 Then
                'Dim new_array(1) As Object
                'new_array(0) = (0, 0, crit)
                'new_array(1) = ("q_gas_sm3day", "p_crit_atma", "critical flow")
                'GLV_q_gas_sm3day = new_array
                GLV_q_gas_sm3day = {0, 0, crit}
                'GLV_q_gas_sm3day = Join(GLV_q_gas_sm3day)

                Exit Function
            End If

            If Pd_Pu <= 0 Then
                GLV_q_gas_sm3day = 0
                Exit Function
            End If

            K = 1.31   ' = Cp/Cv (approx 1.31 for natural gases(R Brown) or 1.25 (Mischenko) )
            K = UnfClassLibrary.Unf_pvt_gas_heat_capacity_ratio(gamma_g, t_C + UnfClassLibrary.const_t_K_zero_C)

            d_in = d_mm * 0.03937
            a = UnfClassLibrary.const_Pi * d_in ^ 2 / 4         'area of choke, sq in.
            Pu_psi = p_in_atma * 14.2233          'upstream pressure, psi
            Pd_psi = p_out_atma * 14.2233          'downstream pressure, psi
            Tu_F = t_C / 100 * 180 + 32
            Pd_Pu_crit = (2 / (K + 1)) ^ (K / (K - 1))
            cd = 0.865
            g = 32.17 'ft/sec^2

            C1 = (Pd_Pu_crit ^ (2 / K) - Pd_Pu_crit ^ (1 + 1 / K)) ^ 0.5
            C2 = (2 * g * K / (K - 1)) ^ 0.5
            Qg_crit = 155.5 * cd * a * Pu_psi * C1 * C2 / (gamma_g * (Tu_F + 460)) ^ 0.5 'critical flow ratio, Mcf/d
            Qg_crit = Qg_crit * c_calibr
            p_crit_out_atma = p_in_atma * Pd_Pu_crit

            If Pd_Pu <= Pd_Pu_crit Then
                Qg = Qg_crit * 28.31993658
                p_out_atma = p_crit_out_atma
                crit = True
            Else
                C0 = ((Pd_Pu ^ (2 / K) - Pd_Pu ^ (1 + 1 / K))) ^ 0.5
                Qg = Qg_crit * 28.31993658 * C0 / C1
                crit = False
            End If
            'Dim new_array(1) As Object
            'new_array(0) = (Qg, p_crit_out_atma, crit)
            'new_array(1) = ("q_gas_sm3day", "p_crit_atma", "critical flow")
            'GLV_q_gas_sm3day = new_array
            GLV_q_gas_sm3day = {Qg, p_crit_out_atma, crit}
            'GLV_q_gas_sm3day = Join(GLV_q_gas_sm3day)

        Catch ex As Exception
            GLV_q_gas_sm3day = -1
            Dim errmsg As String
            errmsg = "error in function : GL_qgas_valve_sm3day"
            Throw New ApplicationException(errmsg)

        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' функция расчета расхода газа через газлифтный клапан
    ' с учетом наличия вкруток на выходе клапана.
    ' результат массив значений и подписей.
    Public Function GLV_q_gas_vkr_sm3day(d_port_mm As Double,
                                     d_vkr_mm As Double,
                                     p_in_atma As Double,
                                     p_out_atma As Double,
                                     gamma_g As Double,
                                     t_C As Double)
        ' d_port_mm - диаметр основного порта клапана, мм
        ' d_vkr_mm  - эффективный диаметр вкруток на выходе, мм
        ' p_in_atma   - давление на входе в клапан (затруб), атма
        ' p_out_atma   - давление на выходе клапана (НКТ), атма
        ' gamma_g   - удельная плотность газа
        ' t_C       - температура клапана, С
        'description_end
        Try
            Dim prm As New UnfClassLibrary.CSolveParam
            prm.Class_Initialize()
            Dim CoeffA(5) As Object
            Dim Func As String
            Dim pv_atma As Double
            Dim q_gas_sm3day As Double
            Dim res1
            Dim res2
            Dim crit1 As Boolean
            Dim crit2 As Boolean

            Func = "calc_dq_gas_pv_vkr_valve"

            CoeffA(0) = d_port_mm
            CoeffA(1) = d_vkr_mm
            CoeffA(2) = p_in_atma
            CoeffA(3) = p_out_atma
            CoeffA(4) = gamma_g
            CoeffA(5) = t_C
            prm.y_tolerance = 0.01

            Call UnfClassLibrary.solve_equation_bisection(Func, p_out_atma, p_in_atma, CoeffA, prm)
            pv_atma = prm.x_solution
            res1 = GLV_q_gas_sm3day(d_port_mm, p_in_atma, pv_atma, gamma_g, t_C)
            res2 = GLV_q_gas_sm3day(d_vkr_mm, pv_atma, p_out_atma, gamma_g, t_C)
            q_gas_sm3day = res1(0) '(0)(0)
            crit1 = res1(2) '(0)(2)
            crit2 = res2(2) '(0)(2)

            'Dim new_array(1) As Object
            'new_array(0) = (q_gas_sm3day, p_in_atma, pv_atma, p_out_atma, q_gas_sm3day, crit1, crit2)
            'new_array(1) = ("q_gas_sm3day", "p_in_atma", "pv_atma", "p_out_atma", "q_gas_sm3day", "crit1", "crit2")
            'GLV_q_gas_vkr_sm3day = new_array
            GLV_q_gas_vkr_sm3day = {q_gas_sm3day, p_in_atma, pv_atma, p_out_atma, q_gas_sm3day, crit1, crit2}
            'GLV_q_gas_vkr_sm3day = Join(GLV_q_gas_vkr_sm3day)

        Catch ex As Exception
            Dim errmsg As String
            errmsg = "error in function : GLV_q_gas_vkr_sm3day"
            Throw New ApplicationException(errmsg)
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' функция расчета давления на входе или на выходе
    ' газлифтного клапана (простого) при закачке газа.
    ' результат массив значений и подписей
    Public Function GLV_p_vkr_atma(ByVal d_port_mm As Double,
                               ByVal d_vkr_mm As Double,
                               ByVal p_calc_atma As Double,
                               ByVal q_gas_sm3day As Double,
                     Optional ByVal gamma_g As Double = 0.6,
                     Optional ByVal t_C As Double = 25,
                     Optional ByVal calc_along_flow As Boolean = False)
        ' d_port_mm     - диаметр порта клапана, мм
        ' d_vkr_mm      - диаметр вкрутки клапана, мм
        ' p_calc_atma   - давление на входе (выходе) клапана, атма
        ' q_gas_sm3day  - расход газа, ст. м3/сут
        ' gamma_g       - удельная плотность газа
        ' t_C           - температура в точке установки клапана
        ' calc_along_flow - направление расчета:
        '              0 - против потока (расчет давления на входе);
        '              1 - по потоку (расчет давления на выходе).
        'description_end
        ' ищем давление внутри клапана
        Dim p_v_atma As Double
        Dim p_in As Double
        Dim p_out As Double
        Dim p_atma As Double
        Dim p2
        Dim p1
        Dim crit1 As Boolean
        Dim crit2 As Boolean
        Dim qg0 As Double
        qg0 = q_gas_sm3day

        Try
            crit1 = False
            crit2 = False
            If calc_along_flow Then
                p_in = p_calc_atma
                p1 = GLV_p_atma(d_port_mm, p_in, q_gas_sm3day, gamma_g, t_C, True)
                p_v_atma = p1(0) '(0)(0)
                If p_v_atma < 0 Then
                    ' critical flow through the port achived
                    q_gas_sm3day = p1(1) '(0)(1)
                    p_v_atma = p1(2) '(0)(2)
                    crit1 = True
                End If

                If d_vkr_mm > 0 Then
                    p2 = GLV_p_atma(d_vkr_mm, p_v_atma, q_gas_sm3day, gamma_g, t_C, True)
                    p_atma = p2(0) '(0)(0)
                    If p_atma < 0 Then
                        ' critical flow through the vkrutka achived
                        q_gas_sm3day = p2(1) '(0)(1)
                        p_atma = p2(2) '(0)(2)
                        crit2 = True
                    End If
                Else
                    p_atma = p_v_atma
                End If
                p_out = p_atma
                If q_gas_sm3day < qg0 Then
                    p_atma = -1
                End If
            Else
                p_out = p_calc_atma
                If d_vkr_mm > 0 Then
                    p1 = GLV_p_atma(d_vkr_mm, p_calc_atma, q_gas_sm3day, gamma_g, t_C, False)
                    p_v_atma = p1(0) '(0)(0)
                    If p_v_atma < 0 Then
                        ' critical flow through the vkrutka achived
                        q_gas_sm3day = p1(1) '(0)(1)
                        p_v_atma = p1(2) '(0)(2)
                        crit2 = True
                    End If
                Else
                    p_v_atma = p_calc_atma
                End If
                p2 = GLV_p_atma(d_port_mm, p_v_atma, q_gas_sm3day, gamma_g, t_C, False)
                p_atma = p2(0) '(0)(0)
                If p_atma < 0 Then
                    ' critical flow through the port achived
                    q_gas_sm3day = p2(1) '(0)(1)
                    p_atma = p2(2) '(0)(2)
                    crit1 = True
                End If
                p_in = p_atma
            End If

            'Dim new_array(1) As Object
            'new_array(0) = (p_atma, p_in, p_v_atma, p_out, q_gas_sm3day, crit1, crit2)
            'new_array(1) = ("p_atma", "p_in_atma", "p_v_atma", "p_out_atma", "q_gas_sm3day", "port critical flow", "vkrutka critical flow")
            'GLV_p_vkr_atma = new_array
            GLV_p_vkr_atma = {p_atma, p_in, p_v_atma, p_out, q_gas_sm3day, crit1, crit2}
            'GLV_p_vkr_atma = Join(GLV_p_vkr_atma)
            Exit Function
        Catch ex As Exception
            GLV_p_vkr_atma = "error"
            Dim errmsg As String
            errmsg = "error in function : GLV_p_vkr_atma"
            Throw New ApplicationException(errmsg)
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' функция расчета давления на входе или на выходе
    ' газлифтного клапана (простого) при закачке газа.
    ' результат массив значений и подписей
    Public Function GLV_p_atma(ByVal d_mm As Double,
                           ByVal p_calc_atma As Double,
                           ByVal q_gas_sm3day As Double,
                           Optional ByVal gamma_g As Double = 0.6,
                           Optional ByVal t_C As Double = 25,
                           Optional ByVal calc_along_flow As Boolean = False,
                           Optional ByVal p_open_atma As Double = 0,
                           Optional ByVal c_calibr As Double = 1)
        ' d_mm          - диаметр клапана, мм
        ' p_calc_atma   - давление на входе (выходе) клапана, атма
        ' q_gas_sm3day  - расход газа, ст. м3/сут
        ' gamma_g       - удельная плотность газа
        ' t_C           - температура в точке установки клапана
        ' calc_along_flow - направление расчета:
        '              0 - против потока (расчет давления на входе);
        '              1 - по потоку (расчет давления на выходе).
        ' p_open_atma    - давление открытия/закрытия клапана, атм
        'description_end

        Try
            Dim Qmax_m3day As Double
            Dim qres
            Dim pd As Double
            Dim Pu As Double
            Dim Pcrit As Double
            Dim K As Double
            Dim Pd_Pu_crit As Double
            Dim crit As Boolean

            Dim prm As New UnfClassLibrary.CSolveParam
            prm.Class_Initialize()
            Dim CoeffA(5) As Object
            Dim Func As String

            K = 1.31   ' = Cp/Cv (approx 1.31 for natural gases(R Brown) or 1.25 (Mischenko) )
            Pd_Pu_crit = (2 / (K + 1)) ^ (K / (K - 1))
            CoeffA(0) = q_gas_sm3day
            CoeffA(1) = d_mm
            CoeffA(3) = gamma_g
            CoeffA(4) = t_C
            CoeffA(5) = c_calibr
            prm.y_tolerance = 0.1

            If calc_along_flow Then
                Pu = p_calc_atma
                pd = 1
                qres = GLV_q_gas_sm3day(d_mm, Pu, pd, gamma_g, t_C)
                Qmax_m3day = qres(0) '(0)(0)
                Pcrit = pd
                If Qmax_m3day > q_gas_sm3day And Pu > p_open_atma Then
                    Func = "calc_dq_gas_pd_valve"
                    CoeffA(2) = Pu
                    crit = False
                    Call UnfClassLibrary.solve_equation_bisection(Func, Pd_Pu_crit * Pu, Pu, CoeffA, prm)

                    'Dim new_array(1) As Object
                    'new_array(0) = (prm.x_solution, Qmax_m3day, Pcrit, crit)
                    'new_array(1) = ("p", "Qmax_m3day", "Pcrit", "critical flow")
                    'GLV_p_atma = new_array
                    GLV_p_atma = {prm.x_solution, Qmax_m3day, Pcrit, crit}
                    'GLV_p_atma = Join(GLV_p_atma)
                Else
                    crit = True

                    'Dim new_array(1) As Object
                    'new_array(0) = (-1, Qmax_m3day, Pcrit, crit)
                    'new_array(1) = ("p, atma", "Qmax_m3day", "Pcrit", "critical flow")
                    'GLV_p_atma = new_array
                    GLV_p_atma = {-1, Qmax_m3day, Pcrit, crit}
                    'GLV_p_atma = Join(GLV_p_atma)
                End If
            Else
                Qmax_m3day = q_gas_sm3day
                pd = p_calc_atma
                Pu = 500
                Func = "calc_dq_gas_pu_valve"
                CoeffA(2) = pd
                crit = False
                Call UnfClassLibrary.solve_equation_bisection(Func, pd, Pu, CoeffA, prm)
                Dim sol As Double
                sol = prm.x_solution
                If sol < p_open_atma Then
                    sol = p_open_atma
                End If

                'Dim new_array(1) As Object
                'new_array(0) = (sol, prm.x_solution, prm.y_solution, Pu, crit)
                'new_array(1) = ("p_opo_atma", "p, atma", "Q_m3day", "Pu max", "critical flow")
                'GLV_p_atma = new_array
                GLV_p_atma = {sol, prm.x_solution, prm.y_solution, Pu, crit}
                'GLV_p_atma = Join(GLV_p_atma)
            End If

            Exit Function
        Catch ex As Exception
            GLV_p_atma = "error"
            Dim errmsg As String
            errmsg = "error in function : GLV_p_atma"
            Throw New ApplicationException(errmsg)
        End Try


    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' функция расчета давления зарядки сильфона на стенде при
    ' стандартной температуре по данным рабочих давления и температуры
    Public Function GLV_p_bellow_atma(ByVal p_atma As Double,
                                  ByVal t_C As Double) As Double
        ' p_atma - рабочее давление открытия клапана в скважине, атм
        ' t_C   - рабочая температура открытия клапана в скважине, С
        'description_end

        Dim t_F As Double
        Dim Ct As Double
        Dim M As Double
        Dim Pb_psia As Double
        If p_atma > 1 Then
            Pb_psia = p_atma * 14.696
            t_F = t_C * 9 / 5 + 32
            If Pb_psia < 1238 Then
                M = 0.0000003054 * Pb_psia ^ 2 + 0.001934 * Pb_psia - 0.00226
            Else
                M = 0.000000184 * Pb_psia ^ 2 + 0.002298 * Pb_psia - 0.267
            End If
            Ct = 1 / (1 + (t_F - 60) * M / Pb_psia)
            GLV_p_bellow_atma = p_atma * Ct
        End If

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' фукнция расчета давления в сильфоне с азотом
    ' в рабочих условиях при заданной температуре
    Public Function GLV_p_close_atma(ByVal p_bellow_atm As Double,
                                 ByVal t_C As Double) As Double
        ' p_bellow_atm  - давление зарядки сильфона при стандартных условиях
        ' t_C           - температура рабочая
        'description_end


        Try
            'Dim p_psi As Double
            Dim t_F As Double
            Dim Ct As Double
            Dim M As Double
            Dim Pb_psia As Double


            Pb_psia = p_bellow_atm * 14.696
            t_F = t_C * 9 / 5 + 32

            If Pb_psia < 1238 Then
                M = 0.0000003054 * Pb_psia ^ 2 + 0.001934 * Pb_psia - 0.00226
            Else
                M = 0.000000184 * Pb_psia ^ 2 + 0.002298 * Pb_psia - 0.267
            End If

            Ct = 1 / (1 + (t_F - 60) * M / Pb_psia)

            GLV_p_close_atma = p_bellow_atm / Ct
            Exit Function
        Catch ex As Exception
            GLV_p_close_atma = 0
        End Try

    End Function


    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    'Функция расчета диаметра порта клапана
    'на основе уравнения Thornhill-Crave
    Public Function GLV_d_choke_mm(ByVal q_gas_sm3day As Double,
                               ByVal p_in_atma As Double,
                               ByVal p_out_atma As Double,
                               Optional ByVal gamma_g As Double = 0.6,
                               Optional ByVal t_C As Double = 25)
        ' q_gas_sm3day  - расход газа, ст. м3/сут
        ' p_in_atma   - давление на входе в клапан (затруб), атма
        ' p_out_atma   - давление на выходе клапана (НКТ), атма
        ' gamma_g   - удельная плотность газа
        ' t_C       - температура клапана, С
        'description_end


        Try
            If q_gas_sm3day <= 0 Then
                GLV_d_choke_mm = 0
                Exit Function
            End If

            If p_in_atma < p_out_atma Then
                GLV_d_choke_mm = -1
                Exit Function
            End If

            Dim K As Double
            K = 1.31   ' = Cp/Cv (approx 1.31 for natural gases(R Brown) or 1.25 (Mischenko) )

            Dim Pu_psi As Double
            Dim Pd_psi As Double
            Pu_psi = p_in_atma * 14.2233 'upstream pressure, psi
            Pd_psi = p_out_atma * 14.2233 'downstream pressure, psi

            Dim Tu_F As Double
            Tu_F = t_C / 100 * 180 + 32

            Dim cd As Double  ' discharge coefficient
            cd = 0.865

            Dim g As Double
            g = 32.17 'ft/sec^2

            Dim Qg_Mcfd As Double
            Qg_Mcfd = q_gas_sm3day / 28.31993658

            Dim Pd_Pu_crit As Double
            Pd_Pu_crit = (2 / (K + 1)) ^ (K / (K - 1))

            Dim Pd_Pu As Double
            Pd_Pu = p_out_atma / p_in_atma

            Dim C0 As Double, C1 As Double, C2 As Double
            C0 = ((Pd_Pu ^ (2 / K) - Pd_Pu ^ (1 + 1 / K))) ^ 0.5
            C1 = (Pd_Pu_crit ^ (2 / K) - Pd_Pu_crit ^ (1 + 1 / K)) ^ 0.5
            C2 = (2 * g * K / (K - 1)) ^ 0.5

            Dim a As Double

            If Pd_Pu <= Pd_Pu_crit Then
                a = Qg_Mcfd / (155.5 * cd * Pu_psi * C1 * C2 / (gamma_g * (Tu_F + 460)) ^ 0.5)
            Else
                a = Qg_Mcfd / (155.5 * cd * Pu_psi * C0 * C2 / (gamma_g * (Tu_F + 460)) ^ 0.5)
            End If

            Dim d_in As Double
            d_in = (a * 4 / UnfClassLibrary.const_Pi) ^ 0.5  '(a * 4 / Application.Pi) ^ 0.5

            GLV_d_choke_mm = d_in / 0.03937


            Exit Function
        Catch ex As Exception
            GLV_d_choke_mm = -1
            Dim errmsg As String
            errmsg = "error in function : GL_dchoke_mm"
            Throw New ApplicationException(errmsg)
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    'Функция расчета давления открытия газлифтного клапана R1
    Public Function GLV_IPO_p_open(ByVal p_bellow_atma As Double,
                          ByVal p_out_atma As Double,
                          ByVal t_C As Double,
                 Optional ByVal GLV_type As Integer = 0,
                 Optional ByVal d_port_mm As Double = 5,
                 Optional ByVal d_vkr1_mm As Double = -1,
                 Optional ByVal d_vkr2_mm As Double = -1,
                 Optional ByVal d_vkr3_mm As Double = -1,
                 Optional ByVal d_vkr4_mm As Double = -1)
        ' p_bellow_atma - давление зарядки сильфона на стенде, атма
        ' p_out_atma    - давление на выходе клапана (НКТ), атма
        ' t_C           - температура клапана в рабочих условиях, С
        ' GLV_type      - тип газлифтного клапана (сейчас только R1)
        ' d_port_mm     - диаметр порта клапана
        ' d_vkr1_mm     - диаметр вкрутки 1, если есть
        ' d_vkr2_mm     - диаметр вкрутки 2, если есть
        ' d_vkr3_mm     - диаметр вкрутки 3, если есть
        ' d_vkr4_mm     - диаметр вкрутки 4, если есть
        'description_end

        Dim GLV As New UnfClassLibrary.CGLvalve
        GLV.Class_init()

        Call GLV.set_GLV_R1(True, d_port_mm, d_vkr1_mm, d_vkr2_mm, d_vkr3_mm, d_vkr4_mm)
        GLV.p_bellow_sc_atma = p_bellow_atma
        GLV.p_out_atma = p_out_atma
        GLV.t_C = t_C
        GLV_IPO_p_open = GLV.p_open_atma

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    'Функция расчета давления открытия газлифтного клапана R1
    Public Function GLV_IPO_p_atma(ByVal p_bellow_atma As Double,
                          ByVal d_port_mm As Double,
                          ByVal p_calc_atma As Double,
                          ByVal q_gas_sm3day As Double,
                          ByVal t_C As Double,
                 Optional ByVal calc_along_flow As Boolean = False,
                 Optional ByVal GLV_type As Integer = 0,
                 Optional ByVal d_vkr1_mm As Double = -1,
                 Optional ByVal d_vkr2_mm As Double = -1,
                 Optional ByVal d_vkr3_mm As Double = -1,
                 Optional ByVal d_vkr4_mm As Double = -1)
        ' p_bellow_atma - давление зарядки сильфона на стенде, атма
        ' p_out_atma    - давление на выходе клапана (НКТ), атма
        ' t_C           - температура клапана в рабочих условиях, С
        ' GLV_type      - тип газлифтного клапана (сейчас только R1)
        ' d_port_mm     - диаметр порта клапана
        ' d_vkr1_mm     - диаметр вкрутки 1, если есть
        ' d_vkr2_mm     - диаметр вкрутки 2, если есть
        ' d_vkr3_mm     - диаметр вкрутки 3, если есть
        ' d_vkr4_mm     - диаметр вкрутки 4, если есть
        'description_end

        Dim GLV As New UnfClassLibrary.CGLvalve
        GLV.Class_init()

        Call GLV.set_GLV_R1(True, d_port_mm, d_vkr1_mm, d_vkr2_mm, d_vkr3_mm, d_vkr4_mm)
        GLV.p_bellow_sc_atma = p_bellow_atma
        Dim res
        If calc_along_flow Then
            res = GLV.calc_p_out_atma(p_calc_atma, q_gas_sm3day)
        Else
            res = GLV.calc_p_in_atma(p_calc_atma, q_gas_sm3day)
        End If
        GLV.t_C = t_C
        GLV_IPO_p_atma = GLV.p_open_atma

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    'Функция расчета давления закрытия газлифтного клапана R1
    Public Function GLV_IPO_p_close(ByVal p_bellow_atma As Double,
                              ByVal p_out_atma As Double,
                              ByVal t_C As Double,
                 Optional ByVal GLV_type As Integer = 0,
                 Optional ByVal d_port_mm As Double = 5,
                 Optional ByVal d_vkr1_mm As Double = -1,
                 Optional ByVal d_vkr2_mm As Double = -1,
                 Optional ByVal d_vkr3_mm As Double = -1,
                 Optional ByVal d_vkr4_mm As Double = -1)
        ' p_bellow_atma - давление зарядки сильфона на стенде, атма
        ' p_out_atma    - давление на выходе клапана (НКТ), атма
        ' t_C           - температура клапана в рабочих условиях, С
        ' GLV_type      - тип газлифтного клапана (сейчас только R1)
        ' d_port_mm     - диаметр порта клапана
        ' d_vkr1_mm     - диаметр вкрутки 1, если есть
        ' d_vkr2_mm     - диаметр вкрутки 2, если есть
        ' d_vkr3_mm     - диаметр вкрутки 3, если есть
        ' d_vkr4_mm     - диаметр вкрутки 4, если есть
        'description_end
        Dim GLV As New UnfClassLibrary.CGLvalve
        GLV.Class_init 

        Call GLV.set_GLV_R1(True, d_port_mm, d_vkr1_mm, d_vkr2_mm, d_vkr3_mm, d_vkr4_mm)
        GLV.p_bellow_sc_atma = p_bellow_atma
        GLV.t_C = t_C
        GLV_IPO_p_close = GLV.p_open_atma
    End Function



    Function GL_dPgasPipe_atmg(ByVal h_m As Double, ByVal P_atmg As Double, ByVal t_C As Double,
                               Optional ByVal d_cas_mm As Double = 125,
                               Optional ByVal dtub_mm As Double = 73,
                               Optional ByVal gamma_gas As Double = 0.8,
                               Optional ByVal q_gas_scm3day As Double = 10000,
                               Optional ByVal roughness As Double = 0.001,
                               Optional ByVal THETA As Double = 90
                               ) As Double

        'de - external diameter, m
        'di - interior diameter, m
        'gamma_gas - relative density of gas
        'qg_sc - gas flow, m3/d
        'eps - pipe roughness, m
        'theta - ,degree
        'length - pipe length, m
        'T - temperature, C
        'P - pressure, atma
        Try
            Dim de, Di, qg_sc, eps, length, t, p
            de = d_cas_mm / 1000
            Di = dtub_mm / 1000

            qg_sc = q_gas_scm3day
            eps = roughness
            length = h_m
            t = t_C
            p = P_atmg

            'convert m3/d to scf/d
            qg_sc = qg_sc * 3.28 ^ 3

            Dim p_MPa As Double, p_psi As Double
            p_MPa = p * 0.1013 'convert atma to Mpa
            p_psi = p * 14.696 ' convert atma to psi


            Dim t_K As Double, t_F As Double
            t_K = t + 273 'convert Celcsius to Kelvin
            t_F = (9 / 5) * t + 32 'convert Celcsius to Fahrengheit

            'Dim T_pc As Double
            'Dim p_pc As Double
            Dim z As Double

            '        T_pc = PseudoTemperatureStanding(gamma_gas)
            '        p_pc = PseudoPressureStanding(gamma_gas)
            '        Z = ZFactorDranchuk(T_K / T_pc, P_MPa / p_pc)
            z = UnfClassLibrary.Unf_pvt_Zgas_d(t_K, p_MPa, gamma_gas)

            eps = eps * 39.3701 'convert m to in

            Dim de_in As Double, di_in As Double
            di_in = Di * 39.3701 'convert m to in
            de_in = de * 39.3701 'convert m to in

            Dim dh As Double, DA As Double, deq As Double
            dh = de_in - di_in
            DA = (de_in ^ 2 - di_in ^ 2) ^ 0.5

            If di_in = 0 Then
                deq = de_in
            Else
                deq = (de_in ^ 2 + di_in ^ 2 - (de_in ^ 2 - di_in ^ 2) / Log(de_in / di_in)) / (de_in - di_in)
            End If


            Dim mu_g As Double
            mu_g = UnfClassLibrary.Unf_pvt_viscosity_gas_cP(t_K, p_MPa, z, gamma_gas)

            Dim Re As Double
            Re = 0.020107 * gamma_gas * Abs(qg_sc) * deq / mu_g / DA ^ 2


            Dim a As Double, b As Double
            a = (2.457 * Log(1 / ((7 / Re) ^ 0.9 + 0.27 * eps / deq))) ^ 16
            b = (37530 / Re) ^ 16

            Dim f_moody As Double
            f_moody = 8 * ((8 / Re) ^ 12 + 1 / ((a + b) ^ 1.5)) ^ (1 / 12)

            Dim gradP As Double

            gradP = -0.018786 * gamma_gas * (p_psi + 14.7) * Sin(THETA * UnfClassLibrary.const_Pi / 180) / (t_F + 460) / z + (1.2595 * 10 ^ (-11)) * f_moody * (t_F + 460) * z * gamma_gas * (qg_sc ^ 2) / (p_psi + 14.7) / dh / DA ^ 4  ' applicacation,pi
            gradP = gradP * 0.068 / 0.3048 'convert psi/ft to atma/m

            GL_dPgasPipe_atmg = p + gradP * length

            Exit Function
        Catch ex As Exception
            GL_dPgasPipe_atmg = -1
            Dim errmsg As String
            errmsg = "error in function : GL_dPgasPipe_atmg"
            Throw New ApplicationException(errmsg)
        End Try

    End Function
End Module

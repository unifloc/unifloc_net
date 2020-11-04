Public Module u7_GL
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
                            Optional ByVal c_calibr As Double = 1) As Object()
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
            'Dim new_array(1) As Object

            If Pd_Pu >= 1 Then
                'new_array(0) = (0, 0, crit)
                'new_array(1) = ("q_gas_sm3day", "p_crit_atma", "critical flow")
                'GLV_q_gas_sm3day = new_array
                GLV_q_gas_sm3day = {0, 0, crit}
                'GLV_q_gas_sm3day = Join(GLV_q_gas_sm3day)

                Exit Function
            End If

            If Pd_Pu <= 0 Then
                GLV_q_gas_sm3day = {0}
                Exit Function
            End If

            K = 1.31   ' = Cp/Cv (approx 1.31 for natural gases(R Brown) or 1.25 (Mischenko) )
            K = Unf_pvt_gas_heat_capacity_ratio(gamma_g, t_C + const_t_K_zero_C)


            d_in = d_mm * 0.03937
            a = const_Pi * d_in ^ 2 / 4         'area of choke, sq in.
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

            'new_array(0) = (Qg, p_crit_out_atma, crit)
            'new_array(1) = ("q_gas_sm3day", "p_crit_atma", "critical flow")
            'GLV_q_gas_sm3day = new_array
            GLV_q_gas_sm3day = {Qg, p_crit_out_atma, crit}
            'GLV_q_gas_sm3day = Join(GLV_q_gas_sm3day)

            Exit Function
        Catch ex As Exception
            GLV_q_gas_sm3day = {-1}
            Dim errmsg As String
            errmsg = "error in function : GL_qgas_valve_sm3day"
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
            Dim qres(2) As Object
            Dim pd As Double
            Dim Pu As Double
            Dim Pcrit As Double
            Dim K As Double
            Dim Pd_Pu_crit As Double
            Dim crit As Boolean

            Dim prm As New CSolveParam
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
            'Dim new_array(1) As Object

            If calc_along_flow Then
                Pu = p_calc_atma
                pd = 1
                qres = GLV_q_gas_sm3day(d_mm, Pu, pd, gamma_g, t_C)
                Qmax_m3day = CDbl(qres(0))
                Pcrit = pd
                If Qmax_m3day > q_gas_sm3day And Pu > p_open_atma Then
                    Func = "calc_dq_gas_pd_valve"
                    CoeffA(2) = Pu
                    crit = False
                    Call solve_equation_bisection(Func, Pd_Pu_crit * Pu, Pu, CoeffA, prm)
                    'new_array(0) = (prm.x_solution, Qmax_m3day, Pcrit, crit)
                    'new_array(1) = ("p", "Qmax_m3day", "Pcrit", "critical flow")
                    'GLV_p_atma = new_array
                    GLV_p_atma = {prm.x_solution, Qmax_m3day, Pcrit, crit}
                    'GLV_p_atma = Join(GLV_p_atma)
                Else
                    crit = True
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
                Call solve_equation_bisection(Func, pd, Pu, CoeffA, prm)
                Dim sol As Double
                sol = prm.x_solution
                If sol < p_open_atma Then
                    sol = p_open_atma
                End If
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

    ' продублировал для использования в CGLvalve
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

    ' продублировал для использования в CGLvalve
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
    ' продублировал для использования в CGLvalve
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
        Dim p2() As Object
        Dim p1() As Object
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
                p_v_atma = p1(0) '(0)
                If p_v_atma < 0 Then
                    ' critical flow through the port achived
                    q_gas_sm3day = p1(0) '(1)
                    p_v_atma = p1(0) '(2)
                    crit1 = True
                End If

                If d_vkr_mm > 0 Then
                    p2 = GLV_p_atma(d_vkr_mm, p_v_atma, q_gas_sm3day, gamma_g, t_C, True)
                    p_atma = p2(0) '(0)
                    If p_atma < 0 Then
                        ' critical flow through the vkrutka achived
                        q_gas_sm3day = p2(0) '(1)
                        p_atma = p2(0) '(2)
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
                    p_v_atma = p1(0) '(0)
                    If p_v_atma < 0 Then
                        ' critical flow through the vkrutka achived
                        q_gas_sm3day = p1(0) '(1)
                        p_v_atma = p1(0) '(2)
                        crit2 = True
                    End If
                Else
                    p_v_atma = p_calc_atma
                End If
                p2 = GLV_p_atma(d_port_mm, p_v_atma, q_gas_sm3day, gamma_g, t_C, False)
                p_atma = p2(0) '(0)
                If p_atma < 0 Then
                    ' critical flow through the port achived
                    q_gas_sm3day = p2(0) '(1)
                    p_atma = p2(0) '(2)
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
    ' продублировал для использования в CGLvalve
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
            Dim res1() As Object
            Dim res2() As Object
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
            q_gas_sm3day = res1(0) '(0)
            crit1 = res1(0) '(2)
            crit2 = res2(0) '(2)

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
End Module

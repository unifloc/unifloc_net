'=======================================================================================
'Unifloc 7.25  coronav                                     khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
'функции для проведения расчетов из интерфейса Excel
'многофазный поток в трубах и элементах инфраструктуры
Option Explicit On

Public Module u7_Excel_functions_MF
    Public Function MF_calibr_choke_fast(
                                        ByVal qliq_sm3day As Double,
                                        ByVal fw_perc As Double,
                                        ByVal d_choke_mm As Double,
                                        Optional ByVal p_in_atma As Double = -1,
                                        Optional ByVal p_out_atma As Double = -1,
                                        Optional ByVal d_pipe_mm As Double = 70,
                                        Optional ByVal t_choke_C As Double = 20,
                                        Optional ByVal str_PVT As String = UnfClassLibrary.u7_const.PVT_DEFAULT,
                                        Optional ByVal q_gas_sm3day As Double = 0)
        Try
            Dim Choke As UnfClassLibrary.CChoke
            Choke = New UnfClassLibrary.CChoke
            'Dim pt As UnfClassLibrary.PTtype
            Dim PVT As UnfClassLibrary.CPVT
            Dim out, out_desc

            PVT = PVT_decode_string(str_PVT)
            If PVT.gas_only Then
                MF_calibr_choke_fast = "not implemented yet"
                Exit Function
            End If

            PVT.q_gas_free_sm3day = q_gas_sm3day
            Choke.fluid = PVT
            Choke.Class_Initialize()
            Choke.fluid.qliq_sm3day = qliq_sm3day
            Choke.fluid.Fw_perc = fw_perc
            Choke.d_down_m = d_pipe_mm / 1000
            Choke.d_up_m = d_pipe_mm / 1000
            Choke.d_choke_m = d_choke_mm / 1000

            If p_in_atma > p_out_atma And p_out_atma >= 1 Then
                Call Choke.calc_choke_calibration(p_in_atma, p_out_atma, t_choke_C)
                out = Choke.c_calibr_fr
                out_desc = "c_calibr_fr"
            End If

            'Dim new_array(1) As Object
            'new_array(0) = (out, p_in_atma, p_out_atma, t_choke_C, Choke.c_calibr_fr)
            'new_array(1) = (out_desc, "p_intake_atma", "p_out_atma", "t_choke_C", "c_calibr_fr")
            'MF_calibr_choke_fast = new_array
            MF_calibr_choke_fast = {out, p_in_atma, p_out_atma, t_choke_C, Choke.c_calibr_fr}
            'MF_calibr_choke_fast = Join(MF_calibr_choke_fast)
            Exit Function

        Catch ex As Exception
            MF_calibr_choke_fast = {-1}
            Dim errmsg As String
            errmsg = "Error:MF_calibr_choke_fast:"
            Throw New ApplicationException(errmsg)

        End Try
    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет давления в штуцере
    Public Function MF_p_choke_atma(
            ByVal qliq_sm3day As Double,
            ByVal fw_perc As Double,
            ByVal d_choke_mm As Double,
            Optional ByVal p_calc_from_atma As Double = -1,
            Optional ByVal calc_along_flow As Boolean = True,
            Optional ByVal d_pipe_mm As Double = 70,
            Optional ByVal t_choke_C As Double = 20,
            Optional ByVal c_calibr_fr As Double = 1,
            Optional ByVal str_PVT As String = "",
            Optional ByVal q_gas_sm3day As Double = 0)
        ' qliq_sm3day   - дебит жидкости в поверхностных условиях
        ' fw_perc       - обводненность
        ' d_choke_mm     - диаметр штуцера (эффективный)
        ' p_calc_from_atma    - давление с которого начинается расчет, атм
        '                 граничное значение для проведения расчета
        '                 либо давление на входе, либое на выходе
        ' calc_along_flow - флаг направления расчета относительно потока
        '     если = 1 то расчет по потоку
        '     ищется давление на выкиде по известному давлению на входе,
        '     ищется линейное давление по известному буферному
        '     если = 0 то расчет против потока
        '     ищется давление на входе по известному давлению на выходе,
        '     ищется буферное давление по известному линейному
        ' d_pipe_mm      - диаметр трубы до и после штуцера
        ' t_choke_C      - температура, С.
        ' c_calibr_fr   - поправочный коэффициент на штуцер
        '                 1 - отсутсвие поправки
        '                 Q_choke_real = c_calibr_fr * Q_choke_model
        ' str_PVT        - закодированная строка с параметрами PVT.
        '                 если задана - перекрывает другие значения
        ' q_gas_sm3day  - свободный газ. дополнительный к PVT потоку.
        ' результат     - число - давления на штуцере на расчетной стороне.
        '                массив значений с параметрами штуцера
        'description_end

        Try
            Dim Choke As UnfClassLibrary.CChoke
            Choke = New UnfClassLibrary.CChoke
            Dim pt As UnfClassLibrary.PTtype
            Dim pres As Double
            Dim PVT As UnfClassLibrary.CPVT
            Dim out, out_desc
            Dim p_in_atma As Double
            Dim p_out_atma As Double

            PVT = PVT_decode_string(str_PVT)
            PVT.q_gas_free_sm3day = q_gas_sm3day
            Choke.fluid = PVT
            Choke.Class_Initialize()
            Choke.fluid.qliq_sm3day = qliq_sm3day
            Choke.fluid.Fw_perc = fw_perc
            Choke.d_down_m = d_pipe_mm / 1000
            Choke.d_up_m = d_pipe_mm / 1000
            Choke.d_choke_m = d_choke_mm / 1000
            Choke.c_calibr_fr = c_calibr_fr
            If PVT.gas_only Then
                pres = UnfClassLibrary.GLV_p_atma(d_choke_mm, p_calc_from_atma, q_gas_sm3day, PVT.gamma_g, t_choke_C, calc_along_flow)(0)(0) ' проверить вывод!
                out = pres
                If calc_along_flow Then
                    p_in_atma = p_calc_from_atma
                    p_out_atma = out
                    out_desc = "Pout, atma"
                Else
                    p_out_atma = p_calc_from_atma
                    p_in_atma = out
                    out_desc = "Pin, atma"
                End If
            Else
                If calc_along_flow Then
                    p_in_atma = p_calc_from_atma
                    pt = Choke.calc_choke_p_lin(UnfClassLibrary.Set_PT(p_in_atma, t_choke_C))
                    out = pt.p_atma
                    p_out_atma = out
                    out_desc = "Pout, atma"
                Else
                    p_out_atma = p_calc_from_atma
                    pt = Choke.calc_choke_p_buf(UnfClassLibrary.Set_PT(p_out_atma, t_choke_C))
                    out = pt.p_atma
                    p_in_atma = out
                    out_desc = "Pin, atma"
                End If
            End If

            'Dim new_array(1) As Object
            'new_array(0) = (out, p_in_atma, p_out_atma, t_choke_C, Choke.c_calibr_fr, PVT.Q_gas_sm3day)
            'new_array(1) = (out_desc, "p_intake_atma", "p_out_atma", "t_choke_C", "c_calibr_fr", "q_gas_sm3day")
            'MF_p_choke_atma = new_array
            MF_p_choke_atma = {out, p_in_atma, p_out_atma, t_choke_C, Choke.c_calibr_fr, PVT.Q_gas_sm3day}
            'MF_p_choke_atma = Join(MF_p_choke_atma)
            Exit Function

        Catch ex As Exception
            MF_p_choke_atma = {-1}
            Dim errmsg As String
            errmsg = "Error:MF_p_choke_atma:"
            Throw New ApplicationException(errmsg)
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' функция расчета дебита жидкости через штуцер
    ' при заданном входном и выходном давлениях
    Public Function MF_q_choke_sm3day(
        ByVal fw_perc As Double,
        ByVal d_choke_mm As Double,
        ByVal p_in_atma As Double,
        ByVal p_out_atma As Double,
        Optional ByVal d_pipe_mm As Double = 70,
        Optional ByVal t_choke_C As Double = 20,
        Optional ByVal c_calibr_fr As Double = 1,
        Optional ByVal str_PVT As String = "",
        Optional ByVal q_gas_sm3day As Double = 0)
        ' fw_perc      - обводненность
        ' d_choke_mm   - диаметр штуцера (эффективный)
        ' p_in_atma    - давление на входе (высокой стороне)
        ' p_out_atma   - давление на выходе (низкой стороне)
        ' d_pipe_mm    - диаметр трубы до и после штуцера
        ' t_choke_C    - температура, С.
        ' c_calibr_fr  - поправочный коэффициент на штуцер
        '                 1 - отсутсвие поправки (по умолчанию)
        '                 Q_choke_real = c_calibr_fr * Q_choke_model
        ' str_PVT      - закодированная строка с параметрами PVT.
        '                 если задана - перекрывает другие значения
        ' q_gas_sm3day - дополнительный поток свободного газа
        ' результат    - расход и массив результатов
        'description_end

        Try
            Dim Choke As UnfClassLibrary.CChoke
            Choke = New UnfClassLibrary.CChoke
            Dim PVT As UnfClassLibrary.CPVT
            Dim q As Double

            PVT = PVT_decode_string(str_PVT)
            Choke.fluid = PVT
            Choke.Class_Initialize()
            Choke.d_down_m = d_pipe_mm / 1000
            Choke.d_up_m = d_pipe_mm / 1000
            Choke.d_choke_m = d_choke_mm / 1000
            Choke.fluid.Fw_perc = fw_perc
            Choke.fluid.q_gas_free_sm3day = q_gas_sm3day
            Choke.c_calibr_fr = c_calibr_fr

            'Dim new_array(1) As Object
            If PVT.gas_only Then
                q = UnfClassLibrary.GLV_q_gas_sm3day(d_choke_mm, p_in_atma, p_out_atma, PVT.gamma_g, t_choke_C, c_calibr_fr)(0)(0)
                'new_array(0) = (q, p_in_atma, p_out_atma, t_choke_C, c_calibr_fr)
                'new_array(1) = ("Qgas", "p_intake_atma", "p_out_atma", "t_choke_C", "c_calibr_fr")
            Else
                q = Choke.calc_choke_qliq_sm3day(p_in_atma, p_out_atma, t_choke_C)
                'new_array(0) = (q, p_in_atma, p_out_atma, t_choke_C, c_calibr_fr)
                'new_array(1) = ("Qliq", "p_intake_atma", "p_out_atma", "t_choke_C", "c_calibr_fr")
            End If

            'MF_p_choke_atma = new_array
            MF_q_choke_sm3day = {q, p_in_atma, p_out_atma, t_choke_C, c_calibr_fr}
            'MF_q_choke_sm3day = Join(MF_q_choke_sm3day)
            Exit Function
        Catch ex As Exception
            MF_q_choke_sm3day = {-1}
            Dim errmsg As String
            errmsg = "Error:MF_q_choke_sm3day:"
            Throw New ApplicationException(errmsg)
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет корректирующего фактора (множителя) модели штуцера под замеры
    ' медленный расчет - калибровка подбирается
    Public Function MF_calibr_choke(
            ByVal qliq_sm3day As Double,
            ByVal fw_perc As Double,
            ByVal d_choke_mm As Double,
            Optional ByVal p_in_atma As Double = -1,
            Optional ByVal p_out_atma As Double = -1,
            Optional ByVal d_pipe_mm As Double = 70,
            Optional ByVal t_choke_C As Double = 20,
            Optional ByVal str_PVT As String = "",
            Optional ByVal q_gas_sm3day As Double = 0,
            Optional ByVal calibr_type As Integer = 0)

        ' qliq_sm3day   - дебит жидкости в ст. условиях
        ' fw_perc       - обводненность
        ' d_choke_mm    - диаметр штуцера (эффективный), мм
        ' p_in_atma     - давление на входе (высокой стороне)
        ' p_out_atma    - давление на выходе (низкой стороне)
        ' d_pipe_mm     - диаметр трубы до и после штуцера, мм
        ' t_choke_C     - температура, С.
        ' str_PVT       - закодированная строка с параметрами PVT,
        '                 если задана - перекрывает другие значения
        ' q_gas_sm3day  - свободный газ. дополнительный к PVT потоку.
        ' calibr_type - тип калибровки
        '             0 - подбор параметра c_calibr
        '             1 - подбор диаметра штуцера
        '             2 - подбор газового фактор
        '             3 - подбор обводненности
        '             4 - подбор дебита жидкости
        '             5 - подбор дебита газа свободного
        ' результат     - число - калибровочный коэффициент для модели.
        '                 штуцера  - множитель на дебит через штуцер
        'description_end

        Try
            Dim Choke As UnfClassLibrary.CChoke
            Choke = New UnfClassLibrary.CChoke
            'Dim pt As UnfClassLibrary.PTtype
            Dim PVT As UnfClassLibrary.CPVT
            'Dim out, out_desc
            Dim CoeffA(0 To 2)
            Dim Func As String
            Dim cal_type_string As String
            Dim val_min As Double, val_max As Double
            Dim prm As New UnfClassLibrary.CSolveParam
            prm.Class_Initialize()


            PVT = PVT_decode_string(str_PVT)

            If PVT.gas_only Then
                MF_calibr_choke = "not implemented yet"
                Exit Function
            End If

            PVT.q_gas_free_sm3day = q_gas_sm3day
            Choke.fluid = PVT
            Choke.Class_Initialize()
            Choke.fluid.qliq_sm3day = qliq_sm3day
            Choke.fluid.Fw_perc = fw_perc
            Choke.d_down_m = d_pipe_mm / 1000
            Choke.d_up_m = d_pipe_mm / 1000
            Choke.d_choke_m = d_choke_mm / 1000
            Choke.t_choke_C = t_choke_C

            ' prepare solution function
            CoeffA(0) = Choke
            CoeffA(1) = p_in_atma
            CoeffA(2) = p_out_atma

            Select Case calibr_type
                Case 0
                    Func = "calc_choke_dp_error_calibr_grav_atm"
                    cal_type_string = "calibr"
                    val_min = 0.5
                    val_max = 1.5
                Case 1
                    Func = "calc_choke_dp_error_diam_atm"
                    cal_type_string = "diam_choke"
                    val_min = Choke.d_choke_m / 2
                    val_max = Choke.d_up_m
                Case 2
                    Func = "calc_choke_dp_error_rp_atm"
                    cal_type_string = "rp"
                    val_min = 20 'pipe.fluid.rp_m3m3 * 0.5
                    val_max = Choke.fluid.Rp_m3m3 * 2
            ' Расширить диапазон поиска по газовому фактору может быть опасно
            ' так как возможна неоднозначность решения
            ' а текущий метод поиска работает только если есть одно решение
                Case 3
                    Func = "calc_choke_dp_error_fw_atm"
                    cal_type_string = "fw"
                    val_min = 0
                    val_max = 1
                    If val_max > 1 Then val_max = 1
                Case 4
                    Func = "calc_choke_dp_error_qliq_atm"
                    cal_type_string = "qliq"
                    val_min = 0
                    val_max = Choke.fluid.qliq_sm3day * 1.5
                Case 5
                    Func = "calc_choke_dp_error_qgas_atm"
                    cal_type_string = "qgas_free"
                    val_min = 0
                    If q_gas_sm3day > 0 Then
                        val_max = q_gas_sm3day * 2
                    Else
                        val_max = 10000
                    End If
                Case Else
                    ' solve_equation_bisection without initialasing func crashes excel
                    MF_calibr_choke = "not implemented"
                    Exit Function
            End Select

            'Dim new_array(1) As Object
            prm.y_tolerance = UnfClassLibrary.const_pressure_tolerance
            If UnfClassLibrary.solve_equation_bisection(Func, val_min, val_max, CoeffA, prm) Then
                'new_array(0) = (prm.x_solution, cal_type_string, prm.y_solution, prm.iterations, prm.msg)
                MF_calibr_choke = {prm.x_solution, cal_type_string, prm.y_solution, prm.iterations, prm.msg}

            Else
                'new_array(0) = ("no solution", cal_type_string, prm.y_solution, prm.iterations, prm.msg)
                MF_calibr_choke = {"no solution", cal_type_string, prm.y_solution, prm.iterations, prm.msg}

            End If

            'new_array(1) = ("solution", "cal_type", "y_solution", "iterations", "description")
            'MF_calibr_choke = new_array
            'MF_calibr_choke = Join(MF_calibr_choke)
            Exit Function
        Catch ex As Exception
            MF_calibr_choke = {-1}
            Dim errmsg As String
            errmsg = "Error:MF_calibr_choke:"
            Throw New ApplicationException(errmsg)
        End Try

    End Function
End Module
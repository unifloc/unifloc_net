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
            Dim pt As UnfClassLibrary.u7_types.PTtype
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
            MF_calibr_choke_fast = -1
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
            Dim pt As UnfClassLibrary.u7_types.PTtype
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
            MF_p_choke_atma = -1
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
            MF_q_choke_sm3day = -1
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
            Dim pt As UnfClassLibrary.u7_types.PTtype
            Dim PVT As UnfClassLibrary.CPVT
            Dim out, out_desc
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
            If solve_equation_bisection(Func, val_min, val_max, CoeffA, prm) Then
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
            MF_calibr_choke = -1
            Dim errmsg As String
            errmsg = "Error:MF_calibr_choke:"
            Throw New ApplicationException(errmsg)
        End Try

    End Function

    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    ' функция расчета коэффициента Джоуля Томсона
    '    Public Function MF_CJT_Katm(
    '             ByVal p_atma As Double,
    '             ByVal t_C As Double,
    '    Optional ByVal str_PVT As String = PVT_DEFAULT,
    '    Optional ByVal qliq_sm3day As Double = 10,
    '    Optional ByVal fw_perc As Double = 0)
    '        ' p_atma      - давление, атм
    '        ' t_C         - температура, С.
    '        ' str_PVT     - encoded to string PVT properties of fluid
    '        ' qliq_sm3day - liquid rate (at surface)
    '        ' fw_perc     - water fraction (watercut)
    '        ' output - number
    '        'description_end


    '        On Error GoTo err1
    '        Dim PVT As New CPVT
    '    Set PVT = PVT_decode_string(str_PVT)
    '    PVT.qliq_sm3day = qliq_sm3day
    '        PVT.fw_perc = fw_perc
    '        Call PVT.calc_PVT(CDbl(p_atma), CDbl(t_C))
    '        MF_CJT_Katm = PVT.cJT_Katm
    '        Exit Function
    'err1:
    '        MF_CJT_Katm = -1
    '        addLogMsg "Error:MF_CJT_Katm:" & Err.Description
    'End Function

    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    ' расчет объемного расхода газожидкостной смеси
    '    ' для заданных термобарических условий
    '    Public Function MF_q_mix_rc_m3day(
    '             ByVal qliq_sm3day As Double,
    '             ByVal fw_perc As Double,
    '             ByVal p_atma As Double,
    '             ByVal t_C As Double,
    '    Optional ByVal str_PVT As String = PVT_DEFAULT)
    '        ' qliq_sm3day- дебит жидкости на поверхности
    '        ' fw_perc    - объемная обводненность
    '        ' p_atma     - давление, атм
    '        ' t_C        - температура, С.
    '        ' str_PVT    - закодированная строка с параметрами PVT.
    '        '              если задана - перекрывает другие значения
    '        ' результат  - число - расход ГЖС, м3/сут.
    '        'description_end

    '        On Error GoTo err1
    '        Dim PVT As CPVT
    '    Set PVT = PVT_decode_string(str_PVT)
    '    PVT.fw_perc = fw_perc
    '        PVT.qliq_sm3day = qliq_sm3day
    '        Call PVT.calc_PVT(p_atma, t_C)
    '        MF_q_mix_rc_m3day = PVT.q_mix_rc_m3day
    '        Exit Function
    'err1:
    '        MF_q_mix_rc_m3day = -1
    '        addLogMsg "Error:MF_q_mix_rc_m3day:" & Err.Description
    'End Function

    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    ' расчет плотности газожидкостной смеси для заданных  условий
    '    Public Function MF_rho_mix_kgm3(
    '             ByVal qliq_sm3day As Double,
    '             ByVal fw_perc As Double,
    '             ByVal p_atma As Double,
    '             ByVal t_C As Double,
    '    Optional ByVal str_PVT As String = PVT_DEFAULT)
    '        ' qliq_sm3day- дебит жидкости на поверхности
    '        ' fw_perc    - объемная обводненность
    '        ' p_atma     - давление, атм
    '        ' t_C        - температура, С.
    '        ' str_PVT    - закодированная строка с параметрами PVT.
    '        '              если задана - перекрывает другие значения
    '        ' результат  - число - плотность ГЖС, кг/м3.
    '        'description_end

    '        On Error GoTo err1
    '        Dim PVT As CPVT
    '    Set PVT = PVT_decode_string(str_PVT)
    '    PVT.fw_perc = fw_perc
    '        PVT.qliq_sm3day = qliq_sm3day
    '        Call PVT.calc_PVT(CDbl(p_atma), CDbl(t_C))
    '        MF_rho_mix_kgm3 = PVT.rho_mix_rc_kgm3
    '        Exit Function
    'err1:
    '        MF_rho_mix_kgm3 = -1
    '        addLogMsg "Error:MF_rho_mix_kgm3:" & Err.Description

    'End Function

    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    ' расчет вязкости газожидкостной смеси
    '    ' для заданных термобарических условий
    '    Public Function MF_mu_mix_cP(
    '            ByVal qliq_sm3day As Double,
    '            ByVal fw_perc As Double,
    '            ByVal p_atma As Double,
    '            ByVal t_C As Double,
    '   Optional ByVal str_PVT As String = PVT_DEFAULT)
    '        ' qliq_sm3day - дебит жидкости на поверхности
    '        ' fw_perc     - объемная обводненность
    '        ' p_atma      - давление, атм
    '        ' t_C         - температура, С.
    '        ' str_PVT     - закодированная строка с параметрами PVT.
    '        '              если задана - перекрывает другие значения
    '        ' результат   - число - вязкость ГЖС, м3/сут.
    '        'description_end

    '        On Error GoTo err1
    '        Dim PVT As CPVT
    '    Set PVT = PVT_decode_string(str_PVT)
    '    PVT.fw_perc = fw_perc
    '        PVT.qliq_sm3day = qliq_sm3day
    '        Call PVT.calc_PVT(CDbl(p_atma), CDbl(t_C))
    '        MF_mu_mix_cP = PVT.mu_mix_cP
    '        Exit Function
    'err1:
    '        MF_mu_mix_cP = -1
    '        addLogMsg "Error:MF_mu_mix_cP:" & Err.Description
    'End Function

    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    ' расчет доли газа в потоке
    '    Public Function MF_gas_fraction_d(
    '              ByVal p_atma As Double,
    '              ByVal t_C As Double,
    '     Optional ByVal fw_perc = 0,
    '     Optional ByVal str_PVT As String = PVT_DEFAULT,
    '     Optional ByVal ksep_add_fr As Double = 0)
    '        ' p_atma   - давление, атм
    '        ' t_C      - температура, С.
    '        ' fw_perc  - обводненность объемная
    '        ' str_PVT  - закодированная строка с параметрами PVT.
    '        '            если задана - перекрывает другие значения
    '        ' ksep_add_fr - коэффициент сепарации дополнительный
    '        '           для сепарации заданной в потоке. применяется
    '        '           для сепарации при искомом давлении
    '        ' результат - число - доля газа в потоке
    '        '              (расходная без проскальзования)
    '        'description_end
    '        On Error GoTo err1
    '        Dim PVT As CPVT
    '    Set PVT = PVT_decode_string(str_PVT)
    '    PVT.fw_perc = fw_perc
    '        Call PVT.calc_PVT(CDbl(p_atma), CDbl(t_C))
    '        MF_gas_fraction_d = PVT.gas_fraction_d(ksep_add_fr)
    '        Exit Function
    'err1:
    '        MF_gas_fraction_d = -1
    '        addLogMsg "Error:MF_gas_fraction_d:" & Err.Description
    'End Function


    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    ' расчет давления при котором
    '    ' достигается заданная доля газа в потоке
    '    Public Function MF_p_gas_fraction_atma(
    '               ByVal free_gas_d As Double,
    '               ByVal t_C As Double,
    '               ByVal fw_perc As Double,
    '      Optional ByVal str_PVT As String = PVT_DEFAULT,
    '      Optional ByVal ksep_add_fr As Double = 0)
    '        ' free_gas_d - допустимая доля газа в потоке;
    '        ' t_C        - температура, С;
    '        ' fw_perc    - объемная обводненность, проценты %;
    '        ' str_PVT    - закодированная строка с параметрами PVT.
    '        '              Если задана - перекрывает другие значения.
    '        ' ksep_add_fr - коэффициент сепарации дополнительный
    '        '           для сепарации заданной в потоке. применяется
    '        '           для сепарации при искомом давлении
    '        ' результат  - число - давление, атма.
    '        'description_end
    '        On Error GoTo err1
    '        Dim PVT As CPVT
    '    Set PVT = PVT_decode_string(str_PVT)
    '    PVT.fw_perc = fw_perc
    '        MF_p_gas_fraction_atma = PVT.p_gas_fraction_atma(free_gas_d, t_C, ksep_add_fr)
    '        Exit Function
    'err1:
    '        MF_p_gas_fraction_atma = -1
    '        addLogMsg "Error:MF_p_gas_fraction_atma:" & Err.Description
    'End Function

    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    ' расчет газового фактора
    '    ' при котором достигается заданная доля газа в потоке
    '    Public Function MF_rp_gas_fraction_m3m3(
    '                ByVal free_gas_d As Double,
    '                ByVal p_atma As Double,
    '                ByVal t_C As Double,
    '                ByVal fw_perc As Double,
    '       Optional ByVal str_PVT As String = PVT_DEFAULT,
    '       Optional ByVal Rp_limit_m3m3 As Double = 500,
    '       Optional ByVal ksep_add_fr As Double = 0)
    '        ' free_gas_d - допустимая доля газа в потоке
    '        ' p_atma     - давление, атм
    '        ' t_C        - температура, С.
    '        ' fw_perc    - объемная обводненность, проценты %;
    '        ' str_PVT    - закодированная строка с параметрами PVT.
    '        '              если задана - перекрывает другие значения
    '        ' Rp_limit_m3m3 - верхняя граница оценки ГФ
    '        ' ksep_add_fr - коэффициент сепарации дополнительный
    '        '           для сепарации заданной в потоке. применяется
    '        '           для сепарации при искомом давлении
    '        ' результат  - число - газовый фактор, м3/м3.
    '        'description_end
    '        On Error GoTo err1
    '        Dim PVT As CPVT
    '    Set PVT = PVT_decode_string(str_PVT)
    '    PVT.fw_perc = fw_perc
    '        MF_rp_gas_fraction_m3m3 = PVT.rp_gas_fraction_m3m3(free_gas_d, p_atma, t_C, ksep_add_fr, Rp_limit_m3m3)
    '        Exit Function
    'err1:
    '        MF_rp_gas_fraction_m3m3 = -1
    '        addLogMsg "Error:MF_rp_gas_fraction_m3m3:" & Err.Description
    'End Function

    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    ' расчет натуральной сепарации газа на приеме насоса
    '    Public Function MF_ksep_natural_d(
    '             ByVal qliq_sm3day As Double,
    '             ByVal fw_perc As Double,
    '             ByVal p_intake_atma As Double,
    '    Optional ByVal t_intake_C As Double = 50,
    '    Optional ByVal d_intake_mm As Double = 90,
    '    Optional ByVal d_cas_mm As Double = 120,
    '    Optional ByVal str_PVT As String = PVT_DEFAULT)
    '        ' qliq_sm3day   - дебит жидкости в поверхностных условиях
    '        ' fw_perc       - обводненность
    '        ' p_intake_atma - давление сепарации
    '        ' t_intake_C    - температура сепарации
    '        ' d_intake_mm   - диаметр приемной сетки
    '        ' d_cas_mm      - диаметр эксплуатационной колонны
    '        ' str_PVT       - закодированная строка с параметрами PVT.
    '        '                 если задана - перекрывает другие значения
    '        ' результат     - число - естественная сепарация
    '        'description_end

    '        On Error GoTo err1
    '        Dim fluid As New CPVT
    '    Set fluid = PVT_decode_string(str_PVT)
    '    fluid.qliq_sm3day = qliq_sm3day
    '        fluid.fw_perc = fw_perc
    '        Call fluid.calc_PVT(p_intake_atma, t_intake_C)
    '        With fluid
    '            MF_ksep_natural_d = unf_natural_separation(d_intake_mm / 1000, d_cas_mm / 1000, .qliq_sm3day, .q_gas_sm3day, .bo_m3m3, .bg_m3m3,
    '                                                    .sigma_oil_gas_Nm, .sigma_wat_gas_Nm, .rho_oil_sckgm3, .rho_gas_sckgm3, .fw_perc)
    '        End With
    '        Exit Function
    'err1:
    '        MF_ksep_natural_d = -1
    '        addLogMsg "Error:MF_ksep_natural_d:" & Err.Description
    'End Function


    ''description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '' расчет общей сепарации на приеме насоса
    'Public Function MF_ksep_total_d(
    '    ByVal SepNat As Double,
    '    ByVal SepGasSep As Double)
    '    ' SepNat        - естественная сепарация
    '    ' SepGasSep     - искусственная сепарация (газосепаратор)
    '    MF_ksep_total_d = SepNat + (1 - SepNat) * SepGasSep
    'End Function


    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    'расчет градиента давления
    '    'с использованием многофазных корреляций
    '    Public Function MF_dpdl_atmm(ByVal d_m As Double,
    '             ByVal p_atma As Double,
    '             ByVal Ql_rc_m3day As Double,
    '             ByVal Qg_rc_m3day As Double,
    '    Optional ByVal mu_oil_cP As Double = const_mu_o,
    '    Optional ByVal mu_gas_cP As Double = const_mu_g,
    '    Optional ByVal sigma_oil_gas_Nm As Double = const_sigma_oil_Nm,
    '    Optional ByVal rho_lrc_kgm3 As Double = const_go_ * 1000,
    '    Optional ByVal rho_grc_kgm3 As Double = const_gg_ * const_rho_air,
    '    Optional ByVal eps_m As Double = 0.0001,
    '    Optional ByVal theta_deg As Double = 90,
    '    Optional ByVal hcorr As Integer = 1,
    '    Optional ByVal param_out As Integer = 0,
    '    Optional ByVal c_calibr_grav As Double = 1,
    '    Optional ByVal c_calibr_fric As Double = 1)
    '        ' расчет градиента давления по одной из корреляций
    '        ' d_m - диаметр трубы в которой идет поток
    '        ' p_atma - давление в точке расчета
    '        ' Ql_rc_m3day - дебит жидкости в рабочих условиях
    '        ' Qg_rc_m3day - дебит газа в рабочих условиях
    '        ' mu_oil_cP - вязкость нефти в рабочих условиях
    '        ' mu_gas_cP - вязкость газа в рабочих условиях
    '        ' sigma_oil_gas_Nm - поверхностное натяжение
    '        '              жидкость газ
    '        ' rho_lrc_kgm3 - плотность нефти
    '        ' rho_grc_kgm3 - плотность газа
    '        ' eps_m     - шероховатость
    '        ' theta_deg - угол от горизонтали
    '        ' hcorr  - тип корреляции
    '        ' param_out - параметр для вывода
    '        ' c_calibr_grav - калибровка гравитации
    '        ' c_calibr_fric - калибровка трения
    '        'description_end

    '        Dim PrGrad

    '        On Error GoTo er1
    '        Select Case hcorr
    '            Case 0

    '                PrGrad = unf_BegsBrillGradient(d_m, theta_deg, eps_m,
    '                                Ql_rc_m3day, Qg_rc_m3day,
    '                                mu_oil_cP, mu_gas_cP,
    '                                sigma_oil_gas_Nm,
    '                                rho_lrc_kgm3,
    '                                rho_grc_kgm3, , , c_calibr_grav, c_calibr_fric)
    '            Case 1

    '                PrGrad = unf_AnsariGradient(d_m, theta_deg, eps_m,
    '                                Ql_rc_m3day, Qg_rc_m3day,
    '                                mu_oil_cP, mu_gas_cP,
    '                                sigma_oil_gas_Nm,
    '                                rho_lrc_kgm3,
    '                                rho_grc_kgm3,
    '                                p_atma, c_calibr_grav, c_calibr_fric)
    '            Case 2

    '                PrGrad = unf_UnifiedTUFFPGradient(d_m, theta_deg, eps_m,
    '                                Ql_rc_m3day, Qg_rc_m3day,
    '                                mu_oil_cP, mu_gas_cP,
    '                                sigma_oil_gas_Nm,
    '                                rho_lrc_kgm3,
    '                                rho_grc_kgm3,
    '                                p_atma, c_calibr_grav, c_calibr_fric)
    '            Case 3

    '                PrGrad = unf_GrayModifiedGradient(d_m, theta_deg, eps_m,
    '                                Ql_rc_m3day, Qg_rc_m3day,
    '                                mu_oil_cP, mu_gas_cP,
    '                                sigma_oil_gas_Nm,
    '                                rho_lrc_kgm3,
    '                                rho_grc_kgm3,
    '                                , , , c_calibr_grav, c_calibr_fric)
    '            Case 4

    '                PrGrad = unf_HagedornandBrawnmodified(d_m, theta_deg, eps_m,
    '                                Ql_rc_m3day, Qg_rc_m3day,
    '                                mu_oil_cP, mu_gas_cP,
    '                                sigma_oil_gas_Nm,
    '                                rho_lrc_kgm3,
    '                                rho_grc_kgm3,
    '                                p_atma, , , , c_calibr_grav, c_calibr_fric)

    '        End Select

    '        If param_out = 0 Then
    '            MF_dpdl_atmm = PrGrad
    '        Else
    '            MF_dpdl_atmm = PrGrad(param_out)
    '        End If
    '        Exit Function
    'er1:
    '        MF_dpdl_atmm = -1
    '        addLogMsg "Error:MF_dpdl_atmm:" & Err.Description
    'End Function



    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    '  подбор параметров потока через трубу при известном
    '    '  перепаде давления с использованием многофазных корреляций
    '    Public Function MF_calibr_pipeline(
    '                 ByVal p_calc_from_atma As Double,
    '                 ByVal p_calc_to_atma As Double,
    '             ByVal t_calc_from_C As Double,
    '             ByVal t_val,
    '             ByVal h_list_m As Variant,
    '             ByVal diam_list_mm As Variant,
    '             ByVal qliq_sm3day As Double,
    '             ByVal fw_perc As Double,
    '    Optional ByVal q_gas_sm3day As Double = 0,
    '    Optional ByVal str_PVT As String = PVT_DEFAULT,
    '    Optional ByVal calc_flow_direction As Integer = 11,
    '    Optional ByVal hydr_corr As H_CORRELATION = 0,
    '    Optional ByVal temp_method As TEMP_CALC_METHOD = StartEndTemp,
    '    Optional ByVal c_calibr = 1,
    '    Optional ByVal roughness_m As Double = 0.0001,
    '    Optional ByVal out_curves As Integer = 1,
    '    Optional ByVal out_curves_num_points As Integer = 20,
    '    Optional ByVal calibr_type As Integer = 0)

    '        'p_calc_from_atma - давление начальное, атм
    '        '                   граничное значение для проведения расчета
    '        'p_calc_to_atma   - давление конечное, атм
    '        '                   граничное значение для проведения расчета
    '        ' t_calc_from_C - температура в точке где задано давление расчета
    '        ' t_val     - температура вдоль трубопровода
    '        '           если число то температура на другом конце трубы
    '        '           если range или таблица [0..N,0..1] то температура
    '        '           окружающей среды по вертикальной глубине, С
    '        ' h_list_m  - траектория трубопровода, если число то измеренная
    '        '           длина, range или таблица [0..N,0..1] то траектория
    '        ' diam_list_mm  - внутрнний диаметр трубы, если число то задается
    '        '           постоянный диаметр, если range или таблица [0..N,0..1]
    '        '           то задается зависимость диаметра от измеренной длины
    '        ' qliq_sm3day - дебит жидкости в поверхностных условиях, нм3/сут
    '        '           если qliq_sm3day =0 и q_gas_sm3day > 0
    '        '           тогда считается барботаж газа через жидкость
    '        ' fw_perc   - обводненность объемная в стандартных условиях
    '        ' q_gas_sm3day  - свободный газ нм3/сут. дополнительный к PVT потоку.
    '        '           учитывается для барботажа или режима потока газа
    '        '           в других случаях добавляется к общему потоку меняя rp
    '        ' str_PVT   - закодированная строка с параметрами PVT.
    '        '           если задана - перекрывает другие значения
    '        '           если задан флаг gas_only = 1 то жидкость не учитывается
    '        ' calc_flow_direction - направление расчета и потока относительно
    '        '           координат.  11 расчет и поток по координате
    '        '                       10 расчет по коордиате, поток против
    '        '                       00 расчет и поток против координате
    '        '                       01 расчет против координат, поток по
    '        ' hydr_corr   - гидравлическая корреляция, H_CORRELATION
    '        '           BeggsBrill = 0,
    '        '           Ansari = 1,
    '        '           Unified = 2,
    '        '           Gray = 3,
    '        '           HagedornBrown = 4,
    '        '           SakharovMokhov = 5
    '        ' temp_method  - метод расчета температуры
    '        '           0 - линейное распределение по длине
    '        '           1 - температура равна температуре окружающей среды
    '        '           2 - расчет температуры с учетом эмиссии в окр. среду
    '        ' c_calibr  - поправка на гравитационную составляющую
    '        '           перепада давления, если дать ссылку на две ячейки,
    '        '           то вторая будет поправка на трение.
    '        ' roughness_m - шероховатость трубы
    '        ' out_curves - флаг вывод значений между концами трубы
    '        '           1 основные, 2 все значения.
    '        '           вывод может замедлять расчет (не сильно)
    '        ' out_curves_num_points - количество точек для вывода значений
    '        '           между концами трубы.
    '        ' calibr_type - тип калибровки
    '        '          0 - подбор параметра c_calibr_grav
    '        '          1 - подбор параметра c_calibr_fric
    '        '          2 - подбор газового фактор
    '        '          3 - подбор обводненности
    '        ' результат   - массив с подобранным парамером и подробностями.
    '        'description_end

    '        Dim pipe As New CPipe
    '        Dim res(), res1()
    '        Dim CoeffA(0 To 2)
    '        Dim Func As String
    '        Dim cal_type_string As String
    '        Dim val_min As Double, val_max As Double
    '        Dim out, out_desc
    '        Dim prm As New CSolveParam

    '        On Error GoTo err1 : 

    '    Set pipe = new_pipeline_with_stream(qliq_sm3day, _
    '                                        fw_perc, _
    '                                        h_list_m, _
    '                                        t_calc_from_C, _
    '                                        calc_flow_direction, _
    '                                        str_PVT, _
    '                                        diam_list_mm, _
    '                                        hydr_corr, _
    '                                        t_val, _
    '                                        temp_method, _
    '                                        c_calibr, _
    '                                        roughness_m, _
    '                                        q_gas_sm3day)

    '    ' prepare solution function
    '    Set CoeffA(0) = pipe
    '        CoeffA(1) = p_calc_from_atma
    '        CoeffA(2) = p_calc_to_atma

    '        Select Case calibr_type
    '            Case 0
    '                Func = "calc_pipe_dp_error_calibr_grav_atm"
    '                cal_type_string = "calibr_grav"
    '                val_min = 0.5
    '                val_max = 1.5
    '            Case 1
    '                Func = "calc_pipe_dp_error_calibr_fric_atm"
    '                cal_type_string = "calibr_fric"
    '                val_min = 0.1
    '                val_max = 1
    '            Case 2
    '                Func = "calc_pipe_dp_error_rp_atm"
    '                cal_type_string = "rp"
    '                val_min = 20 'pipe.fluid.rp_m3m3 * 0.5
    '                val_max = pipe.fluid.rp_m3m3 * 2
    '            ' Расширить диапазон поиска по газовому фактору может быть опасно
    '            ' так как возможна неоднозначность решения
    '            ' а текущий метод поиска работает только если есть одно решение
    '            Case 3
    '                Func = "calc_pipe_dp_error_fw_atm"
    '                cal_type_string = "fw"
    '                val_min = 0
    '                val_max = 1
    '                If val_max > 1 Then val_max = 1
    '            Case 4
    '                Func = "calc_pipe_dp_error_qliq_atm"
    '                cal_type_string = "qliq"
    '                val_min = 0
    '                val_max = pipe.fluid.qliq_sm3day * 1.5
    '            Case 5
    '                Func = "calc_pipe_dp_error_qgas_atm"
    '                cal_type_string = "qgas"
    '                val_min = 0
    '                If q_gas_sm3day > 0 Then
    '                    val_max = q_gas_sm3day * 2
    '                Else
    '                    val_max = 10000
    '                End If
    '            Case Else
    '                ' solve_equation_bisection without initialasing func crashes excel
    '                MF_calibr_pipeline = "not implemented"
    '                Exit Function
    '        End Select

    '        prm.y_tolerance = const_pressure_tolerance
    '        If solve_equation_bisection(Func, val_min, val_max, CoeffA, prm) Then

    '            out = Array(prm.x_solution,
    '                    cal_type_string,
    '                    prm.y_solution,
    '                    prm.iterations,
    '                    prm.msg)

    '        Else
    '            out = Array("no solution",
    '                    cal_type_string,
    '                    prm.y_solution,
    '                    prm.iterations,
    '                    prm.msg)
    '        End If

    '        out_desc = Array("solution",
    '                     "cal_type",
    '                     "y_solution",
    '                     "iterations",
    '                     "description")

    '        MF_calibr_pipeline = array_join(Array(out, out_desc))

    '        Exit Function
    'err1:
    '        MF_calibr_pipeline = Array(-1, "error")
    '        addLogMsg "Error:MF_calibr_pipeline:" & Err.Description

    'End Function



    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    '  подбор параметров потока через трубу при известном
    '    '  перепаде давления с использованием многофазных корреляций
    '    ' (лучше не использовать - используйте MF_calibr_pipeline)
    '    Public Function MF_calibr_pipe(
    '        ByVal p_calc_from_atma As Double,
    '        ByVal p_calc_to_atma As Double,
    '        ByVal t_calc_from_C As Double,
    '        ByVal t_calc_to_C As Double,
    '        ByVal length_m As Double,
    '        ByVal theta_deg As Double,
    '        ByVal d_mm As Double,
    '        ByVal qliq_sm3day As Double,
    '        ByVal fw_perc As Double,
    '        Optional ByVal q_gas_sm3day As Double = 0,
    '        Optional ByVal str_PVT As String = PVT_DEFAULT,
    '        Optional ByVal calc_flow_direction As Integer = 11,
    '        Optional ByVal hydr_corr As H_CORRELATION = 0,
    '        Optional ByVal c_calibr = 1,
    '        Optional ByVal roughness_m As Double = 0.0001,
    '        Optional ByVal calibr_type As Integer = 0)
    '        'p_calc_from_atma - давление с которого начинается расчет, атм
    '        '                   граничное значение для проведения расчета
    '        ' t_calc_from_C   - температура в точке где задано давление, С
    '        ' t_calc_to_C     - температура на другом конце трубы
    '        '                 по умолчанию температура вдоль трубы постоянна
    '        '                   если задано то меняется линейно по трубе
    '        ' length_m        - Длина трубы, измеренная, м
    '        ' theta_deg       - угол направления потока к горизонтали
    '        ' d_mm            - внутренний диаметр трубы
    '        ' qliq_sm3day     - дебит жидкости в поверхностных условиях
    '        '              если qliq_sm3day =0 и q_gas_sm3day > 0
    '        '              тогда считается барботаж газа через жидкость
    '        ' fw_perc         - обводненность
    '        ' q_gas_sm3day    - свободный газ. дополнительный к PVT потоку.
    '        ' str_PVT         - закодированная строка с параметрами PVT.
    '        '                   если задана - перекрывает другие значения
    '        '        если задан флаг gas_only = 1 то жидкость не учитывается
    '        ' calc_flow_direction - направление расчета и потока
    '        '                   относительно координат
    '        '                   если = 11 расчет и поток по координате
    '        '                   если = 10 расчет по, поток против координат
    '        '                   если = 00 расчет и поток против координате
    '        '                   если = 01 расчет против, поток по координате
    '        ' hydr_corr       - гидравлическая корреляция, H_CORRELATION
    '        '                    BeggsBrill = 0
    '        '                    Ansari = 1
    '        '                    Unified = 2
    '        '                    Gray = 3
    '        '                    HagedornBrown = 4
    '        '                    SakharovMokhov = 5
    '        ' c_calibr        - поправка на гравитационную составляющую
    '        '           перепада давления, если дать ссылку на две ячейки,
    '        '           то вторая будет поправка на трение
    '        ' roughness_m     - шероховатость трубы
    '        ' out_curves      - флаг вывод значений между концами трубы
    '        '                   0 минимум, 1 основные, 2 все значения.
    '        '                   вывод может замедлять расчет (не сильно)
    '        ' out_curves_num_points - количество точек для вывода значений
    '        '                   между концами трубы.
    '        ' результат       - число - давление на другом конце трубы atma.
    '        '                  или массив - первая строка значения
    '        '                               вторая строка - подписи
    '        'description_end


    '        Dim pipe As New CPipe
    '        Dim prm As New CSolveParam
    '        Dim CoeffA(0 To 2)
    '        Dim Func As String
    '        Dim cal_type_string As String
    '        Dim val_min As Double, val_max As Double
    '        Dim out, out_desc

    '        On Error GoTo err1
    '        ' check pipe length
    '        length_m = Abs(length_m)                ' length must be positive
    '        If length_m = 0 Then
    '            MF_calibr_pipe = Array(p_calc_from_atma, t_calc_from_C)
    '            Exit Function
    '        End If
    '    ' initialize pipe
    '    Set pipe = new_pipe_with_stream(qliq_sm3day, fw_perc, length_m, calc_flow_direction, _
    '                                    str_PVT, theta_deg, d_mm, hydr_corr, _
    '                                    t_calc_from_C, t_calc_to_C, _
    '                                    c_calibr, _
    '                                    roughness_m, q_gas_sm3day)

    '    ' prepare solution function
    '    Set CoeffA(0) = pipe
    '        CoeffA(1) = p_calc_from_atma
    '        CoeffA(2) = p_calc_to_atma

    '        Select Case calibr_type
    '            Case 0
    '                Func = "calc_pipe_dp_error_calibr_grav_atm"
    '                cal_type_string = "calibr_grav"
    '                val_min = 0.5
    '                val_max = 1.5
    '            Case 1
    '                Func = "calc_pipe_dp_error_calibr_fric_atm"
    '                cal_type_string = "calibr_fric"
    '                val_min = 0.5
    '                val_max = 1.5
    '            Case 2
    '                Func = "calc_pipe_dp_error_rp_atm"
    '                cal_type_string = "rp"
    '                val_min = 20 'pipe.fluid.rp_m3m3 * 0.5
    '                val_max = pipe.fluid.rp_m3m3 * 2
    '            ' Расширить диапазон поиска по газовому фактору может быть опасно
    '            ' так как возможна неоднозначность решения
    '            ' а текущий метод поиска работает только если есть одно решение
    '            Case 3
    '                Func = "calc_pipe_dp_error_fw_atm"
    '                cal_type_string = "fw"
    '                val_min = 0
    '                val_max = 1
    '                If val_max > 1 Then val_max = 1
    '            Case 4
    '                Func = "calc_pipe_dp_error_qliq_atm"
    '                cal_type_string = "qliq"
    '                val_min = 1
    '                val_max = pipe.fluid.qliq_sm3day * 1.5
    '            Case 5
    '                Func = "calc_pipe_dp_error_qgas_atm"
    '                cal_type_string = "qgas"
    '                val_min = 0
    '                If q_gas_sm3day > 0 Then
    '                    val_max = q_gas_sm3day * 2
    '                Else
    '                    val_max = 10000
    '                End If
    '            Case Else
    '                ' solve_equation_bisection without initialasing func crashes excel
    '                MF_calibr_pipe = "not implemented"
    '                Exit Function
    '        End Select

    '        prm.y_tolerance = const_pressure_tolerance
    '        If solve_equation_bisection(Func, val_min, val_max, CoeffA, prm) Then

    '            out = Array(prm.x_solution,
    '                    cal_type_string,
    '                    prm.y_solution,
    '                    prm.iterations,
    '                    prm.msg)

    '        Else
    '            out = Array("no solution",
    '                    cal_type_string,
    '                    CStr(prm.y_solution),
    '                    CStr(prm.iterations),
    '                    prm.msg)
    '        End If

    '        out_desc = Array("solution",
    '                     "cal_type",
    '                     "y_solution",
    '                     "iterations",
    '                     "description")
    '        MF_calibr_pipe = array_join(Array(out, out_desc))

    '        Exit Function
    'err1:
    '        MF_calibr_pipe = Array(-1, "error")
    '        addLogMsg "Error:MF_calibr_pipe:" & Err.Description


    'End Function



    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    '  расчет распределения давления и температуры в трубопроводе
    '    '  с использованием многофазных корреляций
    '    Public Function MF_p_pipeline_atma(
    '                 ByVal p_calc_from_atma As Double,
    '                 ByVal t_calc_from_C As Double,
    '                 ByVal t_val_C As Variant,
    '                 ByVal h_list_m As Variant,
    '                 ByVal diam_list_mm As Variant,
    '                 ByVal qliq_sm3day As Double,
    '                 ByVal fw_perc As Double,
    '        Optional ByVal q_gas_sm3day As Double = 0,
    '        Optional ByVal str_PVT As String = PVT_DEFAULT,
    '        Optional ByVal calc_flow_direction As Integer = 11,
    '        Optional ByVal hydr_corr As H_CORRELATION = 0,
    '        Optional ByVal temp_method As TEMP_CALC_METHOD = StartEndTemp,
    '        Optional ByVal c_calibr = 1,
    '        Optional ByVal roughness_m As Double = 0.0001,
    '        Optional ByVal out_curves As Integer = 1,
    '        Optional ByVal out_curves_num_points As Integer = 20,
    '        Optional ByVal num_value As Integer = 0,
    '        Optional ByVal znlf As Boolean = False)
    '        ' p_calc_from_atma  - давление с которого начинается расчет, атм
    '        '           граничное значение для проведения расчета
    '        ' t_calc_from_C - температура в точке где задано давление расчета
    '        ' t_val_C   - температура вдоль трубопровода
    '        '           если число то температура на другом конце трубы
    '        '           если range или таблица [0..N,0..1] то температура
    '        '           окружающей среды по вертикальной глубине, С
    '        ' h_list_m  - траектория трубопровода, если число то измеренная
    '        '           длина, range или таблица [0..N,0..1] то траектория
    '        ' diam_list_mm  - внутрнний диаметр трубы, если число то задается
    '        '           постоянный диаметр, если range или таблица [0..N,0..1]
    '        '           то задается зависимость диаметра от измеренной длины
    '        ' qliq_sm3day - дебит жидкости в поверхностных условиях, нм3/сут
    '        '           если qliq_sm3day =0 и q_gas_sm3day > 0
    '        '           тогда считается барботаж газа через жидкость
    '        ' fw_perc   - обводненность объемная в стандартных условиях
    '        ' q_gas_sm3day  - свободный газ нм3/сут. дополнительный к PVT потоку.
    '        '           учитывается для барботажа или режима потока газа
    '        '           в других случаях добавляется к общему потоку меняя rp
    '        ' str_PVT   - закодированная строка с параметрами PVT.
    '        '           если задана - перекрывает другие значения
    '        '           если задан флаг gas_only = 1 то жидкость не учитывается
    '        ' calc_flow_direction - направление расчета и потока относительно
    '        '           координат.  11 расчет и поток по координате
    '        '                       10 расчет по коордиате, поток против
    '        '                       00 расчет и поток против координате
    '        '                       01 расчет против координат, поток по
    '        ' hydr_corr - гидравлическая корреляция, H_CORRELATION
    '        '           BeggsBrill = 0,
    '        '           Ansari = 1,
    '        '           Unified = 2,
    '        '           Gray = 3,
    '        '           HagedornBrown = 4,
    '        '           SakharovMokhov = 5
    '        ' temp_method  - метод расчета температуры
    '        '           0 - линейное распределение по длине
    '        '           1 - температура равна температуре окружающей среды
    '        '           2 - расчет температуры с учетом эмиссии в окр. среду
    '        ' c_calibr  - поправка на гравитационную составляющую
    '        '           перепада давления, если дать ссылку на две ячейки,
    '        '           то вторая будет поправка на трение.
    '        ' roughness_m - шероховатость трубы
    '        ' out_curves - флаг вывод значений между концами трубы
    '        '           1 основные, 2 все значения.
    '        '           вывод может замедлять расчет (не сильно)
    '        ' out_curves_num_points - количество точек для вывода значений
    '        '           между концами трубы.
    '        ' num_value       - значение которое будет выводиться первым
    '        ' znlf    - флаг для расчета вертикального барботажа (дин уровень)
    '        ' результат - число - давление на другом конце трубы atma.
    '        '           и распределение параметров по трубе
    '        'description_end

    '        Dim pipe As New CPipe
    '        Dim res, res1

    '        On Error GoTo err1 : 

    '    Set pipe = new_pipeline_with_stream(qliq_sm3day, _
    '                                        fw_perc, _
    '                                        h_list_m, _
    '                                        t_calc_from_C, _
    '                                        calc_flow_direction, _
    '                                        str_PVT, _
    '                                        diam_list_mm, _
    '                                        hydr_corr, _
    '                                        t_val_C, _
    '                                        temp_method, _
    '                                        c_calibr, _
    '                                        roughness_m, _
    '                                        q_gas_sm3day, _
    '                                        znlf)

    '    ' calc pressure distribution
    '    res1 = PT_to_array(pipe.calc_dPipe(p_calc_from_atma, t_calc_from_C, allCurves))
    '        If out_curves = 2 Then
    '            res = pipe.array_out(out_curves_num_points, all_curves_out:=True)
    '        Else
    '            res = pipe.array_out(out_curves_num_points)
    '        End If
    '        res(0, 0) = res1(0)

    '        res(0, 0) = res(0, num_value)
    '        res(1, 0) = res(1, num_value)
    '        MF_p_pipeline_atma = res

    '        Exit Function
    'err1:
    '        MF_p_pipeline_atma = Array(-1, "error")
    '        addLogMsg "Error:MF_p_pipeline_atma:" & Err.Description
    'End Function


    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    ' расчет распределения давления и температуры в трубе
    '    ' (лучше не использовать - используйте MF_p_pipeline_atma)
    '    Public Function MF_p_pipe_atma(
    '        ByVal p_calc_from_atma As Double,
    '        ByVal t_calc_from_C As Double,
    '        ByVal t_calc_to_C As Double,
    '        ByVal length_m As Double,
    '        ByVal theta_deg As Double,
    '        ByVal d_mm As Double,
    '        ByVal qliq_sm3day As Double,
    '        ByVal fw_perc As Double,
    '        Optional ByVal q_gas_sm3day As Double = 0,
    '        Optional ByVal str_PVT As String = PVT_DEFAULT,
    '        Optional ByVal calc_flow_direction As Integer = 11,
    '        Optional ByVal hydr_corr As H_CORRELATION = 0,
    '        Optional ByVal c_calibr = 1,
    '        Optional ByVal roughness_m As Double = 0.0001,
    '        Optional ByVal out_curves As Integer = 1,
    '        Optional ByVal out_curves_num_points As Integer = 20,
    '        Optional ByVal num_value As Integer = 0)
    '        'p_calc_from_atma - давление с которого начинается расчет, атм
    '        '                   граничное значение для проведения расчета
    '        ' t_calc_from_C   - температура в точке где задано давление, С
    '        ' t_calc_to_C     - температура на другом конце трубы
    '        '                 по умолчанию температура вдоль трубы постоянна
    '        '                   если задано то меняется линейно по трубе
    '        ' length_m        - Длина трубы, измеренная, м
    '        ' theta_deg       - угол направления потока к горизонтали
    '        ' d_mm            - внутренний диаметр трубы
    '        ' qliq_sm3day     - дебит жидкости в поверхностных условиях
    '        '              если qliq_sm3day =0 и q_gas_sm3day > 0
    '        '              тогда считается барботаж газа через жидкость
    '        ' fw_perc         - обводненность
    '        ' q_gas_sm3day    - свободный газ. дополнительный к PVT потоку.
    '        ' str_PVT         - закодированная строка с параметрами PVT.
    '        '                   если задана - перекрывает другие значения
    '        '        если задан флаг gas_only = 1 то жидкость не учитывается
    '        ' calc_flow_direction - направление расчета и потока
    '        '                   относительно координат
    '        '                   если = 11 расчет и поток по координате
    '        '                   если = 10 расчет по, поток против координат
    '        '                   если = 00 расчет и поток против координате
    '        '                   если = 01 расчет против, поток по координате
    '        ' hydr_corr       - гидравлическая корреляция, H_CORRELATION
    '        '                    BeggsBrill = 0
    '        '                    Ansari = 1
    '        '                    Unified = 2
    '        '                    Gray = 3
    '        '                    HagedornBrown = 4
    '        '                    SakharovMokhov = 5
    '        ' c_calibr        - поправка на гравитационную составляющую
    '        '           перепада давления, если дать ссылку на две ячейки,
    '        '           то вторая будет поправка на трение
    '        ' roughness_m     - шероховатость трубы
    '        ' out_curves      - флаг вывод значений между концами трубы
    '        '                   0 минимум, 1 основные, 2 все значения.
    '        '                   вывод может замедлять расчет (не сильно)
    '        ' out_curves_num_points - количество точек для вывода значений
    '        '                   между концами трубы.
    '        ' num_value       - значение которое будет выводиться первым
    '        ' результат       - число - давление на другом конце трубы atma.
    '        '                  или массив - первая строка значения
    '        '                               вторая строка - подписи
    '        'description_end

    '        Dim pipe As CPipe
    '        Dim PVT As New CPVT
    '        Dim PTcalc As PTtype
    '        Dim PTin As PTtype
    '        Dim PTout As PTtype
    '        Dim TM As TEMP_CALC_METHOD
    '        Dim out, out_desc
    '        Dim out_curves_type As CALC_RESULTS
    '        Dim res
    '        Dim arr

    '        On Error GoTo err1

    '        ' check pipe length
    '        length_m = Abs(length_m)                ' length must be positive
    '        If length_m = 0 Then
    '            MF_p_pipe_atma = Array(p_calc_from_atma, t_calc_from_C)
    '            Exit Function
    '        End If
    '    ' initialize pipe
    '    Set pipe = new_pipe_with_stream(qliq_sm3day, fw_perc, length_m, calc_flow_direction, _
    '                                    str_PVT, theta_deg, d_mm, hydr_corr, _
    '                                    t_calc_from_C, t_calc_to_C, _
    '                                    c_calibr, _
    '                                    roughness_m, q_gas_sm3day)
    '    ' prep output
    '    If out_curves Then
    '            out_curves_type = allCurves
    '        Else
    '            out_curves_type = nocurves
    '        End If
    '        ' calc pressure distribution
    '        PTcalc = pipe.calc_dPipe(p_calc_from_atma, , out_curves_type)  ' ( -allcurves * out_curves ) can be used for simpcity but not

    '        ' prep results for output
    '        If calc_flow_direction \ 10 = 1 Then
    '            PTout = PTcalc
    '            PTin.p_atma = p_calc_from_atma
    '            PTin.t_C = t_calc_from_C
    '        Else
    '            PTin = PTcalc
    '            PTout.p_atma = p_calc_from_atma
    '            PTout.t_C = t_calc_from_C
    '        End If

    '        ' out results based on out_curves value
    '        If out_curves = 1 Then
    '            res = pipe.array_out(out_curves_num_points)
    '            res(0, 0) = PTcalc.p_atma
    '            res(0, 1) = PTcalc.t_C
    '            arr = res
    '        ElseIf out_curves = 2 Then
    '            res = pipe.array_out(out_curves_num_points, all_curves_out:=True)
    '            res(0, 0) = PTcalc.p_atma
    '            res(0, 1) = PTcalc.t_C
    '            arr = res
    '        Else
    '            out = Array(PTcalc.p_atma,
    '                    PTcalc.t_C,
    '                    PTin.p_atma,
    '                    PTin.t_C,
    '                    PTout.p_atma,
    '                    PTout.t_C,
    '                    pipe.c_calibr_grav,
    '                    pipe.c_calibr_fric)

    '            out_desc = Array("p_calc_atma",
    '                         "t_calc_C",
    '                         "p_in_atma",
    '                         "t_in_C",
    '                         "p_out_atma",
    '                         "t_out_C",
    '                         "c_calibr_grav",
    '                         "c_calibr_fric")

    '            arr = array_join(Array(out, out_desc))

    '        End If
    '        arr(0, 0) = arr(0, num_value)
    '        arr(1, 0) = arr(1, num_value)
    '        MF_p_pipe_atma = arr
    '        Exit Function
    'err1:
    '        MF_p_pipe_atma = Array(-1, "error")
    '        addLogMsg "Error:MF_p_pipe_atma:" & Err.Description
    'End Function



    ' ==============  функции для расчета штуцера ==========================
    ' =====================================================================


End Module
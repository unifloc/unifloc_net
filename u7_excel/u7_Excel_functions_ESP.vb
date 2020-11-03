﻿'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
' Модуль для расчетов по режимам работы УЭЦН Excel
Option Explicit On
Public Module u7_Excel_functions_ESP

    Public Function ESP_head_m(
        ByVal qliq_m3day As Double,
        Optional ByVal num_stages As Integer = 1,
        Optional ByVal freq_Hz As Double = 50,
        Optional ByVal pump_id As Integer = 0,
        Optional ByVal mu_cSt As Double = -1,
        Optional ByVal c_calibr As Double = 1) As Double
        ' qliq_m3day - дебит жидкости в условиях насоса (стенд)
        ' num_stages  - количество ступеней
        ' freq_Hz    - частота вращения насоса
        ' pump_id    - номер насоса в базе данных
        ' mu_cSt     - вязкость жидкости, сСт;
        ' c_calibr    - коэффициент поправки на напор..
        '               если массив то второе значение - поправка на подачу (множитель)
        '               третье на мощность (множитель)
        'description_end

        Try
            Dim esp As New UnfClassLibrary.CESPpump
            esp.Class_Initialize()
            Call esp.set_ID(pump_id)
            If esp Is Nothing Then
                ESP_head_m = 0
                Exit Function
            End If

            'Dim c_calibr_head As Double
            'Dim c_calibr_rate As Double
            'Dim c_calibr_power As Double
            ' Call read_ESP_calibr(c_calibr, c_calibr_head, c_calibr_rate, c_calibr_power)  не работает функция калибровки, ссылаемся на clbr, который пуст

            esp.freq_Hz = freq_Hz
            esp.stage_num = num_stages
            qliq_m3day = qliq_m3day / esp.c_calibr_rate
            ESP_head_m = esp.get_ESP_head_m(qliq_m3day, num_stages, mu_cSt)
            ESP_head_m = ESP_head_m * esp.c_calibr_head
            Exit Function
        Catch ex As Exception
            ESP_head_m = -1
            Dim msg As String
            msg = "ESP_head_m: error with " & "ESP_head_m = " & CStr(ESP_head_m) & ": " & ex.Message

            Throw New ApplicationException(msg)
        End Try

    End Function


    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' номинальная мощность потребляемая ЭЦН с вала (на основе каталога ЭЦН)
    ' учитывается поправка на вязкость
    Public Function ESP_power_W(
            ByVal qliq_m3day As Double,
            Optional ByVal num_stages As Integer = 1,
            Optional ByVal freq_Hz As Double = 50,
            Optional ByVal pump_id As Integer = 737,
            Optional ByVal mu_cSt As Double = -1,
            Optional ByVal c_calibr As Double = 1) As Double
        ' мощность УЭЦН номинальная потребляемая
        ' qliq_m3day - дебит жидкости в условиях насоса (стенд)
        ' num_stages  - количество ступеней
        ' freq_Hz       - частота вращения насоса
        ' pump_id     - номер насоса в базе данных
        ' mu_cSt     - вязкость жидкости
        ' c_calibr    - коэффициент поправки на напор.
        '               если массив то второе значение - поправыка на подачу (множитель)
        '               третье на мощность (множитель)
        'description_end

        Try
            Dim esp As New UnfClassLibrary.CESPpump
            esp.Class_Initialize()
            Call esp.set_ID(pump_id)
            If esp Is Nothing Then
                ESP_power_W = 0
                Exit Function
            End If

            Dim c_calibr_head As Double
            Dim c_calibr_rate As Double
            Dim c_calibr_power As Double
            Call read_ESP_calibr(c_calibr, c_calibr_head, c_calibr_rate, c_calibr_power)


            esp.freq_Hz = freq_Hz
            esp.stage_num = num_stages
            qliq_m3day = qliq_m3day / c_calibr_rate
            ESP_power_W = esp.get_ESP_power_W(qliq_m3day, num_stages, mu_cSt)
            ESP_power_W = ESP_power_W * c_calibr_power
            Exit Function
        Catch ex As Exception
            Dim msg As String
            ESP_power_W = -1
            msg = "Error:ESP_power_W:"

            Throw New ApplicationException(msg)
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' номинальный КПД ЭЦН (на основе каталога ЭЦН)
    ' учитывается поправка на вязкость
    Public Function ESP_eff_fr(
            ByVal qliq_m3day As Double,
            Optional ByVal num_stages As Integer = 1,
            Optional ByVal freq_Hz As Double = 50,
            Optional ByVal pump_id As Integer = 737,
            Optional ByVal mu_cSt As Double = -1,
            Optional ByVal c_calibr As Double = 1) As Double
        ' qliq_m3day - дебит жидкости в условиях насоса (стенд)
        ' num_stages  - количество ступеней
        ' freq_Hz       - частота вращения насоса
        ' pump_id     - номер насоса в базе данных
        ' mu_cSt     - вязкость жидкости
        ' c_calibr    - коэффициент поправки на напор.
        '               если массив то второе значение - поправыка на подачу (множитель)
        '               третье на мощность (множитель)
        'description_end

        Try
            Dim esp As New UnfClassLibrary.CESPpump
            esp.Class_Initialize()
            Call esp.set_ID(pump_id)
            If esp Is Nothing Then
                ESP_eff_fr = 0
                Exit Function
            End If

            Dim c_calibr_head As Double
            Dim c_calibr_rate As Double
            Dim c_calibr_power As Double
            Call read_ESP_calibr(c_calibr, c_calibr_head, c_calibr_rate, c_calibr_power)

            esp.freq_Hz = freq_Hz
            esp.stage_num = num_stages
            qliq_m3day = qliq_m3day / c_calibr_rate
            esp.correct_visc_let = True
            ESP_eff_fr = esp.get_ESP_effeciency_fr(qliq_m3day, mu_cSt)
            ESP_eff_fr = ESP_eff_fr * c_calibr_head * c_calibr_rate / c_calibr_power
            Exit Function
        Catch ex As Exception
            Dim msg As String
            msg = "Error:ESP_eff_fr:"

            Throw New ApplicationException(msg)
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' название ЭЦН по номеру
    Public Function ESP_name(ByVal pump_id As Integer) As String
        ' pump_id    - идентификатор насоса в базе данных
        ' результат - название насоса
        'description_end

        Try
            Dim esp As New UnfClassLibrary.CESPpump
            esp.Class_Initialize()
            Call esp.set_ID(pump_id)
            If esp Is Nothing Then
                ESP_name = "no name"
                Exit Function
            End If

            ESP_name = esp.db.name

            Exit Function
        Catch ex As Exception
            Dim msg As String
            ESP_name = -1
            msg = "Error:ESP_name:"

            Throw New ApplicationException(msg)
        End Try

    End Function

    ' description_to_manual - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' максимальный дебит ЭЦН для заданной частоты
    ' по номинальной кривой РНХ
    Public Function ESP_rate_max_sm3day(
        Optional ByVal freq_Hz As Double = 50,
        Optional ByVal pump_id As Integer = 737,
        Optional ByVal mu_cSt As Double = -1) As Double
        ' freq_Hz   - частота вращения ЭЦН
        ' pump_id    - идентификатор насоса в базе данных
        'description_end
        Try
            Dim esp As New UnfClassLibrary.CESPpump
            esp.Class_Initialize()
            Call esp.set_ID(pump_id)
            If esp Is Nothing Then
                ESP_rate_max_sm3day = 0
                Exit Function
            End If
            esp.freq_Hz = freq_Hz
            ESP_rate_max_sm3day = esp.rate_max_sm3day(mu_cSt)
            Exit Function
        Catch ex As Exception
            Dim msg As String
            msg = "Error:ESP_rate_max_sm3day:"

            Throw New ApplicationException(msg)
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' оптимальный дебит ЭЦН для заданной частоты
    ' по номинальной кривой РНХ
    Public Function ESP_optRate_m3day(
        Optional ByVal freq_Hz As Double = 50,
        Optional ByVal pump_id As Integer = 737) As Double
        ' freq_Hz   - частота вращения ЭЦН
        ' pump_id    - идентификатор насоса в базе данных
        'description_end

        Try
            Dim esp As New UnfClassLibrary.CESPpump
            esp.Class_Initialize()
            Call esp.set_ID(pump_id)
            If esp Is Nothing Then
                ESP_optRate_m3day = 0
                Exit Function
            End If
            esp.freq_Hz = freq_Hz
            ESP_optRate_m3day = esp.rate_nom_sm3day
            Exit Function
        Catch ex As Exception
            Dim msg As String
            msg = "Error:ESP_optRate_m3day:"

            Throw New ApplicationException(msg)
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' функция возвращает идентификатор типового насоса по значению
    ' номинального дебита
    Public Function ESP_id_by_rate(q As Double)
        ' возвращает ID в зависимости от номинального дебита.
        ' насосы подобраны вручную из текущей базы.
        ' Q - номинальный дебит
        'description_end

        If q > 0 And q < 20 Then ESP_id_by_rate = 738 :         ' ЭЦН5-15
        If q >= 20 And q < 40 Then ESP_id_by_rate = 740 :         ' ЭЦН5-30
        If q >= 40 And q < 60 Then ESP_id_by_rate = 1005 :         ' ЭЦН5-50
        If q >= 60 And q < 100 Then ESP_id_by_rate = 1006 :         ' ЭЦН5-80
        If q >= 100 And q < 150 Then ESP_id_by_rate = 737 :         ' ЭЦН5-125
        If q >= 150 And q < 250 Then ESP_id_by_rate = 748 :         ' ЭЦН5A-200
        If q >= 250 And q < 350 Then ESP_id_by_rate = 750 :         ' ЭЦН5A-320Э
        If q >= 350 And q < 600 Then ESP_id_by_rate = 753 :         ' ЭЦН5А-500
        If q >= 600 And q < 800 Then ESP_id_by_rate = 754 :         ' ЭЦН5А-700
        If q >= 800 And q < 1200 Then ESP_id_by_rate = 755 :         ' ЭЦН6-1000
        If q > 1200 Then ESP_id_by_rate = 758
    End Function
    'description_end

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    'функция расчета давления на выходе/входе ЭЦН в рабочих условиях
    Public Function ESP_p_atma(
                     ByVal qliq_sm3day As Double,
                     ByVal fw_perc As Double,
                     ByVal p_calc_atma As Double,
            Optional ByVal num_stages As Integer = 1,
            Optional ByVal freq_Hz As Double = 50,
            Optional ByVal pump_id As Integer = 737,
            Optional ByVal str_PVT As String = UnfClassLibrary.PVT_DEFAULT,
            Optional ByVal t_intake_C As Double = 50,
            Optional ByVal t_dis_C As Double = 50,
            Optional ByVal calc_along_flow As Boolean = 1,
            Optional ByVal ESP_gas_correct As Double = 1,
            Optional ByVal c_calibr As Double = 1,
            Optional ByVal dnum_stages_integrate As Integer = 1,
            Optional ByVal out_curves_num_points As Integer = 20,
            Optional ByVal num_value As Integer = 0,
            Optional ByVal q_gas_sm3day As Double = 0)
        ' qliq_sm3day       - дебит жидкости на поверхности
        ' fw_perc           - обводненность
        ' p_calc_atma        - давление для которого делается расчет
        '                     либо давление на приеме насоса
        '                     либо давление на выкиде насоса
        '                     определяется параметром calc_along_flow
        ' num_stages        - количество ступеней
        ' freq_Hz           - частота вращения вала ЭЦН, Гц
        ' pump_id           - идентификатор насоса
        ' str_PVT            - набор данных PVT
        ' t_intake_C        - температура на приеме насоа
        ' t_dis_C            - температура на выкиде насоса.
        '                     если = 0 и calc_along_flow = 1 то рассчитывается
        ' calc_along_flow    - режим расчета снизу вверх или сверху вниз
        '                 calc_along_flow = True => p_atma давление на приеме
        '                 calc_along_flow = False => p_atma давление на выкиде
        ' ESP_gas_correct  - деградация по газу:
        '      0 - 2 задает значение вручную;
        '      10 стандартный ЭЦН (предел 25%);
        '      20 ЭЦН с газостабилизирующим модулем (предел 50%);
        '      30 ЭЦН с осевым модулем (предел 75%);
        '      40 ЭЦН с модифицированным ступенями (предел 40%).
        '      110+, тогда модель n-100 применяется ко всем ступеням отдельно
        '         Предел по доле газа на входе в насос после сепарации
        '         на основе статьи SPE 117414 (с корректировкой)
        '         поправка дополнительная к деградации (суммируется).
        ' c_calibr  - коэффициент поправки на напор.
        '       если массив то второе значение - поправыка на подачу (множитель)
        '       третье на мощность (множитель)
        ' dnum_stages_integrate - шаг интегрирования ЭЦН
        '           если >1 будет быстрее но менее точно
        ' out_curves_num_points - количество точек для вывода значений
        '                   по ступеням
        ' num_value       - значение которое будет выводиться первым
        ' q_gas_sm3day    - свободный газ в потоке
        ' результат   - массив значений включающий
        'description_end
        Dim arr(,) As Object
        'Dim clbr
        Dim esp As New UnfClassLibrary.CESPpump
        esp.Class_Initialize()
        Dim c_calibr_head As Double
        Dim c_calibr_rate As Double
        Dim c_calibr_power As Double

        Try
            ' get ESP from database
            Call esp.set_ID(pump_id)

            If esp Is Nothing Then
                ESP_p_atma = "no ESP"
                Exit Function
            End If

            With esp
                If str_PVT <> "" Then
                    .fluid = PVT_decode_string(str_PVT)
                End If

                Call read_ESP_calibr(c_calibr, c_calibr_head, c_calibr_rate, c_calibr_power)

                .c_calibr_head = c_calibr_head
                .c_calibr_rate = c_calibr_rate
                .c_calibr_power = c_calibr_power

                .fluid.qliq_sm3day = qliq_sm3day
                .fluid.Fw_perc = fw_perc
                .fluid.q_gas_free_sm3day = q_gas_sm3day

                .freq_Hz = freq_Hz
                .stage_num = num_stages
                .gas_correct = ESP_gas_correct
                .dnum_stages_integrate = dnum_stages_integrate

                Call .calc_ESP(p_calc_atma, t_intake_C, t_dis_C, calc_along_flow, saveCurve:=True)

                'arr = .array_out(out_curves_num_points)
                If calc_along_flow Then
                    arr(0, 0) = .p_dis_atma
                    arr(1, 0) = "p_dis_atma"
                Else
                    arr(0, 0) = .p_int_atma
                    arr(1, 0) = "p_intake_atma"
                End If
            End With

            arr(0, 0) = arr(0, num_value)
            arr(1, 0) = arr(1, num_value)

            ESP_p_atma = arr
            Exit Function
        Catch ex As Exception
            Dim msg As String
            ESP_p_atma = "error"
            msg = "Error:ESP_dp_atm:"

            Throw New ApplicationException(msg)
        End Try

    End Function


    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    ' расчет подстроечных параметров системы УЭЦН
    '    Public Function ESP_calibr_pump(
    '                     ByVal qliq_sm3day As Double,
    '                     ByVal fw_perc As Double,
    '                     ByVal p_int_atma As Double,
    '                     ByVal p_dis_atma As Double,
    '            Optional ByVal num_stages As Integer = 1,
    '            Optional ByVal freq_Hz As Double = 50,
    '            Optional ByVal pump_id As Long = 674,
    '            Optional ByVal str_PVT As String = PVT_DEFAULT,
    '            Optional ByVal t_intake_C As Double = 50,
    '            Optional ByVal t_dis_C As Double = 50,
    '            Optional ByVal calc_along_flow As Boolean = 1,
    '            Optional ByVal ESP_gas_correct As Double = 1,
    '            Optional ByVal c_calibr = 1,
    '            Optional ByVal dnum_stages_integrate As Integer = 1,
    '            Optional ByVal calibr_type As Integer = 0)
    '        ' qliq_sm3day       - дебит жидкости на поверхности
    '        ' fw_perc           - обводненность
    '        ' p_int_atma        - давление на приеме насоса
    '        ' p_dis_atma        - давление на выкиде насоса
    '        ' num_stages        - количество ступеней
    '        ' freq_Hz           - частота вращения вала ЭЦН, Гц
    '        ' pump_id           - идентификатор насоса
    '        ' str_PVT            - набор данных PVT
    '        ' t_intake_C        - температура на приеме насоа
    '        ' t_dis_C            - температура на выкиде насоса.
    '        '                     если = 0 и calc_along_flow = 1 то рассчитывается
    '        ' calc_along_flow    - режим расчета снизу вверх или сверху вниз
    '        '                 calc_along_flow = True => p_atma давление на приеме
    '        '                 calc_along_flow = False => p_atma давление на выкиде
    '        ' ESP_gas_correct  - деградация по газу:
    '        '     0 - 2 задает значение вручную;
    '        '     10 стандартный ЭЦН (предел 25%);
    '        '     20 ЭЦН с газостабилизирующим модулем (предел 50%);
    '        '     30 ЭЦН с осевым модулем (предел 75%);
    '        '     40 ЭЦН с модифицированным ступенями (предел 40%).
    '        '     110+, тогда модель n-100 применяется ко всем ступеням отдельно
    '        '     Предел по доле газа на входе в насос после сепарации
    '        '     на основе статьи SPE 117414 (с корректировкой)
    '        '     поправка дополнительная к деградации (суммируется).
    '        ' c_calibr  - коэффициент поправки на напор.
    '        '     если массив то второе значение - поправыка на подачу (множитель)
    '        '     третье на мощность (множитель)
    '        ' dnum_stages_integrate - шаг интегрирования ЭЦН
    '        '           если >1 будет быстрее но менее точно
    '        ' calibr_type - тип калибровки
    '        ' результат   - массив значений включающий
    '        'description_end


    '        Dim esp As New CESPpump
    '        Dim c_calibr_head As Double
    '        Dim c_calibr_rate As Double
    '        Dim c_calibr_power As Double

    '        Dim prm As New CSolveParam

    '        Dim CoeffA(0 To 4)
    '        Dim Func As String
    '        Dim cal_type_string As String
    '        Dim val_min As Double, val_max As Double
    '        Dim out, out_desc

    '        On Error GoTo er1
    '        ' get ESP from database
    '        Call esp.set_ID(pump_id)

    '        If esp Is Nothing Then
    '            ESP_calibr_pump = "no ESP"
    '            Exit Function
    '        End If

    '        With esp
    '            If str_PVT <> "" Then
    '             Set .fluid = PVT_decode_string(str_PVT)
    '        End If

    '            Call read_ESP_calibr(c_calibr, c_calibr_head, c_calibr_rate, c_calibr_power)

    '            .c_calibr_head = c_calibr_head
    '            .c_calibr_rate = c_calibr_rate
    '            .c_calibr_power = c_calibr_power

    '            .fluid.qliq_sm3day = qliq_sm3day
    '            .fluid.Fw_perc = fw_perc

    '            .freq_Hz = freq_Hz
    '            .stage_num = num_stages
    '            .gas_correct = ESP_gas_correct

    '            .dnum_stages_integrate = dnum_stages_integrate
    '        End With
    '        ' prepare solution function
    '    Set CoeffA(0) = esp
    '    CoeffA(1) = p_int_atma
    '        CoeffA(2) = p_dis_atma
    '        CoeffA(3) = t_intake_C
    '        CoeffA(4) = t_dis_C

    '        Select Case calibr_type
    '            Case 0
    '                Func = "calc_ESP_dp_error_calibr_head_atm"
    '                cal_type_string = "calibr_head"
    '                val_min = 0.5
    '                val_max = 1.5
    '            Case 1
    '                Func = "calc_ESP_dp_error_calibr_rate_atm"
    '                cal_type_string = "calibr_rate"
    '                val_min = 0.5
    '                val_max = 1.5
    '        End Select

    '        prm.y_tolerance = const_pressure_tolerance

    '        If solve_equation_bisection(Func, val_min, val_max, CoeffA, prm) Then

    '            out = Array(prm.x_solution,
    '                        cal_type_string,
    '                        prm.y_solution,
    '                        prm.iterations,
    '                        prm.msg)

    '        Else
    '            out = Array("no solution",
    '                        cal_type_string,
    '                        CStr(prm.y_solution),
    '                        CStr(prm.iterations),
    '                        prm.msg)
    '        End If

    '        out_desc = Array("solution",
    '                         "cal_type",
    '                         "y_solution",
    '                         "iterations",
    '                         "description")
    '        ESP_calibr_pump = array_join(Array(out, out_desc))

    '        Exit Function
    'er1:
    '        ESP_calibr_pump = "error"
    '        AddLogMsg "Error:ESP_calibr_pump:" & Err.Description

    'End Function


    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    ' расчет производительности системы УЭЦН
    '    ' считает перепад давления, электрические параметры и деградацию КПД
    '    Public Function ESP_system_calc(
    '                     ByVal qliq_sm3day As Double,
    '                     ByVal fw_perc As Double,
    '                     ByVal qgas_free_sm3day As Double,
    '                     ByVal p_calc_atma As Double,
    '                     ByVal t_intake_C As Double,
    '            Optional ByVal t_dis_C As Double = -1,
    '            Optional ByVal str_PVT As String,
    '            Optional ByVal str_ESP As String,
    '            Optional ByVal str_motor As String,
    '            Optional ByVal str_cable As String,
    '            Optional ByVal str_gassep As String,
    '            Optional ByVal calc_along_flow As Boolean = 1,
    '            Optional ByVal out_curves_num_points As Integer = 20,
    '            Optional ByVal num_value As Integer = 0)
    '        ' qliq_sm3day       - дебит жидкости на поверхности
    '        ' fw_perc           - обводненность
    '        ' qgas_free_sm3day  - свободный газ в потоке
    '        ' p_calc_atma       - давление для которого делается расчет
    '        '                     либо давление на приеме насоса
    '        '                     либо давление на выкиде насоса
    '        '                     определяется параметром calc_along_flow
    '        ' str_PVT            - набор данных PVT
    '        ' str_ESP            - набор данных ЭЦН

    '        ' calc_along_flow    - режим расчета снизу вверх или сверху вниз
    '        '            calc_along_flow = True => p_atma давление на приеме
    '        '           calc_along_flow = False => p_atma давление на выкиде
    '        ' out_curves_num_points - количество точек для вывода значений
    '        '            по ступеня.
    '        ' num_value       - значение которое будет выводиться первым
    '        ' результат   - массив значений включающий
    '        '            перепад давления
    '        '            перепад температур
    '        '            мощность потребляемая с вала, Вт
    '        '            мощность гидравлическая по перекачке жидкости, Вт
    '        '            КПД ЭЦН
    '        '            список неполон
    '        'description_end
    '        Dim arr
    '        Dim i As Integer
    '        Dim nrows As Integer
    '        Dim fr_Hz As Double
    '        On Error GoTo er1
    '        Dim ESPsys As New CESPsystem

    '        Dim fluid As CPVT

    '    Set fluid = PVT_decode_string(str_PVT)
    '    fluid.qliq_sm3day = qliq_sm3day
    '        fluid.Fw_perc = fw_perc
    '        fluid.q_gas_free_sm3day = qgas_free_sm3day


    '        Call ESPsys.init_json(str_ESP, str_motor, str_cable, str_gassep, fluid)

    '        Call ESPsys.calc_ESPsys(p_calc_atma, t_intake_C, t_dis_C, calc_along_flow, saveCurve:=True)

    '        arr = ESPsys.array_out(out_curves_num_points)


    '        arr(0, 0) = arr(0, num_value)
    '        arr(1, 0) = arr(1, num_value)

    '        ESP_system_calc = arr

    '        Exit Function
    'er1:
    '        ESP_system_calc = -1
    '        AddLogMsg "Error:ESP_system_calc:" & Err.Description

    'End Function



    '    '=======================================================
    '    '--------------------- Двигатель -----------------------
    '    '=======================================================



    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    ' функция расчета параметров двигателя по заданному моменту на валу
    '    Public Function ESP_motor_calc_mom(ByVal mom_Nm As Double,
    '                              Optional ByVal freq_Hz As Double = 50,
    '                              Optional ByVal U_V As Double = -1,
    '                              Optional ByVal U_nom_V As Double = 500,
    '                              Optional ByVal P_nom_kW As Double = 10,
    '                              Optional ByVal f_nom_Hz As Double = 50,
    '                              Optional ByVal motorID As Integer = 0,
    '                              Optional ByVal eff_nom_fr As Double = 0.85,
    '                              Optional ByVal cosphi_nom_fr As Double = 0.8,
    '                              Optional ByVal slip_nom_fr As Double = 0.05,
    '                              Optional ByVal d_od_mm As Double = 117,
    '                              Optional ByVal lambda As Double = 2,
    '                              Optional ByVal alpha0 As Double = 0.4,
    '                              Optional ByVal xi0 As Double = 1.05,
    '                              Optional ByVal Ixcf As Double = 0.4) _
    '                      As Variant
    '        ' mom_Nm      - момент развиваемый двигателем на валу, Нм
    '        ' freq_Hz     - частота вращения внешнего поля
    '        ' U_V         - напряжение рабочее, линейное, В
    '        ' U_nom_V     - номинальное напряжение питания двигателя, линейное, В
    '        ' P_nom_kW    - номинальная мощность двигателя кВт
    '        ' f_nom_Hz    - номинальная частота вращения поля, Гц
    '        ' motorID     - тип 0 - постоянные значения,
    '        '                   1 - задается по каталожным кривым, ассинхронный
    '        '                   2 - задается по схеме замещения, ассинхронный
    '        ' eff_nom_fr  - КПД при номинальном режиме работы
    '        ' cosphi_nom_fr - коэффициент мощности при номинальном режиме работы
    '        ' slip_nom_fr - скольжение при номинальном режиме работы
    '        ' d_od_mm     - внешний диаметр - габарит ПЭД
    '        ' lambda      - для motorID = 2 перегрузочный коэффициент
    '        '               отношение макс момента к номинальному
    '        ' alpha0  - параметр. влияет на положение макс КПД.для motorID = 2
    '        ' xi0     - параметр. определяет потери момента при холостом ходе.
    '        '           для motorID = 2
    '        ' Ixcf    - поправка на поправку тока холостого хода
    '        '           при изменении напряжения и частоты от минимальной.
    '        '           для motorID = 2' результат   - момент на валу двигателя
    '        'description_end
    '        On Error GoTo er1

    '        Dim arr, arr_name

    '        Dim motor As New CESPMotor

    '        Call motor.InitMotor(motorID, U_nom_V, P_nom_kW, f_nom_Hz, eff_nom_fr, cosphi_nom_fr, slip_nom_fr, d_od_mm, lambda, alpha0, xi0)
    '        Call motor.calc_motor_mom_Nm(mom_Nm, freq_Hz, U_V)
    '        With motor
    '            arr = Array(.I_lin_A, .CosPhi_d, .eff_d, .s_d, .M_Nm, .Pshaft_kW)
    '            arr_name = Array("I_lin_A", "CosPhi_d", "eff_d", "slip", "M_Nm", "Pshaft_kW")
    '        End With
    '        ESP_motor_calc_mom = array_join(Array(arr, arr_name))
    '        Exit Function
    'er1:
    '        ESP_motor_calc_mom = -1
    '        AddLogMsg "Error:motor_calc_mom:" & Err.Description

    'End Function



    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    ' расчет полной характеристики двигателя от проскальзования
    '    ' по заданной величине скольжения (на основе схемы замещения)
    '    Public Function ESP_motor_calc_slip(ByVal S As Double,
    '                               Optional ByVal freq_Hz As Double = 50,
    '                               Optional ByVal U_V As Double = -1,
    '                               Optional ByVal U_nom_V As Double = 500,
    '                               Optional ByVal P_nom_kW As Double = 10,
    '                               Optional ByVal f_nom_Hz As Double = 50,
    '                               Optional ByVal eff_nom_fr As Double = 0.85,
    '                               Optional ByVal cosphi_nom_fr As Double = 0.8,
    '                               Optional ByVal slip_nom_fr As Double = 0.05,
    '                               Optional ByVal d_od_mm As Double = 117,
    '                               Optional ByVal lambda As Double = 2,
    '                               Optional ByVal alpha0 As Double = 0.4,
    '                               Optional ByVal xi0 As Double = 1.05,
    '                               Optional ByVal Ixcf As Double = 0.4)
    '        ' s           - скольжение двигателя
    '        ' freq_Hz     - частота вращения внешнего поля
    '        ' U_V         - напряжение рабочее, линейное, В
    '        ' U_nom_V     - номинальное напряжение питания двигателя, линейное, В
    '        ' P_nom_kW    - номинальная мощность двигателя кВт
    '        ' f_nom_Hz    - номинальная частота вращения поля, Гц
    '        ' eff_nom_fr  - КПД при номинальном режиме работы
    '        ' cosphi_nom_fr - коэффициент мощности при номинальном режиме работы
    '        ' slip_nom_fr - скольжение при номинальном режиме работы
    '        ' d_od_mm     - внешний диаметр - габарит ПЭД
    '        ' lambda      - для motorID = 2 перегрузочный коэффициент
    '        '               отношение макс момента к номинальному
    '        ' alpha0  - параметр. влияет на положение макс КПД.для motorID = 2
    '        ' xi0     - параметр. определяет потери момента при холостом ходе.
    '        '           для motorID = 2
    '        ' Ixcf    - поправка на поправку тока холостого хода
    '        '           при изменении напряжения и частоты от минимальной.
    '        '           для motorID = 2
    '        ' результат   - массив параметров ПЭД
    '        'description_end


    '        On Error GoTo er1

    '        Dim arr, arr_name
    '        Dim motor As New CESPMotor
    '        Dim sk
    '        With motor
    '            Call motor.InitMotor(2, U_nom_V, P_nom_kW, f_nom_Hz, eff_nom_fr, cosphi_nom_fr, slip_nom_fr, d_od_mm, lambda, alpha0, xi0)
    '            sk = .calc_s_M_krit(U_V, freq_Hz)
    '            motor.calc_motor_slip S, freq_Hz, U_V

    '        arr = Array(.I_lin_A, .CosPhi_d, .eff_d, .s_d, .M_Nm, .Pshaft_kW, sk(1, 1), sk(1, 2))
    '            arr_name = Array("I_lin_A", "CosPhi_d", "eff_d", "slip", "M_Nm", "Pshaft_kW", sk(2, 1), sk(2, 2))
    '        End With

    '        ESP_motor_calc_slip = array_join(Array(arr, arr_name))

    '        Exit Function
    'er1:
    '        ESP_motor_calc_slip = -1
    '        AddLogMsg "Error:ESP_motor_calc_slip:" & Err.Description
    'End Function

    '    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '    ' функция выдает номинальные параметры ПЭД
    '    Public Function ESP_motor_nameplate(
    '                                   Optional ByVal Unom_V As Double = 500,
    '                                   Optional ByVal Pnom_kW As Double = 10,
    '                                   Optional ByVal Fnom_Hz As Double = 50,
    '                                   Optional ByVal motorID As Integer = 0,
    '                                   Optional ByVal eff_fr As Double = 0.85,
    '                                   Optional ByVal cosphi_fr As Double = 0.8,
    '                                   Optional ByVal slip_fr As Double = 0.05,
    '                                   Optional ByVal d_od_mm As Double = 117,
    '                                   Optional num As Integer = 1)
    '        ' опциональные параметры
    '        ' Unom_V      - номинальное напряжение питания двигателя, линейное, В
    '        ' Pnom_kW     - номинальная мощность двигателя кВт
    '        ' fnom_Hz     - номинальная частота вращения поля, Гц
    '        ' motorID     - тип 0 - постоянные значения,
    '        '                   1 - задается по каталожным кривым, ассинхронный
    '        '                   2 - задается по схеме замещения, ассинхронный
    '        ' eff_fr      - КПД для типа 0
    '        ' cosphi_fr   - коэффициент мощности для типа 0
    '        ' slip_fr     - проскальзывание для типа 0
    '        ' d_od_mm     - внешний диаметр ПЭД
    '        ' num   - номер который выводится первым
    '        '   результат   - формальное название ПЭД
    '        'description_end
    '        On Error GoTo er1


    '        Dim motor As New CESPMotor
    '        Dim arr, arr_name
    '        Dim sk
    '        Call motor.InitMotor(motorID, Unom_V, Pnom_kW, Fnom_Hz, eff_fr, cosphi_fr, slip_fr)
    '        'sk = motor.calc_s_M_krit(Unom_V, Fnom_Hz)
    '        With motor
    '            arr = Array(.name,
    '                        .name,
    '                        .manufacturer_name,
    '                   CStr(.Pnom_kW),
    '                   CStr(.Unom_lin_V),
    '                   CStr(.Inom_lin_A),
    '                   CStr(.Snom_d),
    '                   CStr(.CosPhinom_d),
    '                   CStr(.Fnom_Hz),
    '                   CStr(.Mnom_Nm),
    '                   CStr(.length_m),
    '                   CStr(.d_od_mm))
    '            arr(0) = arr(num)

    '            arr_name = Array("name",
    '                             "name",
    '                             "manufacturer_name",
    '                             "Pnom_kW",
    '                             "Unom_lin_V",
    '                             "Inom_lin_A",
    '                             "Snom_d",
    '                             "CosPhinom_d",
    '                             "Fnom_Hz",
    '                             "Mnom_Nm",
    '                             "length_m",
    '                             "d_od_mm")
    '            arr_name(0) = arr_name(num)

    '        End With
    '        ESP_motor_nameplate = array_join(Array(arr, arr_name))
    '        Exit Function
    'er1:
    '        ESP_motor_nameplate = -1
    '        AddLogMsg "Error:ESP_motor_nameplate:" & Err.Description
    'End Function





    '=======================================================
    '--------------------- Газосепаратор -------------------
    '=======================================================





    ''description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '' расчет коэффициента сепарации газосепаратора
    '' по результатам стендовых испытаний РГУ нефти и газа
    'Public Function ESP_gassep_ksep_d(
    '                ByVal gsep_type_TYPE As Integer,
    '                ByVal gas_frac_d As Double,
    '                ByVal qliq_sm3day As Double,
    '       Optional ByVal freq_Hz As Double = 50) As Double
    '    ' MY_SEPFACTOR - Вычисление коэффициента сепрации в точке
    '    '   gsep_type_TYPE    - тип сепаратора (номер от 1 до 29)
    '    '    1  - 'GDNK5'
    '    '    2  - 'VGSA (VORTEX)'
    '    '    3  - 'GDNK5A'
    '    '    4  - 'GSA5-1'
    '    '    5  - 'GSA5-3'
    '    '    6  - 'GSA5-4'
    '    '    7  - 'GSAN-5A'
    '    '    8  - 'GSD-5A'
    '    '    9  - 'GSD5'
    '    '    10 - '3MNGB5'
    '    '    11 - '3MNGB5A'
    '    '    12 - '3MNGDB5'
    '    '    13 - '3MNGDB5A'
    '    '    14 - 'MNGSL5A-M'
    '    '    15 - 'MNGSL5A-TM'
    '    '    16 - 'MNGSL5-M'
    '    '    17 - 'MNGSL5-TM'
    '    '    18 - 'MNGSLM 5'
    '    '    19 - 'MNGD 5'
    '    '    20 - 'GSIK 5A'
    '    '    21 - '338DSR'
    '    '    22 - '400GSR'
    '    '    23 - '400GSV'
    '    '    24 - '400GSVHV'
    '    '    25 - '538 GSR'
    '    '    26 - '538 GSVHV'
    '    '    27 - '400FSR(OLD)'
    '    '    28 - '513GRS(OLD)'
    '    '    29 - '675HRS'
    '    '
    '    '   gas_frac_d       - газосодержание на входе в газосепаратор
    '    '   qliq_sm3day      - дебит жидкости в стандартных условиях
    '    '   freq_Hz          - частота врашения, Гц
    '    'description_end
    '    Dim GS As New CESPGasSep
    '    ESP_gassep_ksep_d = GS.my_sepfactor(gsep_type_TYPE, gas_frac_d * 100, qliq_sm3day, freq_Hz * 60) / 100

    'End Function

    ''description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    '' название газосопаратора
    'Public Function ESP_gassep_name(
    '                ByVal gsep_type_TYPE As Integer)
    '    ' MY_SEPFACTOR - Вычисление коэффициента сепрации в точке
    '    '   gsep_type_TYPE    - тип сепаратора (номер от 1 до 29)
    '    '    1  - 'GDNK5'
    '    '    2  - 'VGSA (VORTEX)'
    '    '    3  - 'GDNK5A'
    '    '    4  - 'GSA5-1'
    '    '    5  - 'GSA5-3'
    '    '    6  - 'GSA5-4'
    '    '    7  - 'GSAN-5A'
    '    '    8  - 'GSD-5A'
    '    '    9  - 'GSD5'
    '    '    10 - '3MNGB5'
    '    '    11 - '3MNGB5A'
    '    '    12 - '3MNGDB5'
    '    '    13 - '3MNGDB5A'
    '    '    14 - 'MNGSL5A-M'
    '    '    15 - 'MNGSL5A-TM'
    '    '    16 - 'MNGSL5-M'
    '    '    17 - 'MNGSL5-TM'
    '    '    18 - 'MNGSLM 5'
    '    '    19 - 'MNGD 5'
    '    '    20 - 'GSIK 5A'
    '    '    21 - '338DSR'
    '    '    22 - '400GSR'
    '    '    23 - '400GSV'
    '    '    24 - '400GSVHV'
    '    '    25 - '538 GSR'
    '    '    26 - '538 GSVHV'
    '    '    27 - '400FSR(OLD)'
    '    '    28 - '513GRS(OLD)'
    '    '    29 - '675HRS'
    '    'description_end
    '    Dim GS As New CESPGasSep
    '    ESP_gassep_name = GS.Separator_Name(gsep_type_TYPE)

    'End Function





    '=======================================================
    '--------------- Вспомогательные функции ---------------
    '=======================================================

    Private Sub read_ESP_calibr(ByVal c_calibr As Double,
                                ByRef c_calibr_head As Double,
                                ByRef c_calibr_rate As Double,
                                ByRef c_calibr_power As Double)


        c_calibr_head = 1
        c_calibr_rate = 1
        c_calibr_power = 1


        Dim clbr() As Double

        ' set calibration properties
        'clbr = array1d_from_range(c_calibr, num_only:=True, no_zero:=False)
        c_calibr_head = clbr(1)
        If clbr.GetUpperBound(0) >= 2 Then
            c_calibr_rate = clbr(2)
        Else
            c_calibr_rate = 1
        End If

        If clbr.GetUpperBound(0) >= 3 Then
            c_calibr_power = clbr(3)
        Else
            c_calibr_power = 1
        End If


    End Sub

End Module
'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
' класс для расчета многофазного потока в трубе (вторая версия после рефакторинга Алма)
'
' История
' 2016.04    Реализован новый механизм расчета распределения давления в стволе с использованием модуля решения ОДУ
'            Упрощена структура хранения массивов
' 2017.01    Модернизация под 7 версию. Исправление ошибок и контроль температуры
' 2019.04    Рефакторинг в сторону упрощения
' 2019.10    Рефакторинг в сторону упрощения - продолжение

Option Explicit On
Imports System.Math

Public Class CPipe

    Public ZeroCoordMes_m As Double                        ' начальная координата трубы, измеренная, от которой будут отсчитываться координаты в выходных массивах
    Public ZeroCoordVert_m As Double                       ' начальная координата трубы, вертикальная, от которой будут отсчитываться координаты в выходных массивах
    Public fluid As CPVT                                   ' базовый флюид в трубе. Определяет свойства и расходы и фазовый состав
    Public ambient_formation As New CAmbientFormation      ' порода за пределями скважины
    Public curve As New CCurves                           ' все кривые планируется прятать тут
    'Public curves As New CCurves
    Public t_calc_C As Double                              ' начальная температура флюида для расчета с учетом эмисии тепла

    Public p_result_atma As Double                         ' давление - результат расчета для отчета
    Public t_result_C As Double                            ' температура - результат расчета для отчета

    Private param_ As PARAMCALC                               ' параметры расчета по трубе

    ' геометрия трубы заданная массивами
    Private h_mes_insert_m_ As CInterpolation              ' измеренная глубина которую надо вставить в расчет трубы
    ' чтобы отловить изменение градиента температуры, например при динамическом уровне
    Private legth_total_m_ As Double                       ' общая длинна трубы
    Private depth_vert_total_m_ As Double                  ' общая глубина трубы
    ' расчетные параметры по трубе  (используются для вывода после проведения расчета)
    Private flow_params_out_() As PIPE_FLOW_PARAMS         ' расчетные параметры по трубе после расчета
    Private dTdLinit_ As Double                            ' распределение градиента температуры по длине начальное

    ' набор расчетных параметров по стволу скважины
    ' Private num_points_curve_ As Integer                    ' количество точек которые должны быть сохранены для распределения давления в трубе в итоговых кривых
    Private step_hm_curve_ As Double                        ' шаг для формирования выходного массива по трубе. м
    Private hm_curve_ As New CInterpolation                ' кривая для хранения набора точек, для которых должны строится все другие кривые
    ' поправочные коэффициенты для расчета распределения давления
    Private c_calibr_grav_ As Double
    Private c_calibr_fric_ As Double

    Public GLVin As CGLvalve  ' link to gas lift valve in pipe


    ' конструктор класса
    ' вызывается при создании класса - гарантирует что все объекты будут созданы
    Public Sub Class_Initialize(Optional ByVal GLVin_ As CGLvalve = Nothing,
                                 Optional ByVal ZeroCoordMes_m_ As Double = 0,
                                 Optional ByVal ZeroCoordVert_m_ As Double = 0,
                                 Optional ByVal legth_total_m As Double = 100,
                                 Optional ByVal StepHmCurve_ As Double = 1000,
                                 Optional ByVal c_calibr_grav As Double = 1,
                                 Optional ByVal c_calibr_fric As Double = 1,
                                 Optional ByVal step_hm_curve As Double = 0,
                                 Optional ByVal dTdLinit As Double = 0,
                                 Optional ByVal depth_vert_total_m As Double = 0,
                                 Optional ByVal t_result_C_ As Double = 0,
                                 Optional ByVal p_result_atma_ As Double = 0,
                                 Optional ByVal t_calc_C_ As Double = 0)

        GLVin = GLVin_
        ZeroCoordMes_m = ZeroCoordMes_m_
        ZeroCoordVert_m = ZeroCoordVert_m_
        legth_total_m_ = legth_total_m
        ReDim flow_params_out_(0)
        curve.Item("c_Roughness").isStepFunction = True ' шероховатость и диаметр трубы меняют ступенчато и не интерполируются
        curve.Item("c_Diam").isStepFunction = True
        curve.Item("c_fpat").isStepFunction = True
        fluid = New CPVT
        fluid.Class_Initialize()

        param_ = Set_calc_flow_param(calc_along_coord:=True,
                                flow_along_coord:=True,
                                hcor:=H_CORRELATION.Ansari,
                                temp_method:=TEMP_CALC_METHOD.StartEndTemp,
                                length_gas_m:=0)
        h_mes_insert_m_ = New CInterpolation
        Call h_mes_insert_m_.AddPoint(0, 0)
        Call h_mes_insert_m_.AddPoint(length_mes_m, 0)
        StepHmCurve = StepHmCurve_    ' по умолчанию шаг 100 м для сохранения кривых
        c_calibr_grav_ = c_calibr_grav
        c_calibr_fric_ = c_calibr_fric

        step_hm_curve_ = step_hm_curve
        depth_vert_total_m_ = depth_vert_total_m
        dTdLinit_ = dTdLinit
        t_result_C = t_result_C_
        p_result_atma = p_result_atma_
        t_calc_C = t_calc_C_

    End Sub

    'Public ReadOnly Property param() As PARAMCALC
    '    Get
    '        param = param_
    '    End Get
    'End Property

    Public Property param() As PARAMCALC
        Get
            param = param_
        End Get
        Set(value As PARAMCALC)
            param_ = value
            If value.length_gas_m > 0 Then
                ' if correlation division point given - stable points in pipe have to be updated
                Call add_h_mes_save_m(value.length_gas_m)
                StepHmCurve = StepHmCurve
            End If
        End Set
    End Property

    'Public ReadOnly Property c_calibr_grav() As Double
    '    Get
    '        c_calibr_grav = c_calibr_grav_
    '    End Get
    'End Property

    Public Property c_calibr_grav() As Double
        Get
            c_calibr_grav = c_calibr_grav_
        End Get
        Set(val As Double)
            If val >= 0 And val < 2 Then ' не стоит подстрочный коэффициент менять слишком сильно - если что тут можно поправить
                c_calibr_grav_ = val
            End If
        End Set
    End Property

    Public Property c_calibr_fric() As Double
        Get
            c_calibr_fric = c_calibr_fric_
        End Get
        Set(val As Double)
            If val >= 0 And val < 2 Then ' не стоит подстрочный коэффициент менять слишком сильно - если что тут можно поправить
                c_calibr_fric_ = val
            End If
        End Set
    End Property

    'Public ReadOnly Property c_calibr_fric() As Double
    '    Get
    '        c_calibr_fric = c_calibr_fric_
    '    End Get
    'End Property

    Public ReadOnly Property h_mes_save_m(ByVal i As Integer) As Double
        Get
            h_mes_save_m = h_mes_insert_m_.PointX(i)
        End Get
    End Property

    Public Function add_h_mes_save_m(ByVal val As Double) As Boolean
        If val > ZeroCoordMes_m And val < (ZeroCoordMes_m + length_mes_m) Then
            Call h_mes_insert_m_.AddPoint(val, 0)   ' Запишем точку, которую надо сохранить
        End If
        StepHmCurve = StepHmCurve
    End Function

    'Public ReadOnly Property StepHmCurve() As Double
    '    Get
    '        StepHmCurve = step_hm_curve_
    '    End Get
    'End Property

    Public Property StepHmCurve() As Double
        Get
            StepHmCurve = step_hm_curve_
        End Get
        Set(val As Double)
            Dim i As Integer
            Dim Hm As Double
            Dim Hm_max As Double

            step_hm_curve_ = val
            ' установили шаг - сразу подготовим массив точек по давлению для которых должен быть проведен расчет
            hm_curve_.ClearPoints()    ' очистили точки по давлению
            For i = 1 To h_mes_insert_m_.Num_points    ' пустили цикл по количеству точек, которые должны быть обязательно
                Hm = h_mes_insert_m_.PointX(i)
                hm_curve_.AddPoint(Hm, 0)               ' добвляем точку в выходной массив. Важен только x, поэтому y задаем произвольно
                ' здесь же должны задаться первая и последняя точки
            Next i
            ' далее добавим все промежуточные точки с заданым шагом
            i = 0
            Hm = hm_curve_.PointX(1)   ' начинаем с первой точки
            Hm_max = hm_curve_.PointX(hm_curve_.Num_points)
            Do
                Hm = Hm + StepHmCurve
                If Hm < Hm_max Then                     ' если новая точка попадает в диапазон, добавляем ее.
                    hm_curve_.AddPoint(Hm, 0)             ' здесь предполагается, что координаты будут возрастать
                End If                                  ' если такая точка есть, то она просто перезапишется
            Loop While Hm < Hm_max
            ' здесь получили в кривой hm_curve_ все точки для которых надо искать параметры

        End Set
    End Property

    ' длина сегмента трубы
    Public ReadOnly Property length_mes_m() As Double
        Get
            length_mes_m = legth_total_m_
        End Get
    End Property

    Public ReadOnly Property depth_vert_m() As Double
        Get
            depth_vert_m = depth_vert_total_m_
        End Get
    End Property

    Public Sub InitTlinearSmart(ByVal t_from_C As Double,
                            ByVal t_to_C As Double,
                            ByVal calc_flow_direction As Integer)

        If calc_flow_direction = 1 Or calc_flow_direction = 0 Then
            InitTlinear(t_to_C, t_from_C)
        Else
            InitTlinear(t_from_C, t_to_C)
        End If

    End Sub


    Public Sub InitTlinear(ByVal T_start_coord_C As Double, ByVal T_end_coord_C As Double)
        ' начальная инициализация распределения температуры в трубе
        If legth_total_m_ > 0 Then
            dTdLinit_ = (T_end_coord_C - T_start_coord_C) / legth_total_m_
        End If
        curve.Item("c_Tinit").ClearPoints()
        curve.Item("c_Tinit").AddPoint(ZeroCoordMes_m, T_start_coord_C)
        curve.Item("c_Tinit").AddPoint(ZeroCoordMes_m + length_mes_m, T_end_coord_C)
        curve.Item("c_Tinit").special = True
    End Sub

    ' general temperature initialisation fuction for UDF
    Public Sub InitT(t_calc_from_C As Double,
                t_val(,) As Double,
                calc_flow_direction As Integer,
                Optional temp_method As TEMP_CALC_METHOD = TEMP_CALC_METHOD.StartEndTemp)
        Dim t_calc_to_C As Double
        Dim temp_crv As New CInterpolation
        Dim amb As New CAmbientFormation
        Dim prm As PARAMCALC

        InitTlinear(t_calc_from_C, t_calc_from_C)
        t_calc_C = t_calc_from_C

        If t_val.GetUpperBound(2) = 1 Then
            t_calc_to_C = t_val(1, 1)

            InitTlinearSmart(t_calc_from_C, t_calc_to_C, calc_flow_direction)

        Else
            Call temp_crv.load_from_range(t_val)
            amb.amb_temp_curve = temp_crv
            ambient_formation = amb
            prm = param
            prm.temp_method = temp_method
            param = prm

        End If

    End Sub

    Private Function dTdL_linear_Cm(lmes_m As Double) As Double
        ' возвращает градиент температуры исходя из линейного приближения
        dTdL_linear_Cm = dTdLinit_
    End Function

    Private Function t_linear_C(lmes_m As Double) As Double
        ' возвращает температуру исходя из линейного приближения
        t_linear_C = curve.Item("c_Tinit").GetPoint(lmes_m)
    End Function

    Private Function t_amb_C(lmes_m As Double) As Double
        ' возвращает температуру исходя из окружения скважины
        Dim Hv_m As Double
        Hv_m = h_vert_h_mes_m(lmes_m)              ' определяем вертикальную глубину для заданной измеренной глубины
        t_amb_C = ambient_formation.amb_temp_C(Hv_m)
    End Function

    Private Function t_init_C(lmes_m As Double) As Double
        Select Case param.temp_method
            Case TEMP_CALC_METHOD.StartEndTemp
                t_init_C = t_linear_C(lmes_m) ' температуру берем извне
            Case TEMP_CALC_METHOD.GeoGradTemp
                t_init_C = t_amb_C(lmes_m)
            Case TEMP_CALC_METHOD.AmbientTemp
                t_init_C = t_amb_C(lmes_m)
        End Select
    End Function

    Private Function dTdL_amb_Cm(lmes_m As Double) As Double
        ' возвращает градиент температуры исходя из окружения
        Dim theta_deg As Double
        Dim Hv_m As Double
        theta_deg = angle_hmes_deg(lmes_m)         ' определяем наклон на заданной глубине
        Hv_m = h_vert_h_mes_m(lmes_m)              ' определяем вертикальную глубину для заданной измеренной глубины
        dTdL_amb_Cm = ambient_formation.amb_temp_grad_Cm(Hv_m) * sind(theta_deg)
    End Function

    Public Function d_hmes_m(ByVal z As Double) As Double
        ' функция возвращает внутренний диаметр трубы по заданной абсолютной измеренной глубине (если труба проходит по этой глубине)
        d_hmes_m = curve.Item("c_Diam").GetPoint(z)
    End Function

    Public Function roughness_h_mes_m(ByVal z As Double) As Double
        ' возвращает шероховатость по измеренной глубине
        roughness_h_mes_m = curve.Item("c_Roughness").GetPoint(z)
    End Function

    Public Function angle_hmes_deg(ByVal z As Double) As Double
        ' возвращает угол по измеренной глубине
        angle_hmes_deg = curve.Item("c_Theta").GetPoint(z)
    End Function

    Public Function h_vert_h_mes_m(ByVal z As Double) As Double
        ' возвращает угол по измеренной глубине
        h_vert_h_mes_m = curve.Item("c_Hvert").GetPoint(z)
    End Function

    Public Function p_h_mes_atma(ByVal z As Double) As Double
        ' возвращает угол по измеренной глубине
        p_h_mes_atma = curve.Item("c_P").GetPoint(z)
    End Function

    Public Function t_h_mes_C(ByVal z As Double) As Double
        ' возвращает угол по измеренной глубине
        t_h_mes_C = curve.Item("c_T").GetPoint(z)
    End Function

    ' инициализация трубы через данные по траектории скважины
    Public Sub init_pipe_constr_by_trajectory(ByVal tr As CPipeTrajectory,
                                              Optional ByVal HmesStart_m As Double = Nothing,
                                              Optional ByVal HmesEnd_m As Double = Nothing,
                                              Optional ByVal tr_cas As CPipeTrajectory = Nothing,
                                              Optional ByVal srv_points_step As Integer = 100)

        Dim i As Integer
        Dim h As Double
        Dim p_pipe_segments_num As Integer

        If HmesStart_m.ToString.Any Then HmesStart_m = tr.h_mes_m(0)
        If HmesEnd_m.ToString.Any Then HmesEnd_m = tr.h_mes_m(tr.num_points - 1)

        curve.Item("c_Diam").isStepFunction = True
        ZeroCoordMes_m = HmesStart_m
        ZeroCoordVert_m = tr.h_abs_hmes_m(HmesStart_m)
        p_pipe_segments_num = tr.num_points - 1
        ' по умолчанию используем все сегменты которые были заданы в траектории
        For i = 0 To p_pipe_segments_num + 2
            If i = 0 Then
                h = HmesStart_m
            ElseIf i = 1 Then
                h = HmesEnd_m
            Else
                h = tr.h_mes_m(i - 2)
            End If
            If h >= HmesStart_m And h <= HmesEnd_m Then
                ' теперь  заполним кривые соответствующие траектории скважины  - в первый раз пишем нулевую точку
                If tr_cas Is Nothing Then ' + tr_cas is missing
                    curve.Item("c_Diam").AddPoint(h, tr.diam_hmes_m(h))    ' НКТ
                Else
                    curve.Item("c_Diam").AddPoint(h, tr_cas.diam_hmes_m(h) - tr.diam_hmes_m(h))   ' затруб
                End If
                curve.Item("c_Roughness").AddPoint(h, tr.roughness_m)
                curve.Item("c_Theta").AddPoint(h, tr.ang_hmes_deg(h))
                curve.Item("c_Hvert").AddPoint(h, tr.h_abs_hmes_m(h))
            End If
        Next i
        curve.Item("c_Diam").special = True
        curve.Item("c_Roughness").special = True
        curve.Item("c_Theta").special = True
        curve.Item("c_Hvert").special = True

        legth_total_m_ = HmesEnd_m - HmesStart_m
        depth_vert_total_m_ = tr.h_abs_hmes_m(HmesEnd_m) - tr.h_abs_hmes_m(HmesStart_m)
        h_mes_insert_m_.ClearPoints()
        h_mes_insert_m_.AddPoint(ZeroCoordMes_m, 0)
        h_mes_insert_m_.AddPoint(ZeroCoordMes_m + length_mes_m, 0)
        Call add_h_mes_save_m(param_.length_gas_m)

        StepHmCurve = srv_points_step
    End Sub


    Public Function init_pipe(ByVal d_mm As Double,
                         ByVal length_m As Double,
                         ByVal theta_deg As Double,
                         Optional ByVal roughness_m As Double = 0.00001,
                         Optional Hmes0_m As Double = 0)
        ' простой метод инициализации трубы по двум точкам
        ' используется для учебных примеров в функциях Excel

        Dim arr_h(0 To 1, 0 To 1) As Double
        Dim arr_d(0 To 1, 0 To 2) As Double
        Dim tr As New CPipeTrajectory


        arr_h(0, 0) = 0
        arr_h(1, 0) = length_m
        arr_h(0, 1) = 0
        arr_h(1, 1) = length_m * (Sin(theta_deg / 180 * const_Pi))

        arr_d(0, 0) = 0
        arr_d(1, 0) = length_m
        arr_d(0, 1) = d_mm
        arr_d(1, 1) = d_mm
        arr_d(0, 2) = roughness_m
        arr_d(1, 2) = roughness_m


        Call tr.init_from_vert_range(arr_h, arr_d)
        tr.roughness_m = roughness_m

        Call init_pipe_constr_by_trajectory(tr)

    End Function


    '=================================================================================================
    ' новый подход - можно обойтись без разделение на прямой участок и кривой - расчет за один проход.
    '=================================================================================================
    Public Function calc_grad(l_m As Double,
                         p_atma As Double,
                         t_C As Double,
                         Optional calc_dtdl As Boolean = True,
                         Optional p_cas_atma As Double = 0.95) As PIPE_FLOW_PARAMS
        ' функция расчета градиента давления и температуры в скважине при заданных параметрах
        ' возвращает все параметры потока в заданной точке трубы при заданых термобарических условиях.
        '
        '
        '
        '  L_m      - измеренная глубина на которой ведется расчет, нужна для привязки по температуре
        '  p_atma   - давление в заданной точке
        '  T_C      - температура в заданной точке
        '  calc_dtdl
        '  p_cas_atma - затрубное давление для оптимизации расчета барботажа в затрубе


        'Allocate variables used to output auxilary values
        Dim dpdlg_out As Double
        Dim dpdlf_out As Double
        Dim dpdla_out As Double
        Dim v_sl_out As Double
        Dim v_sg_out As Double
        Dim vl_msec As Double
        Dim vg_msec As Double
        Dim h_l_out As Double
        Dim fpat_out
        Dim d_m As Double   ' диаметр трубы по которой идет поток
        Dim theta_deg As Double ' угол наклона трубы в расчете
        Dim theta_sign As Integer
        Dim rough_m As Double   ' шероховатость
        Dim Hv_m As Double

        Dim dp_dl As Double, dp_dl_arr(7) As Double
        Dim dt_dl As Double
        Dim v As Double, dvdL As Double

        If param_.FlowAlongCoord Then
            theta_sign = -1
        Else
            theta_sign = 1
        End If

        d_m = d_hmes_m(l_m)                ' определяем диаметр на указанной глубине
        theta_deg = theta_sign * angle_hmes_deg(l_m)   ' определяем наклон на заданной глубине
        rough_m = roughness_h_mes_m(l_m)       ' определяем шероховатость на заданной глубине
        Hv_m = h_vert_h_mes_m(l_m)              ' определяем вертикальную глубину для заданной измеренной глубины

        With fluid
            'проверим на корректность исходных данных
            If p_atma < const_minPpipe_atma Then
                dp_dl = 0
                GoTo endlab
            End If

            Call .Calc_PVT(p_atma, t_C)             ' найдем все PVT в заданных условиях
            If .Q_mix_rc_m3day = 0 Then
                t_C = ambient_formation.amb_temp_C(Hv_m)
            End If

            dp_dl_arr(0) = 0
            dp_dl_arr(1) = 0
            dp_dl_arr(2) = 0
            dp_dl_arr(3) = 0
            dp_dl_arr(4) = 0
            dp_dl_arr(5) = 0
            dp_dl_arr(6) = 0
            dp_dl_arr(7) = 101

            Dim corr As H_CORRELATION
            If ((l_m >= param_.start_length_gas_m) And (l_m < param_.start_length_gas_m + param_.length_gas_m)) Or fluid.gas_only Then
                corr = H_CORRELATION.gas
            Else
                corr = param_.correlation
            End If


            Select Case corr
                Case H_CORRELATION.BeggsBrill
                    dp_dl_arr = unf_BegsBrillGradient(d_m, theta_deg, rough_m, .Qliq_rc_m3day, .Q_gas_rc_m3day, .Mu_liq_cP, .Mu_gas_cP, .Sigma_liq_Nm,
                                  .Rho_liq_rc_kgm3, .Rho_gas_rc_kgm3, 0, 1, c_calibr_grav, c_calibr_fric)
                Case H_CORRELATION.Ansari

                    If p_atma > p_cas_atma Then
                        dp_dl_arr = unf_AnsariGradient(d_m, theta_deg, rough_m, .Qliq_rc_m3day, .Q_gas_rc_m3day, .Mu_liq_cP, .Mu_gas_cP, .Sigma_liq_Nm,
                                  .Rho_liq_rc_kgm3, .Rho_gas_rc_kgm3, p_atma, c_calibr_grav, c_calibr_fric)
                    End If
                Case H_CORRELATION.gas
                    If p_atma > p_cas_atma Then
                        dp_dl_arr = unf_GasGradient(d_m, theta_deg, rough_m, .Q_gas_rc_m3day, .Mu_gas_cP,
                                           .Rho_gas_rc_kgm3, p_atma)
                        ' gas gradient do not use calibration coeficients
                    End If
                Case H_CORRELATION.Unified
                    dp_dl_arr = unf_UnifiedTUFFPGradient(d_m, theta_deg, rough_m, .Qliq_rc_m3day, .Q_gas_rc_m3day, .Mu_liq_cP, .Mu_gas_cP, .Sigma_liq_Nm,
                                  .Rho_liq_rc_kgm3, .Rho_gas_rc_kgm3, p_atma, c_calibr_grav, c_calibr_fric)
                Case H_CORRELATION.Gray
                    dp_dl_arr = unf_GrayModifiedGradient(d_m, theta_deg, rough_m, .Qliq_rc_m3day, .Q_gas_rc_m3day, .Mu_liq_cP, .Mu_gas_cP, .Sigma_liq_Nm,
                                  .Rho_liq_rc_kgm3, .Rho_gas_rc_kgm3, 0, 1, , c_calibr_grav, c_calibr_fric)
                Case H_CORRELATION.HagedornBrown
                    dp_dl_arr = unf_HagedornandBrawnmodified(d_m, theta_deg, rough_m, .Qliq_rc_m3day, .Q_gas_rc_m3day, .Mu_liq_cP, .Mu_gas_cP, .Sigma_liq_Nm,
                                  .Rho_liq_rc_kgm3, .Rho_gas_rc_kgm3, p_atma, 0, 1, , c_calibr_grav, c_calibr_fric)
                Case H_CORRELATION.SakharovMokhov
                    dp_dl_arr = unf_Saharov_Mokhov_Gradient(d_m, theta_deg, rough_m, p_atma, .Q_oil_sm3day, .Q_wat_sm3day, .Q_gas_sm3day, .Bo_m3m3,
                                      .Bw_m3m3, .Bg_m3m3, .Rs_m3m3, .Mu_oil_cP, .Mu_wat_cP, .Mu_gas_cP, .Sigma_oil_gas_Nm, .Sigma_wat_gas_Nm, .Rho_oil_sckgm3, .Rho_wat_sckgm3, .Rho_gas_sckgm3,
                                      , , , c_calibr_grav, c_calibr_fric)
            End Select

            dp_dl = theta_sign * dp_dl_arr(0)

            dpdlg_out = theta_sign * dp_dl_arr(1)
            dpdlf_out = theta_sign * dp_dl_arr(2)
            dpdla_out = theta_sign * dp_dl_arr(3)
            v_sl_out = dp_dl_arr(4)
            v_sg_out = dp_dl_arr(5)
            h_l_out = dp_dl_arr(6)
            fpat_out = dp_dl_arr(7)

            vl_msec = v_sl_out * const_Pi * d_m ^ 2 / 4 ' скорость жидкости реальная
            vg_msec = v_sg_out * const_Pi * d_m ^ 2 / 4 ' скорость жидкости реальная
            ' для оценки температуры оценим скорость потока и ускорение

            ' теперь зададим изменение температуры в потоке
            If calc_dtdl Then
                Select Case param.temp_method
                    Case TEMP_CALC_METHOD.StartEndTemp
                        dt_dl = dTdL_linear_Cm(Hv_m)
                    Case TEMP_CALC_METHOD.GeoGradTemp
                        dt_dl = dTdL_amb_Cm(Hv_m)
                    Case TEMP_CALC_METHOD.AmbientTemp
                        v = vg_msec    ' оценка сверху
                        dvdL = -v / p_atma * dp_dl
                        dt_dl = ambient_formation.calc_dtdl_Cm(Hv_m, sind(theta_deg), t_C, .Wm_kgsec, .Cmix_JkgC,
                                                          dp_dl, v, dvdL, .CJT_Katm, param_.FlowAlongCoord)
                End Select
            End If

            ' тут надо записать в результаты все расчетные параметры
            Dim res As PIPE_FLOW_PARAMS



            res.md_m = l_m                                 ' pipe measured depth (from start - top)
            res.vd_m = Hv_m                                ' pipe vertical depth from start - top
            res.diam_mm = d_m * 1000
            res.dpdl_a_atmm = dpdla_out                    ' acceleration gradient at measured depth
            res.dpdl_f_atmm = dpdlf_out                    ' friction gradient at measured depth
            res.dpdl_g_atmm = dpdlg_out                    ' gravity gradient at measured depth
            res.fpat = fpat_out                            ' flow pattern code
            res.gasfrac = fluid.Gas_fraction_d()
            res.h_l_d = h_l_out                            ' liquid hold up
            res.Qg_m3day = fluid.Q_gas_rc_m3day
            res.p_atma = p_atma                              '  pipe pressure at measured depth
            res.t_C = t_C                                  ' pipe temp at measured depth
            res.v_sl_msec = v_sl_out                       ' superficial liquid velosity
            res.v_sg_msec = v_sg_out                       ' superficial gas velosity
            res.thete_deg = theta_deg                      '
            res.roughness_m = rough_m                      '
            res.rs_m3m3 = fluid.Rs_m3m3                    ' растворенный газ в нефти в потоке
            res.gasfrac = fluid.Gas_fraction_d              ' расходное содержание газа в потоке
            res.mu_oil_cP = fluid.Mu_oil_cP                      ' вязкость нефть в потоке
            res.mu_wat_cP = fluid.Mu_wat_cP                      ' вязкость воды в потоке
            res.mu_gas_cP = fluid.Mu_gas_cP                      ' вязкость газа в потоке
            res.mu_mix_cP = fluid.Mu_mix_cP                  ' вязкость смеси в потоке
            res.Rhoo_kgm3 = fluid.Rho_oil_rc_kgm3             ' плотность нефти
            res.Rhow_kgm3 = fluid.Rho_wat_rc_kgm3           ' плотность воды
            res.rhol_kgm3 = fluid.Rho_liq_rc_kgm3             ' плотность жидкости
            res.Rhog_kgm3 = fluid.Rho_gas_rc_kgm3             ' плотность газа
            res.rhomix_kgm3 = fluid.Rho_mix_rc_kgm3           ' плотность смеси в потоке
            res.q_oil_m3day = fluid.Q_oil_rc_m3day                  ' расход нефти в рабочих условиях
            res.qw_m3day = fluid.Q_wat_rc_m3day                  ' расход воды в рабочих условиях
            res.Qg_m3day = fluid.Q_gas_rc_m3day            ' расход газа в рабочих условиях
            res.mo_kgsec = fluid.Mo_kgsec                  ' массовый расход нефти в рабочих условиях
            res.mw_kgsec = fluid.Mw_kgsec                  ' массовый расход воды в рабочих условиях
            res.mg_kgsec = fluid.Mg_kgsec                  ' массовый расход газа в рабочих условиях
            res.vl_msec = vl_msec  ' скорость жидкости реальная
            res.vg_msec = vg_msec  ' скорость газа реальная

endlab:
            res.dp_dl = dp_dl
            res.dt_dl = dt_dl
            calc_grad = res

        End With

    End Function

    Public Sub set_ZNLF()

        param_.correlation = H_CORRELATION.Ansari
        If param_.temp_method = 2 Then param_.temp_method = 1
        fluid = fluid.Clone
        fluid.qliq_sm3day = const_ZNLF_rate
        fluid.Fw_fr = 0 ' только нефть при барботаже
    End Sub


    Public Function calc_dPipe(ByVal p_atma As Double,
                               Optional ByVal t_C As Double = Nothing,
                               Optional ByVal saveCurve As CALC_RESULTS = CALC_RESULTS.nocurves) As PTtype
        ' здесь выбираем метод расчета
        ' если не надо рассчитывать эмисию тепла - то можно расчет делать только по давлению - это быстрее
        ' если температуру рассчитываем то решаем систему и по давлению и по температуре - медленнее
        ' если для расчета нужна стартовая температура флюида то берется из t_calc_C
        Dim PTres As PTtype

        If Not t_C.ToString.Any Then
            t_calc_C = t_C
        End If

        If length_mes_m = 0 Then
            PTres.p_atma = p_atma
            PTres.t_C = t_calc_C
        Else
            If param.temp_method = TEMP_CALC_METHOD.AmbientTemp And param.CalcAlongCoord = param.FlowAlongCoord Then
                ' расчет температуры с учетом эмисиии тепла в окружающее пространство возможен
                ' только если расчет делается по направлению потока
                PTres = calc_dPipe_2d(p_atma, t_calc_C, saveCurve)
            Else
                PTres = calc_dPipe_1d(p_atma, saveCurve)
            End If
        End If
        p_result_atma = PTres.p_atma
        t_result_C = PTres.t_C
        calc_dPipe = PTres
    End Function



    Private Function calc_dPipe_2d(p_atma As Double, t_C As Double, Optional saveCurve As CALC_RESULTS = CALC_RESULTS.nocurves) As PTtype
        ' новая версия расчета перепада давления в трубе, сразу с учетом инклинометрии
        ' основан на применении ODEsolver
        ' PT   - термобарические условия в точке задания условия по давлению
        ' SaveCurve - флаг показывающий необходимость сохранения детальных результатов расчета
        ' t_other_C  - опциональное значение температуры на другом конце трубы, необходимо при линейном
        '             распределении температуры

        Dim Y0(1) As Double   ' начальные значения для проведения расчета
        Dim N, M As Integer
        Dim X() As Double, y(,) As Double     ' массив глубин для которых нужны значения
        Dim eps As Double
        Dim Stepp As Double
        Dim State As odesolverstate
        Dim Rtn As Boolean
        Dim i As Integer
        Dim pfp As PIPE_FLOW_PARAMS
        Dim Rep As odesolverreport
        Dim stPt As Boolean


        'ReDim Y0(1)

        Try
            eps = const_pressure_tolerance '0.01
            Stepp = 10
            Y0(0) = p_atma
            Y0(1) = t_C
            N = 2                   ' размер системы  - две переменные - давление и температура
            M = hm_curve_.Num_points ' количество точек для которых надо выдать ответ
            ' формируем массив глубин для расчета давления
            ' учитываем, что массив глубин зависит от направления в котором отсчитываем координаты
            ReDim X(M - 1)

            If param.CalcAlongCoord Then
                For i = 0 To M - 1
                    X(i) = hm_curve_.PointX(i + 1)
                Next i
            Else
                For i = 0 To M - 1
                    X(i) = hm_curve_.PointX(M - i)
                Next i
            End If
            ' проверка - если поток в скважине нулевой, тогда температура равна температуре окружающей среды
            ' без такой проверки расчет градиента температуры сходит с ума
            If fluid.qliq_sm3day = 0 Then
                Y0(1) = ambient_formation.amb_temp_C(X(0))
            End If

            '   Y = solve_ode("calc_grad_2d", Y0, x, coeffA, Eps)

            Call odesolverrkck(Y0, N, X, M, eps, Stepp, State)
            ' Loop through the AlgLib solver routine and the external ODE
            ' evaluation routine until the solver routine returns "False",
            ' which indicates that it has finished.
            ' The VBA function named in "FuncName" is called using
            ' the Application.Run method.
            Rtn = True
            i = 0
            Do While Rtn = True And i < 10000
                Rtn = odesolveriteration(State)
                pfp = calc_grad(State.csobj.x, State.csobj.y(0), State.csobj.y(1))    ' Application.Run(FuncName, State.X, State.Y, CoeffA)
                State.csobj.dy(0) = pfp.dp_dl
                State.csobj.dy(1) = pfp.dt_dl
                i = i + 1
            Loop
            ' Extract the desired results from the State
            ' object using the appropriate AlgLib routine
            Call odesolverresults(State, M, X, y, Rep)
            ' If necessary convert the AlgLib output array(s) to
            ' a form suitable for Excel.  In this case YA2 is
            ' a 2D base 0 array, which may be assigned to the
            ' function return value without further processing.
            ' Assign the output array to the function return value
            ' ODE = YA2

            ' подговим выходные результаты функции
            calc_dPipe_2d.p_atma = y(M - 1, 0)
            calc_dPipe_2d.t_C = y(M - 1, 1)
            If saveCurve > 0 Then
                ' сохраним результаты расчета для отображения на графиках
                curve.Item("c_P").ClearPoints()
                curve.Item("c_T").ClearPoints()
                curve.Item("c_Tamb").ClearPoints()

                For i = 0 To M - 1
                    stPt = i = 0 Or i = M - 1
                    curve.Item("c_P").AddPoint(X(i), y(i, 0), stPt)
                    curve.Item("c_T").AddPoint(X(i), y(i, 1), stPt)
                    If param.temp_method = TEMP_CALC_METHOD.AmbientTemp Then
                        curve.Item("c_Tamb").AddPoint(X(i), ambient_formation.amb_temp_C(curve.Item("c_Hvert").GetPoint(X(i))), stPt)
                    Else
                        curve.Item("c_Tamb").AddPoint(X(i), y(i, 1), stPt)
                    End If
                Next i
                curve.Item("c_P").special = True
                curve.Item("c_T").special = True
                curve.Item("c_Tamb").special = True
                If saveCurve > 1 Then
                    Call FillDetailedCurve()
                End If
            End If
            Exit Function
        Catch ex As Exception
            Dim errmsg As String
            errmsg = "CPipe.calc_dPipe_2d: ошибка какая то"
            Throw New ApplicationException(errmsg)
        End Try

    End Function



    Private Function calc_dPipe_1d(p_atma As Double, Optional saveCurve As CALC_RESULTS = CALC_RESULTS.nocurves) As PTtype
        ' новая версия расчета перепада давления в трубе, сразу с учетом инклинометрии
        ' основан на применении ODEsolver
        ' проверка работы одномерного решателя - ради скорости расчета

        Dim Y0(0) As Double   ' начальные значения для проведения расчета
        Dim N, M As Integer
        Dim X() As Double, y(,) As Double     ' массив глубин для которых нужны значения
        Dim eps As Double
        Dim Stepp As Double
        Dim State As odesolverstate
        Dim Rtn As Boolean
        Dim i As Integer
        Dim pfp As PIPE_FLOW_PARAMS
        Dim Rep As odesolverreport
        Dim stPt As Boolean

        Try

        Catch ex As Exception
            Dim errmsg As String
            errmsg = "CPipe.calc_dPipe_1d: error -> "
            Throw New ApplicationException(errmsg)
        End Try
        eps = const_pressure_tolerance  '0.001
        Stepp = 10
        Y0(0) = p_atma

        '    Y0(1) = PT.T_C
        N = 1                   ' размер системы  - одна переменные - давление и температура
        M = hm_curve_.Num_points ' количество точек для которых надо выдать ответ
        ' формируем массив глубин для расчета давления
        ' учитываем, что массив глубин зависит от направления в котором отсчитываем координаты
        ReDim X(M - 1)
        If param.CalcAlongCoord Then
            For i = 0 To M - 1
                X(i) = hm_curve_.PointX(i + 1)
            Next i
        Else
            For i = 0 To M - 1
                X(i) = hm_curve_.PointX(M - i)
            Next i
        End If
        ' проверка - если поток в скважине нулевой, тогда температура равна температуре окружающей среды
        ' без такой проверки расчет градиента температуры сходит с ума
        '    If fluid.qliq_sm3day = 0 Then
        '        Y0(1) = ambient_formation.amb_temp_C(X(0))
        '    End If
        Call odesolverrkck(Y0, N, X, M, eps, Stepp, State)
        ' Loop through the AlgLib solver routine and the external ODE
        ' evaluation routine until the solver routine returns "False",
        ' which indicates that it has finished.
        ' The VBA function named in "FuncName" is called using
        ' the Application.Run method.
        Rtn = True
        i = 0
        Do While Rtn = True And i < 10000
            Rtn = odesolveriteration(State)
            If State.csobj.y(0) < const_minPpipe_atma Then
                ' при расчете давления получили отрицательные значения
                ' может происходить про расчете в затрубе
                ' тогда имитируем правильное завершение работы цикла
                State.csobj.y(0) = 0
                'State.RepTerminationType = 2
                'Rtn = False
            End If
            pfp = calc_grad(State.csobj.x, State.csobj.y(0), t_init_C(State.csobj.x), calc_dtdl:=False)
            State.csobj.dy(0) = pfp.dp_dl
            i = i + 1
        Loop
        ' Extract the desired results from the State
        ' object using the appropriate AlgLib routine
        Call odesolverresults(State, M, X, y, Rep)
        ' If necessary convert the AlgLib output array(s) to
        ' a form suitable for Excel.  In this case YA2 is
        ' a 2D base 0 array, which may be assigned to the
        ' function return value without further processing.
        ' Assign the output array to the function return value
        ' ODE = YA2
        ' Debug.Print i    ' 133  iteration approximatly min
        ' подговим выходные результаты функции
        calc_dPipe_1d.p_atma = y(M - 1, 0)
        calc_dPipe_1d.t_C = t_init_C(X(M - 1))
        If saveCurve > 0 Then
            ' сохраним результаты расчета для отображения на графиках
            curve.Item("c_P").ClearPoints()
            curve.Item("c_T").ClearPoints()
            curve.Item("c_Tamb").ClearPoints()

            For i = 0 To M - 1
                '            stPt = h_mes_insert_m_.TestPoint(x(i)) >= 0
                If h_mes_insert_m_.TestPoint(X(i)) >= 0 Then
                    stPt = True
                Else
                    stPt = False
                End If
                curve.Item("c_P").AddPoint(X(i), y(i, 0), stPt)
                curve.Item("c_T").AddPoint(X(i), t_init_C(X(i)), stPt)
                curve.Item("c_Tamb").AddPoint(X(i), t_init_C(X(i)), stPt)
            Next i
            curve.Item("c_P").special = True
            curve.Item("c_T").special = True
            curve.Item("c_Tamb").special = True
            If saveCurve > 1 Then
                Call FillDetailedCurve()
            End If
        End If
        Exit Function

    End Function

    Private Sub FillDetailedCurve()
        ' функция расчета  детальных распределений параметров по длине трубы
        Dim i As Integer
        Dim M As Integer
        Dim FlowParams_out As PIPE_FLOW_PARAMS

        M = curve.Item("c_P").Num_points
        Call curve.ClearPoints() ' ClearPoints_unprotected
        For i = 1 To M
            FlowParams_out = calc_grad(curve.Item("c_P").PointX(i),
                                    curve.Item("c_P").PointY(i),
                                    curve.Item("c_T").PointY(i))
            With FlowParams_out
                curve.Item("c_udl_m").AddPoint(.md_m, .md_m - .vd_m)
                curve.Item("c_dpdl_g").AddPoint(.md_m, .dpdl_g_atmm)
                curve.Item("c_dpdl_f").AddPoint(.md_m, .dpdl_f_atmm)
                curve.Item("c_dpdl_a").AddPoint(.md_m, .dpdl_a_atmm)
                curve.Item("c_vsl").AddPoint(.md_m, .v_sl_msec)
                curve.Item("c_vsg").AddPoint(.md_m, .v_sg_msec)
                curve.Item("c_Hl").AddPoint(.md_m, .h_l_d)
                curve.Item("c_fpat").AddPoint(.md_m, .fpat)
                'curve("c_Theta").AddPoint .md_m, .thete_deg
                'curve("c_Roughness").AddPoint .md_m, .roughness_m
                curve.Item("c_Rs").AddPoint(.md_m, .rs_m3m3)
                curve.Item("c_gasfrac").AddPoint(.md_m, .gasfrac)
                curve.Item("c_muo").AddPoint(.md_m, .mu_oil_cP)
                curve.Item("c_muw").AddPoint(.md_m, .mu_wat_cP)
                curve.Item("c_mug").AddPoint(.md_m, .mu_gas_cP)
                curve.Item("c_mumix").AddPoint(.md_m, .mu_mix_cP)
                curve.Item("c_rhoo").AddPoint(.md_m, .Rhoo_kgm3)
                curve.Item("c_rhow").AddPoint(.md_m, .Rhow_kgm3)
                curve.Item("c_rhol").AddPoint(.md_m, .rhol_kgm3)
                curve.Item("c_rhog").AddPoint(.md_m, .Rhog_kgm3)
                curve.Item("c_rhomix").AddPoint(.md_m, .rhomix_kgm3)
                curve.Item("c_qo").AddPoint(.md_m, .q_oil_m3day)
                curve.Item("c_qw").AddPoint(.md_m, .qw_m3day)
                curve.Item("c_qg").AddPoint(.md_m, .Qg_m3day)
                curve.Item("c_mo").AddPoint(.md_m, .mo_kgsec)
                curve.Item("c_mw").AddPoint(.md_m, .mw_kgsec)
                curve.Item("c_mg").AddPoint(.md_m, .mg_kgsec)
                curve.Item("c_vl").AddPoint(.md_m, .vl_msec)
                curve.Item("c_vg").AddPoint(.md_m, .vg_msec)
            End With
        Next i
    End Sub

    Public ReadOnly Property p_curve() As CInterpolation
        Get
            p_curve = curve.Item("c_P")
        End Get
    End Property


    Public Function array_out(Optional ByVal num_points As Integer = 20,
                          Optional ByVal all_curves_out As Boolean = False)
        ' подготовка массива для вывода в Excel
        ' num_points - количество точек в выходных массивах для вывода
        '
        Dim arr(,)
        Dim M As Integer
        '  Dim FlowParams_out As PIPE_FLOW_PARAMS
        Dim offset As Integer
        Dim i As Integer
        Dim hh As Double

        offset = 2

        Dim crv_P As CInterpolation
            Dim crv_T As CInterpolation

            ' rearrange output curves one time here - will be used later
            crv_P = curve.Item("c_P").ClonePointsToNum(num_points)
            crv_T = curve.Item("c_T").ClonePointsToNum(num_points)

            M = crv_P.Num_points
            If Not all_curves_out Then
                ReDim arr(M + offset, 8)
            Else
                ReDim arr(M + offset, 40)    ' get ready to output all
            End If

            arr(0, 0) = p_result_atma
            arr(0, 1) = t_result_C
            arr(0, 2) = crv_P.PointY(1)
            arr(0, 3) = crv_T.PointY(1)
            arr(0, 4) = crv_P.PointY(M)
            arr(0, 5) = crv_T.PointY(M)
            arr(0, 6) = c_calibr_grav_
            arr(0, 7) = c_calibr_fric_
            arr(0, 8) = fluid.Q_gas_sm3day

            If all_curves_out Then
                arr(0, 9) = 0
                arr(0, 10) = 0
                arr(0, 11) = 0
                arr(0, 12) = 0
                arr(0, 13) = 0
                arr(0, 14) = 0
                arr(0, 15) = 0
                arr(0, 16) = 0
                arr(0, 17) = 0
                arr(0, 18) = 0
                arr(0, 19) = 0
                arr(0, 20) = 0
                arr(0, 21) = 0
                arr(0, 22) = 0
                arr(0, 23) = 0
                arr(0, 24) = 0
                arr(0, 25) = 0
                arr(0, 26) = 0
                arr(0, 27) = 0
                arr(0, 28) = 0
                arr(0, 29) = 0
                arr(0, 30) = 0
                arr(0, 31) = 0
                arr(0, 32) = 0
                arr(0, 33) = 0
                arr(0, 34) = 0
                arr(0, 35) = 0
                arr(0, 36) = 0
                arr(0, 37) = 0
                arr(0, 38) = 0
                arr(0, 39) = 0
                arr(0, 40) = 0
            End If



            arr(1, 0) = "p_result_atma"
            arr(1, 1) = "t_result_C"
            arr(1, 2) = "p_1, atma"
            arr(1, 3) = "t_1, C"
            arr(1, 4) = "p_2, atma"
            arr(1, 5) = "t_2, C"
            arr(1, 6) = "c_calibr_grav"
            arr(1, 7) = "c_calibr_fric"
            arr(1, 8) = "q_gas_sm3day"

            If all_curves_out Then
                arr(1, 9) = "q_gas_sm3day"
                arr(1, 10) = 0
                arr(1, 11) = 0
                arr(1, 12) = 0
                arr(1, 13) = 0
                arr(1, 14) = 0
                arr(1, 15) = 0
                arr(1, 16) = 0
                arr(1, 17) = 0
                arr(1, 18) = 0
                arr(1, 19) = 0
                arr(1, 20) = 0
                arr(1, 21) = 0
                arr(1, 22) = 0
                arr(1, 23) = 0
                arr(1, 24) = 0
                arr(1, 25) = 0
                arr(1, 26) = 0
                arr(1, 27) = 0
                arr(1, 28) = 0
                arr(1, 29) = 0
                arr(1, 30) = 0
                arr(1, 31) = 0
                arr(1, 32) = 0
                arr(1, 33) = 0
                arr(1, 34) = 0
                arr(1, 35) = 0
                arr(1, 36) = 0
                arr(1, 37) = 0
                arr(1, 38) = 0
                arr(1, 39) = 0
                arr(1, 40) = 0
            End If


            arr(offset, 0) = "num"
            arr(offset, 1) = "h,m"
            arr(offset, 2) = "hvert,m"
            arr(offset, 3) = "p,atma"
            arr(offset, 4) = "t,C"
            arr(offset, 5) = "Hl"
            arr(offset, 6) = "fpat"
            arr(offset, 7) = "t_amb, C"
            arr(offset, 8) = "diam, m"

            If all_curves_out Then
                arr(offset, 9) = "c_Roughness" 'python_output
                arr(offset, 10) = "c_Theta"
                arr(offset, 11) = "c_Tinit"
                arr(offset, 12) = "c_P"
                arr(offset, 13) = "c_T"
                arr(offset, 14) = "c_Tamb"
                arr(offset, 15) = "c_udl_m"
                arr(offset, 16) = "c_dpdl_g"
                arr(offset, 17) = "c_dpdl_f"
                arr(offset, 18) = "c_dpdl_a"
                arr(offset, 19) = "c_vsl"
                arr(offset, 20) = "c_vsg"
                arr(offset, 21) = "c_Hl"
                arr(offset, 22) = "c_gasfrac"
                arr(offset, 23) = "c_muo"
                arr(offset, 24) = "c_muw"
                arr(offset, 25) = "c_mug"
                arr(offset, 26) = "c_mumix"
                arr(offset, 27) = "c_rhoo"
                arr(offset, 28) = "c_rhow"
                arr(offset, 29) = "c_rhol"
                arr(offset, 30) = "c_rhog"
                arr(offset, 31) = "c_rhomix"
                arr(offset, 32) = "c_qo"
                arr(offset, 33) = "c_qw"
                arr(offset, 34) = "c_qg"
                arr(offset, 35) = "c_mo"
                arr(offset, 36) = "c_mw"
                arr(offset, 37) = "c_mg"
                arr(offset, 38) = "c_vl"
                arr(offset, 39) = "c_vg"
                arr(offset, 40) = "c_Rs"
            End If

            For i = 1 To M
                arr(offset + i, 0) = i
                hh = crv_P.PointX(i)
                arr(offset + i, 1) = hh
                arr(offset + i, 2) = curve.Item("c_Hvert").GetPoint(hh)
                arr(offset + i, 3) = crv_P.PointY(i)
                arr(offset + i, 4) = curve.Item("c_T").GetPoint(hh)
                arr(offset + i, 5) = curve.Item("c_Hl").GetPoint(hh)
                arr(offset + i, 6) = curve.Item("c_fpat").GetPoint(hh)
                arr(offset + i, 7) = curve.Item("c_Tamb").GetPoint(hh)
                arr(offset + i, 8) = curve.Item("c_Diam").GetPoint(hh)
                If all_curves_out Then
                    arr(offset + i, 9) = curve.Item("c_Roughness").GetPoint(hh)
                    arr(offset + i, 10) = curve.Item("c_Theta").GetPoint(hh)
                    arr(offset + i, 11) = curve.Item("c_Tinit").GetPoint(hh)
                    arr(offset + i, 12) = curve.Item("c_P").GetPoint(hh)
                    arr(offset + i, 13) = curve.Item("c_T").GetPoint(hh)
                    arr(offset + i, 14) = curve.Item("c_Tamb").GetPoint(hh)
                    arr(offset + i, 15) = curve.Item("c_udl_m").GetPoint(hh)
                    arr(offset + i, 16) = curve.Item("c_dpdl_g").GetPoint(hh)
                    arr(offset + i, 17) = curve.Item("c_dpdl_f").GetPoint(hh)
                    arr(offset + i, 18) = curve.Item("c_dpdl_a").GetPoint(hh)
                    arr(offset + i, 19) = curve.Item("c_vsl").GetPoint(hh)
                    arr(offset + i, 20) = curve.Item("c_vsg").GetPoint(hh)
                    arr(offset + i, 21) = curve.Item("c_Hl").GetPoint(hh)
                    arr(offset + i, 22) = curve.Item("c_gasfrac").GetPoint(hh)
                    arr(offset + i, 23) = curve.Item("c_muo").GetPoint(hh)
                    arr(offset + i, 24) = curve.Item("c_muw").GetPoint(hh)
                    arr(offset + i, 25) = curve.Item("c_mug").GetPoint(hh)
                    arr(offset + i, 26) = curve.Item("c_mumix").GetPoint(hh)
                    arr(offset + i, 27) = curve.Item("c_rhoo").GetPoint(hh)
                    arr(offset + i, 28) = curve.Item("c_rhow").GetPoint(hh)
                    arr(offset + i, 29) = curve.Item("c_rhol").GetPoint(hh)
                    arr(offset + i, 30) = curve.Item("c_rhog").GetPoint(hh)
                    arr(offset + i, 31) = curve.Item("c_rhomix").GetPoint(hh)
                    arr(offset + i, 32) = curve.Item("c_qo").GetPoint(hh)
                    arr(offset + i, 33) = curve.Item("c_qw").GetPoint(hh)
                    arr(offset + i, 34) = curve.Item("c_qg").GetPoint(hh)
                    arr(offset + i, 35) = curve.Item("c_mo").GetPoint(hh)
                    arr(offset + i, 36) = curve.Item("c_mw").GetPoint(hh)
                    arr(offset + i, 37) = curve.Item("c_mg").GetPoint(hh)
                    arr(offset + i, 38) = curve.Item("c_vl").GetPoint(hh)
                    arr(offset + i, 39) = curve.Item("c_vg").GetPoint(hh)

                    arr(offset + i, 40) = curve.Item("c_Rs").GetPoint(hh)
                End If

            Next i
        array_out = arr
    End Function
End Class

'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
'
'
' класс для моделирования работы погружной части ЭЦН
' описывает работу набора одинаковых ступеней

'==============  CESPpump  ==============
' 
Option Explicit On
Imports System.Math
Imports System.IO
Imports Newtonsoft.Json
Imports System.Linq
Imports System.Data.Linq

Public Class CESPpump
    ' геометрические параметры насоса
    Public h_mes_top_m As Double           ' глубина установки ЭЦН (по верхней части)

    Public angle_deg As Double             ' угол установки УЭЦН (предполагается, что по глубине угол не меняется) ' not used for 7.24

    ' общие параметры
    Public fluid As New CPVT                   ' флюид движущийся через насос (с учтом сепарации газа)
    Public c_calibr_head As Double         ' деградация характеристики УЭЦН по напору 
    Public c_calibr_rate As Double         ' деградация характеристики УЭЦН по дебиту
    Public c_calibr_power As Double        ' деградация по мощности (она же по КПД системы)
    ' набор параметров для задания поведения в зоне турбинного вращения (когда дебит больше максимального)
    Public turb_head_factor As Double
    Public turb_rate_factor As Double
    Public dnum_stages_integrate As Integer

    Public freq_Hz As Double               ' частота вращения вала насоса (испольщуется для расчета)
    Public curves As New CCurves           ' все кривые планируется прятать тут
    ' параметры конструкции ЭЦН
    Public stage_num As Integer            ' количество ступеней в насосе (испольщуется для расчета характеристики насоса)

    Private t_int_C_ As Double             ' температура потока на приемной сетке УЭЦН (учитывается нагрев двигателем)
    Private t_dis_C_ As Double             ' температура потока на выкиде насоса (учитывается нагрев в насосе)
    ' параметры работы насоса для которых был проведен расчет
    Private p_int_atma_ As Double          ' давление на приеме насоса (используется для расчета рабочих характеристик)
    Private p_dis_atma_ As Double          ' давление на выкиде насоса

    Private power_fluid_Wt_ As Double      ' Мощность передаваемая УЭЦН жидкости
    Private power_ESP_Wt_ As Double        ' Мощность потребляемая ЭЦН с вала (механическая)
    Private eff_ESP_d_ As Double           ' КПД УЭЦН по факту

    Private head_real_m_ As Double
    ' параметры определяющие установку УЭЦН
    Private db_ As ESP_PARAMS            ' набор параметров насоса из базы данных
    Private db_json_string As String

    Private gas_frac_intake_ As Double
    Private gas_corr_ As Double
    Public gas_correct As Double       ' тип для коррекции по газу
    ' ESP_gas_correct       - тип насоса по работе с газом
    '      ESP_gas_correct = 0 нет коррекции
    '      ESP_gas_correct = 1 стандартный ЭЦН (предел 25%)
    '      ESP_gas_correct = 2 ЭЦН с газостабилизирующим модулем (предел 50%)
    '      ESP_gas_correct = 3 ЭЦН с осевым модулем (предел 75%)
    '      ESP_gas_correct = 4 ЭЦН с модифицированными ступенями (предел 40%)
    Private correct_visc_ As Boolean
    Private corr_visc_h_ As Double          ' поправочный коэффициент для напорной характеристики на вязкость для текущего дебита и текущего расчета
    Private corr_visc_q_ As Double          ' поправочный коэффициент для дебита
    Private corr_visc_pow_ As Double        ' поправочный коэффициент для мощности
    Private corr_visc_eff_ As Double        ' поправочный коэффициент для КПД
    Private h_corr_qd_curve_ As New CInterpolation     ' зависимость поправочного коэффицента для напора от дебита (для расчета по модели американского института нефти)
    Private p_curve_ As New CInterpolation  ' кривая распределения давления вдоль насоса (как снаружи, так и внутри)
    Private t_curve_ As New CInterpolation  ' кривая распределения температуры флюида вдоль насоса

    Private calc_from_dis_ As Boolean

    Public gassep_M_Nm As Double

    Public ESP_base_dictionary As New Dictionary(Of String, ESP_dict)
    Public qliq_m3day As Double

    Public Sub Class_Initialize(Optional ByVal correct_visc As Boolean = True,
                                 Optional ByVal c_calibr_head_ As Double = 1,
                                 Optional ByVal c_calibr_rate_ As Double = 1,
                                 Optional ByVal c_calibr_power_ As Double = 1,
                                 Optional ByVal stage_num_ As Integer = 1,
                                 Optional ByVal freq_Hz_ As Double = 50,
                                 Optional ByVal gas_correct_ As Double = 1,
                                 Optional ByVal turb_head_factor_ As Double = 1,
                                 Optional ByVal turb_rate_factor_ As Double = 1,
                                 Optional ByVal dnum_stages_integrate_ As Integer = 1,
                                 Optional ByVal h_mes_top_m_ As Double = 1000,
                                 Optional ByVal corr_visc_eff As Double = 1,
                                 Optional ByVal angle_deg_ As Double = 0,
                                 Optional ByVal t_int_C As Double = 0,
                                 Optional ByVal t_dis_C As Double = 0,
                                 Optional ByVal p_int_atma As Double = 0,
                                 Optional ByVal p_dis_atma As Double = 0,
                                 Optional ByVal power_fluid_Wt As Double = 0,
                                 Optional ByVal power_ESP_Wt As Double = 0,
                                 Optional ByVal eff_ESP_d As Double = 0,
                                 Optional ByVal head_real_m As Double = 0,
                                 Optional ByVal db_json_string_ As String = "",
                                 Optional ByVal gas_frac_intake As Double = 0,
                                 Optional ByVal gas_corr As Double = 0,
                                 Optional ByVal gassep_M_N_ As Double = 0,
                                 Optional ByVal calc_from_dis As Boolean = False,
                                 Optional ByVal qliq_m3day_ As Double = 0)

        correct_visc_ = correct_visc

        c_calibr_head = c_calibr_head_ ' по умолчанию нет деградации
        c_calibr_rate = c_calibr_rate_ ' по умолчанию нет деградации
        c_calibr_power = c_calibr_power_ ' по умолчанию нет деградации

        stage_num = stage_num_
        freq_Hz = freq_Hz_

        Call corrections_clear()
        corr_visc_eff_ = corr_visc_eff
        gas_correct = gas_correct_

        turb_head_factor = turb_head_factor_ ' 2 ' 0.5
        turb_rate_factor = turb_rate_factor_ ' 1.1 '0.9
        dnum_stages_integrate = dnum_stages_integrate_

        h_mes_top_m = h_mes_top_m_

        angle_deg = angle_deg_
        t_int_C_ = t_int_C
        t_dis_C_ = t_dis_C
        p_int_atma_ = p_int_atma
        p_dis_atma_ = p_dis_atma
        power_fluid_Wt_ = power_fluid_Wt
        power_ESP_Wt_ = power_ESP_Wt
        eff_ESP_d_ = eff_ESP_d
        head_real_m_ = head_real_m
        db_json_string = db_json_string_
        gas_frac_intake_ = gas_frac_intake
        gas_corr_ = gas_corr
        gassep_M_Nm = gassep_M_N_
        calc_from_dis_ = calc_from_dis
        qliq_m3day = qliq_m3day_
    End Sub

    Private Sub corrections_clear(Optional ByVal corr_visc_h As Double = 1,
                                  Optional ByVal corr_visc_q As Double = 1,
                                  Optional ByVal corr_visc_pow As Double = 1)

        corr_visc_h_ = corr_visc_h
        corr_visc_q_ = corr_visc_q
        corr_visc_pow_ = corr_visc_pow
        '    corr_visc_eff_ = 1
        '    c_calibr_head = 1
        '    c_calibr_power = 1
        '    c_calibr_rate = 1

    End Sub
    ' заменил название из-за let, проверить 
    Public ReadOnly Property db() As ESP_PARAMS
        Get
            db = db_
        End Get
    End Property

    Public Property db_let As ESP_PARAMS
        Get
            Return db_
        End Get
        Set(ByVal val As ESP_PARAMS)
            db_ = val
        End Set
    End Property
    Public ReadOnly Property correct_visc() As Boolean
        Get
            correct_visc = correct_visc_
        End Get
    End Property

    Public Property correct_visc_let As Boolean
        Get
            Return correct_visc_
        End Get
        Set(value As Boolean)
            correct_visc_ = value
            If Not correct_visc_ Then
                Call corrections_clear()
            End If
        End Set
    End Property

    ' =======================  геометрия
    Public ReadOnly Property length_m() As Double
        Get
            length_m = db_.stage_height_m * stage_num
        End Get
    End Property

    ' глубина нижней точки установки
    Public ReadOnly Property h_mes_down_m() As Double
        Get
            h_mes_down_m = h_mes_top_m + length_m
        End Get
    End Property

    ' функция для расчета высоты сборки из num ступеней
    Private Function stages_heigth_m(ByVal num As Double) As Double
        If num <= stage_num Then
            stages_heigth_m = length_m / stage_num * num
        Else
            stages_heigth_m = length_m
        End If
    End Function

    ' свойство для расчета измеренной глубины расположения i ступени
    Public ReadOnly Property HmesStage_m(i As Double) As Double
        Get
            HmesStage_m = h_mes_down_m - stages_heigth_m(i) ' тут надо отнять длину ступеней выше контрольной
        End Get
    End Property

    Public ReadOnly Property area_shaft_m2() As Double
        Get
            area_shaft_m2 = db_.d_shaft_m * db_.d_shaft_m / 4 * const_Pi
        End Get
    End Property

    Public ReadOnly Property angle_vert_deg() As Double
        Get
            angle_vert_deg = angle_vert_deg - 90
        End Get
    End Property

    ' ========================  конец блока описания геометрии
    Public ReadOnly Property head_m() As Double
        Get
            head_m = head_real_m_
        End Get
    End Property

    Public Function points_num() As Integer
        points_num = db_.head_points.GetUpperBound(0) + 1
    End Function

    Public ReadOnly Property eff_ESP_d() As Double
        Get
            eff_ESP_d = eff_ESP_d_
        End Get
    End Property

    Public ReadOnly Property power_fluid_W() As Double
        Get
            power_fluid_W = power_fluid_Wt_
        End Get
    End Property

    Public ReadOnly Property power_ESP_W() As Double
        Get
            power_ESP_W = power_ESP_Wt_
        End Get
    End Property

    Public ReadOnly Property p_int_atma() As Double
        Get
            p_int_atma = p_int_atma_
        End Get
    End Property

    Public ReadOnly Property p_dis_atma() As Double
        Get
            p_dis_atma = p_dis_atma_
        End Get
    End Property

    Public ReadOnly Property t_int_C() As Double
        Get
            t_int_C = t_int_C_
        End Get
    End Property

    Public ReadOnly Property t_dis_C() As Double
        Get
            t_dis_C = t_dis_C_
        End Get
    End Property
    ' заменил название из-за let, проверить 
    Public ReadOnly Property w_obmin() As Double
        Get
            w_obmin = freq_Hz * 60
        End Get
    End Property

    Public Property w_obmin_let As Double
        Get
            Return freq_Hz
        End Get
        Set(ByVal val As Double)
            freq_Hz = val / 60
        End Set
    End Property

    Public ReadOnly Property w_radsec() As Double
        Get
            w_radsec = freq_Hz * 2 * const_Pi
        End Get
    End Property

    Public ReadOnly Property rate_max_sm3day(ByVal mu_cSt As Double) As Double
        Get
            If correct_visc_ And (mu_cSt > 0) Then        ' если большая вязкость - сделаем коррекцию
                Call calc_CorrVisc_PetrInst(0, mu_cSt)   ' метод меняет константы класса, которые влияют на характеристики насоса
            End If
            rate_max_sm3day = db_.rate_max_sm3day * freq_Hz / db_.freq_Hz * corr_visc_q_
        End Get
    End Property

    Public ReadOnly Property rate_nom_sm3day(Optional ByVal mu_cSt As Double = -1) As Double
        Get
            If correct_visc_ And (mu_cSt > 0) Then        ' если большая вязкость - сделаем коррекцию
                Call calc_CorrVisc_PetrInst(0, mu_cSt)   ' метод меняет константы класса, которые влияют на характеристики насоса
            End If
            rate_nom_sm3day = db_.rate_nom_sm3day * freq_Hz / db_.freq_Hz * corr_visc_q_
        End Get
    End Property

    'Public Property Get gas_degr() As Double
    '   gas_degr = gas_degr_
    'End Property
    '
    'Public Property Let gas_degr(val As Double)
    '   If val >= 0 Then
    '       gas_degr_ = val
    '   End If
    'End Property

    Private Function calc_ESP_head_nominal_m(ByVal q_m3day As Double, Optional ByVal stage_num As Integer = 1) As Double
        ' функция для расчета номинального напора насоса
        Dim b As Double                  ' отношение частот
        With db_
            b = .freq_Hz / freq_Hz  ' определим отношение реальной частоты УЭЦН к номинальной для которой заданы характеристики
            calc_ESP_head_nominal_m = b ^ (-2) * stage_num * crv_interpolation(.rate_points, .head_points, b * q_m3day, 2)
            calc_ESP_head_nominal_m = calc_ESP_head_nominal_m '* corr_visc_h_  ' учтем коррекцию на вязкость
        End With
    End Function

    Public Function get_ESP_head_m(ByVal q_m3day As Double, Optional ByVal stage_num As Integer = -1, Optional ByVal mu_cSt As Double = -1) As Double
        'Dim b As Double                 ' отношение частот
        Dim stage_num_to_calc As Integer ' число ступеней с которым будет проводиться расчет
        Dim maxQ As Double
        Dim q_calc_m3day As Double

        Call corrections_clear()

        If q_m3day < 0 Then             ' проверим исходные данные на релевантность
            get_ESP_head_m = 0
            AddLogMsg("CPumpESP.get_ESP_head_m: расчет характеристики насоса с отрицательным дебитом  Q_m3day = " & String.Format("{0:0.##}", q_m3day) & "Напор установлен = 0")

            Exit Function
        End If
        ' определяем число ступеней с которым будем проводить расчет
        If stage_num > 0 Then           ' если в явном виде задан параметр то его используем
            stage_num_to_calc = stage_num
        Else                            ' иначе использует количество ступеней из характеристики насоса
            stage_num_to_calc = Me.stage_num
        End If

        If correct_visc_ And (mu_cSt > 0) Then   ' если большая вязкость - сделаем коррекцию
            Call calc_CorrVisc_PetrInst(q_m3day, mu_cSt)   ' метод меняет константы класса, которые влияют на характеристики насоса
        End If

        q_calc_m3day = q_m3day / corr_visc_q_    ' делаем коррекцию по вязкости для дебита
        maxQ = db_.rate_max_sm3day * freq_Hz / db_.freq_Hz                   ' здесь коррекция на вязкость тоже уже учтена
        If q_calc_m3day < maxQ Then
            get_ESP_head_m = calc_ESP_head_nominal_m(q_calc_m3day, stage_num_to_calc)
        ElseIf maxQ - turb_rate_factor * (q_calc_m3day - maxQ) > 0 Then
            ' apply correction for far rigth interval
            get_ESP_head_m = -turb_head_factor * calc_ESP_head_nominal_m(maxQ - turb_rate_factor * (q_calc_m3day - maxQ), stage_num_to_calc)
        Else
            get_ESP_head_m = -turb_head_factor * calc_ESP_head_nominal_m(0, stage_num_to_calc)
        End If
        get_ESP_head_m = get_ESP_head_m * corr_visc_h_
    End Function

    Private Sub calc_CorrVisc_PetrInst(ByVal q_mix_ As Double, ByVal nu_cSt As Double)
        ' метод для расчета корректировки напорной характеристики УЭЦН на вязкость для текущего насоса
        ' расчет для одной ступени

        Dim GAMMA As Double
        Dim QwBEP_100gpm As Double, HwBEP_ft As Double
        Dim Qstar As Double
        Dim Q0 As Double, Q0_6 As Double, Q0_8 As Double, Q1_0 As Double, Q1_2 As Double, qmax As Double
        Dim H0 As Double, H0_6 As Double, H0_8 As Double, H1_0 As Double, H1_2 As Double, Hmax As Double

        Dim corr_visc_h__ As Double    ' поправочный коэффициень для напорной характеристики на вязкость для текущего дебита и текущего расчета
        Dim corr_visc_q__ As Double    ' для дебита
        Dim corr_visc_pow__ As Double  ' для мощности
        Dim corr_visc_eff__ As Double  ' для КПД
        Try
            ' turn off object correction factors
            corr_visc_h_ = 1             ' поправочный коэффициень для напорной характеристики на вязкость для текущего дебита и текущего расчета
            corr_visc_q_ = 1               ' для дебита
            corr_visc_pow_ = 1             ' для мощности
            corr_visc_eff_ = 1             ' для КПД

            ' turn off local corr factors as well
            corr_visc_h__ = 1             ' поправочный коэффициень для напорной характеристики на вязкость для текущего дебита и текущего расчета
            corr_visc_q__ = 1               ' для дебита
            corr_visc_pow__ = 1             ' для мощности
            corr_visc_eff__ = 1             ' для КПД

            If nu_cSt < 5 Then Exit Sub

            QwBEP_100gpm = Me.rate_nom_sm3day * const_convert_m3day_gpm '/ 100   '   похоже к книге Такаса ошибка - не надо делить на 100 тут
            HwBEP_ft = Me.get_ESP_head_m(Me.rate_nom_sm3day, 1) * const_convert_m_ft
            GAMMA = -7.5946 + 6.6504 * Log(HwBEP_ft) + 12.8429 * Log(QwBEP_100gpm)
            Qstar = Exp((39.5276 + 26.5606 * Log(nu_cSt) - GAMMA) / 51.6565)
            corr_visc_q__ = 1 - 4.0327 * 10 ^ (-3) * Qstar - 1.724 * 10 ^ (-4) * Qstar ^ 2

            If (corr_visc_q__ < 0) Then
                corr_visc_h__ = 0
                'exit without changes to object state
                Exit Sub
            End If

            corr_visc_eff__ = 1 - 3.3075 * 10 ^ (-2) * Qstar + 2.8875 * 10 ^ (-4) * Qstar ^ 2
            corr_visc_pow__ = 1 / corr_visc_eff__


            Q0 = 0
            ' rate_nom_sm3day has inside correction corr_visc_q_ - but not here
            Q1_0 = rate_nom_sm3day * corr_visc_q__
            H1_0 = 1 - 7.00763 * 10 ^ (-3) * Qstar - 1.41 * 10 ^ (-5) * Qstar ^ 2
            Q0_8 = Q1_0 * 0.8
            H0_8 = 1 - 4.4726 * 10 ^ (-3) * Qstar - 4.18 * 10 ^ (-5) * Qstar ^ 2
            Q0_6 = Q1_0 * 0.6
            H0_6 = 1 - 3.68 * 10 ^ (-3) * Qstar - 4.36 * 10 ^ (-5) * Qstar ^ 2
            Q1_2 = Q1_0 * 1.2
            H1_2 = 1 - 9.01 * 10 ^ (-3) * Qstar + 1.31 * 10 ^ (-5) * Qstar ^ 2
            qmax = rate_max_sm3day(-1) * corr_visc_q__
            Hmax = H1_2


            If qmax < Q1_2 Then
                AddLogMsg("CESPpump.calc_CorrVisc_PetrInst error. qmax >= Qmom * 1.2. Correction neglected")
                Exit Sub
                ' тут что то не так с характеристиков насоса - номинальный и максимальный дебит не соответствуют друг другу
            End If

            h_corr_qd_curve_.ClearPoints()

            'Call h_corr_qd_curve_.AddPoint(Qmax, Hmax)
            Call h_corr_qd_curve_.AddPoint(Q1_2, H1_2)
            Call h_corr_qd_curve_.AddPoint(Q1_0, H1_0)
            Call h_corr_qd_curve_.AddPoint(Q0_8, H0_8)
            Call h_corr_qd_curve_.AddPoint(Q0_6, H0_6)
            H0 = h_corr_qd_curve_.GetPoint(Q0) ' пытаемся экстраполировать
            If H0 < 0 Then H0 = H0_6
            Call h_corr_qd_curve_.AddPoint(Q0, H0)

            If q_mix_ > qmax Then q_mix_ = qmax

            corr_visc_h__ = h_corr_qd_curve_.GetPoint(q_mix_)

            corr_visc_h_ = corr_visc_h__             ' поправочный коэффициень для напорной характеристики на вязкость для текущего дебита и текущего расчета
            corr_visc_q_ = corr_visc_q__               ' для дебита
            corr_visc_pow_ = corr_visc_pow__             ' для мощности
            corr_visc_eff_ = corr_visc_eff__             ' для КПД

            Exit Sub
        Catch ex As Exception
            Dim msg As String
            msg = "CESPpump.calc_CorrVisc_PetrInst error with params q_mix_ = " & CStr(q_mix_) & " nu_cSt = " & CStr(nu_cSt)

            Throw New ApplicationException(msg)
        End Try
    End Sub

    Private Sub read_json_to_dict()
        ' read ESP database in json format
        ' ESP database in json format can be prepared with ESP_db.xlsm file

        'Dim ss As String
        'Dim line_from_file As String
        'Dim lines_all As String
        Dim fname As String

        fname = Directory.GetCurrentDirectory & esp_db_name
        Console.WriteLine(fname)
        Try
            'Using reader As New StreamReader(fname)
            '    lines_all = reader.ReadToEnd
            'End Using
            'open fname for input as #1
            'Do While Not eof(1)
            '    line input #1, line_from_file
            '    lines_all = lines_all & line_from_file & vbCrLf
            'Loop
            'close #1
            'Dim list As New List(Of ESP_dict)
            Dim json As String
            json = File.ReadAllText(fname)
            ESP_base_dictionary = JsonConvert.DeserializeObject(Of Dictionary(Of String, ESP_dict))(json)
            'ESP_base_dictionary = JsonConvert.DeserializeObject(lines_all)

            Exit Sub
        Catch ex As Exception
            Dim msg As String
            msg = "read_json_to_dict:" & fname & "read error"

            Throw New ApplicationException(msg)
        End Try
    End Sub

    Public Sub set_ID(ESP_ID As String)

        Dim esp_db As ESP_PARAMS
        Dim dict As ESP_dict
        Dim j As Integer, num As Integer

        If ESP_base_dictionary.Count = 0 Then
            Call read_json_to_dict()
        End If
        'dict = CType((From esp In ESP_base_dictionary
        'Where esp.ID = ESP_ID
        'Select Case esp.Data), Data_prop)
        dict = ESP_base_dictionary.Item(ESP_ID)
        'dict = ESP_base_dictionary.ID(ESP_ID)
        db_json_string = JsonConvert.SerializeObject(dict)
        Try
            num = dict.rate_points.Count
            With esp_db

                ReDim .head_points(0 To num)
                ReDim .rate_points(0 To num)
                ReDim .power_points(0 To num)
                ReDim .eff_points(0 To num)

                For j = 0 To num - 1
                    .head_points(j) = dict.head_points(j)
                    .rate_points(j) = dict.rate_points(j)
                    .power_points(j) = dict.power_points(j)
                    .eff_points(j) = dict.eff_points(j)
                Next j

                ' read all data from first line in DB table
                .ID = ESP_ID
                .manufacturer = dict.manufacturer
                .name = dict.name
                .stages_max = dict.stages_max
                .rate_nom_sm3day = dict.rate_nom_sm3day
                .rate_opt_min_sm3day = dict.rate_opt_min_sm3day
                .rate_opt_max_sm3day = dict.rate_opt_max_sm3day
                .rate_max_sm3day = dict.rate_max_sm3day
                .slip_nom_rpm = dict.slip_nom_rpm
                .freq_Hz = dict.freq_Hz
                .eff_max = dict.eff_max
                .d_od_m = dict.d_od_mm / 1000 ' читаем габарит насоса
                .d_motor_od_m = dict.d_motor_od_mm / 1000  ' читаем габарит насоса
                .d_cas_min_m = dict.d_cas_min_mm / 1000
                .d_shaft_m = dict.d_shaft_mm / 1000
                .power_limit_shaft_kW = dict.power_limit_shaft_kW
                .power_limit_shaft_max_kW = dict.power_limit_shaft_max_kW
                .pressure_limit_housing_atma = dict.pressure_limit_housing_atma

                If Not dict.height_stage_m = 0 Then
                    .height_stage_m = dict.height_stage_m
                Else
                    .height_stage_m = 0.05   ' по умолчанию 5 см высота. Можно сделать зависимость от дебита
                End If
            End With

            db_ = esp_db

            Exit Sub
        Catch ex As Exception
            Dim msg As String
            msg = "CESPpump.set_ID error - Problem while loading pump . " & esp_db.ID

            Throw New ApplicationException(msg)
        End Try

    End Sub

    Public Sub set_num_stages(head_m As Double)
        '  функция расчета необходимого числа ступеней для обеспечения заданного напора
        Dim Head1st As Double
        Head1st = get_ESP_head_m(rate_nom_sm3day, 1)
        If Head1st > 0 Then
            stage_num = CInt(head_m / Head1st)
        End If
    End Sub

    Public Sub init_json(json As String)

        Dim d As Dictionary(Of String, ESP_dict)

        d = CType(JsonConvert.DeserializeObject(json), Dictionary(Of String, ESP_dict))

        Call init_dictionary(d)
    End Sub

    Private Sub init_dictionary(dict As Dictionary(Of String, ESP_dict))

        Dim ESP_ID As String
        Dim head_nom_m As Double

        Try

            For Each esp As KeyValuePair(Of String, ESP_dict) In dict
                If esp.Value.ID.ToString IsNot Nothing Then
                    ESP_ID = esp.Value.ID
                    Call set_ID(ESP_ID)
                Else
                    AddLogMsg("CESPpump.init_dictionary error - wrong input. No ESP_ID key in pump json")
                    Throw New Exception
                End If
                stage_num = 0
                If esp.Value.num_stages.ToString IsNot Nothing Then
                    stage_num = esp.Value.num_stages
                End If
                If stage_num <= 0 And esp.Value.head_nom_m.ToString IsNot Nothing Then
                    head_nom_m = esp.Value.head_nom_m
                    Call set_num_stages(head_nom_m)
                End If
                If stage_num <= 0 Then
                    AddLogMsg("CESPpump.init_dictionary error - wrong input. No stages in ESP defined")
                End If
                If esp.Value.ID = ESP_ID Then
                    If esp.Value.freq_Hz.ToString IsNot Nothing Then
                        freq_Hz = esp.Value.freq_Hz
                    End If
                    If esp.Value.gas_correct.ToString IsNot Nothing Then
                        gas_correct = esp.Value.gas_correct
                    End If
                    If esp.Value.c_calibr_head.ToString IsNot Nothing Then
                        c_calibr_head = esp.Value.c_calibr_head
                    End If
                    If esp.Value.c_calibr_rate.ToString IsNot Nothing Then
                        c_calibr_rate = esp.Value.c_calibr_rate
                    End If
                    If esp.Value.c_calibr_power.ToString IsNot Nothing Then
                        c_calibr_power = esp.Value.c_calibr_power
                    End If
                    If esp.Value.dnum_stages_integrate.ToString IsNot Nothing Then
                        dnum_stages_integrate = esp.Value.dnum_stages_integrate
                    End If
                End If
            Next

            '    If Not IsMissing(dict_gassep) Then
            '        If Not dict_gassep Is Nothing Then
            '            With dict_gassep
            '                If .Exists("gassep_type") Then gassep_type = .Item("gassep_type")
            '                If .Exists("ksep_man_d") Then gassep_ksep_man_d = .Item("ksep_man_d")
            '                gassep_M_Nm = 0
            '                If .Exists("M_Nm") Then gassep_M_Nm = .Item("M_Nm")
            '            End With
            '        End If
            '    End If

            Exit Sub
        Catch ex As Exception
            Dim msg As String
            msg = "Error:CESPpump.init_dictionary: init error " & sDELIM

            Throw New ApplicationException(msg)
        End Try
    End Sub

    Public Sub calc_ESP(p_atma As Double,
                    t_intake_C As Double,
           Optional t_dis_C As Double = 0,
           Optional calc_from_intake As Boolean = True,
           Optional saveCurve As Boolean = False,
           Optional f_Hz As Double = -1)
        ' метод расчета работы насоса
        If f_Hz > 0 Then
            Me.freq_Hz = f_Hz
        End If
        Call ESP_dPIntegration(p_atma, t_intake_C, t_dis_C, Not calc_from_intake, saveCurve)
    End Sub

    Private Sub ESP_dPIntegration(ByVal p_atma As Double,
                              ByVal t_intake_C As Double,
                   Optional t_dis_C As Double = 0,
                   Optional calc_from_dis As Boolean = False,
                   Optional saveCurve As Boolean = False)
        ' Функция расчете распределения давления в УЭЦН - расчет снизу вверх от входного давления до выходного
        ' заодно считает и потребляемую мощность и КПД установки
        ' p_atma         pressure at pump intake
        ' t_intake_C          temprature at pump intake
        ' t_dis_C         температура на выходе, если задана учитывается, если нет то рассчитывается
        ' calc_from_dis  показывает будет ли предпринята попытка проинтегрировать сверху вниз насос
        ' p_int_estimation_atma приближения для давления на приеме, используется для расчета сверху вниз

        Dim i As Integer
        Dim head_mix As Double
        Dim dPStage As Double
        Dim PowfluidWt As Double, PowfluidTot_Wt As Double  ' полезная мощность передаваемая насосом жидкости
        Dim PowESP_Wt As Double, PowESPTot_Wt As Double     ' механическая мощность потребляемая насосом
        Dim EffESP_d As Double      ' КПД УЭЦН
        Dim EffStage As Double
        Dim dTpump_C As Double, dTpumpSum_C As Double
        Dim Pst_atma As Double
        Dim Tst_C As Double         ' температура по ступеням
        Dim sign_int As Integer
        Dim q_mix_ As Double, q_mix__degr As Double
        Dim dNst As Integer  ' шаг ускорения при интегрировании большими шагами
        Dim Nst As Integer   ' шаг на текущей итерации
        Dim N As Integer     ' текущий номер ступени
        Dim nn As Integer    ' номер ступени от приема для записи в архив
        Dim dPav As Double   ' поправки на давление и температуру при интегрировании
        Dim dTav As Double

        dNst = dnum_stages_integrate ' для начала пытаемся интегрировать такими шагами
        ' если тут поставить 10 будет быстрее считать за счет снижение числа шагов
        ' при этом может копиться ошибка, особенно если мало ступеней
        'dNst = 10 ' может быть надо будет когда то сделать глобальную настройку по скорости точности
        dPav = 0
        dTav = 0
        gas_corr_ = 1
        Try
            calc_from_dis_ = calc_from_dis  ' save state to object

            ' rearrange variables to calc direction
            If calc_from_dis Then
                If t_dis_C < 0 Then t_dis_C = t_intake_C
                Tst_C = t_dis_C
                p_dis_atma_ = p_atma
                sign_int = -1
            Else
                Tst_C = t_intake_C
                p_int_atma_ = p_atma
                sign_int = 1
            End If

            ' init auxiliary variables
            Call corrections_clear()
            Pst_atma = p_atma
            dTpumpSum_C = 0
            head_real_m_ = 0
            t_int_C_ = t_intake_C
            t_dis_C_ = t_dis_C
            PowfluidWt = 0
            PowfluidTot_Wt = 0
            PowESP_Wt = 0
            PowESPTot_Wt = 0
            dTpumpSum_C = 0
            N = 0
            i = 0
            If saveCurve Then
                curves.Item("gas_fractionInPump").ClearPoints()
                curves.Item("PressureInPump").ClearPoints()
                curves.Item("TempInPump").ClearPoints()
                curves.Item("PowerfluidInPump").ClearPoints()
                curves.Item("PowerESPInPump").ClearPoints()
                curves.Item("EffESPInPump").ClearPoints()
                curves.Item("q_mix_InPump").ClearPoints()

                curves.Item("mu_stage_cP").ClearPoints()
                curves.Item("corr_visc_h_").ClearPoints()
                curves.Item("corr_visc_q_").ClearPoints()
                curves.Item("corr_visc_pow_").ClearPoints()
                curves.Item("corr_visc_eff_").ClearPoints()
                curves.Item("gas_corr_").ClearPoints()

                p_curve_.ClearPoints()
                t_curve_.ClearPoints()
            End If

            ' init gas correction coefficient
            ' if calc from intake - it will be reinitialased later
            ' if calc from discharge - it can be set manually here
            gas_corr_ = GasCorrection_d(0, gas_correct)

            With fluid
                ' if  calc from intake - then save intake condition to 0 stage
                ' and calc intake gas fraction  and gas correction coefficient

                If Not calc_from_dis And saveCurve Then
                    ' calc PVT for intake conditions
                    Call .Calc_PVT(Pst_atma, Tst_C)    ' calc properties for inteke conditions
                    gas_frac_intake_ = .Gas_fraction_d
                    gas_corr_ = GasCorrection_d(gas_frac_intake_, gas_correct) ' reinit gas correction

                    curves.Item("gas_fractionInPump").AddPoint(N, .F_g)
                    curves.Item("PressureInPump").AddPoint(N, Pst_atma)
                    curves.Item("TempInPump").AddPoint(N, Tst_C)
                    curves.Item("PowerfluidInPump").AddPoint(N, 0)
                    curves.Item("PowerESPInPump").AddPoint(N, 0)
                    curves.Item("EffESPInPump").AddPoint(N, 0)
                    curves.Item("q_mix_InPump").AddPoint(N, .Q_mix_rc_m3day)

                    curves.Item("mu_stage_cP").AddPoint(N, .Mu_mix_cP)
                    curves.Item("corr_visc_h_").AddPoint(N, corr_visc_h_)
                    curves.Item("corr_visc_q_").AddPoint(N, corr_visc_q_)
                    curves.Item("corr_visc_pow_").AddPoint(N, corr_visc_pow_)
                    curves.Item("corr_visc_eff_").AddPoint(N, corr_visc_eff_)
                    curves.Item("gas_corr_").AddPoint(N, gas_corr_)

                    p_curve_.AddPoint(HmesStage_m(CDbl(N)), Pst_atma)
                    t_curve_.AddPoint(HmesStage_m(CDbl(N)), Tst_C)
                End If

                If calc_from_dis And saveCurve Then
                    ' calc PVT for intake conditions
                    Call .Calc_PVT(Pst_atma, Tst_C)    ' calc properties for inteke conditions

                    curves.Item("gas_fractionInPump").AddPoint(stage_num, .F_g)
                    curves.Item("PressureInPump").AddPoint(stage_num, Pst_atma)
                    curves.Item("TempInPump").AddPoint(stage_num, Tst_C)
                    curves.Item("PowerfluidInPump").AddPoint(stage_num, 0)
                    curves.Item("PowerESPInPump").AddPoint(stage_num, 0)
                    curves.Item("EffESPInPump").AddPoint(stage_num, 0)
                    curves.Item("q_mix_InPump").AddPoint(stage_num, .Q_mix_rc_m3day)

                    curves.Item("mu_stage_cP").AddPoint(stage_num, .Mu_mix_cP)
                    curves.Item("corr_visc_h_").AddPoint(stage_num, corr_visc_h_)
                    curves.Item("corr_visc_q_").AddPoint(stage_num, corr_visc_q_)
                    curves.Item("corr_visc_pow_").AddPoint(stage_num, corr_visc_pow_)
                    curves.Item("corr_visc_eff_").AddPoint(stage_num, corr_visc_eff_)
                    curves.Item("gas_corr_").AddPoint(stage_num, gas_corr_)

                    p_curve_.AddPoint(HmesStage_m(stage_num), Pst_atma)
                    t_curve_.AddPoint(HmesStage_m(stage_num), Tst_C)
                End If

                ' ====================== start main loop ================================================
                Do While N < stage_num '+ 1
                    If calc_from_dis Then
                        If stage_num - N - dNst > 0 Then  ' смотрим какой будет величина следующего шага
                            Nst = dNst                          ' мелкие шаги оставляем в зоне низкий давлений
                        Else
                            Nst = 1
                        End If
                    Else
                        If (stage_num - N) Mod dNst = 0 Then  ' смотрим какой будет величина следующего шага
                            Nst = dNst                          ' мелкие шаги оставляем в зоне низкий давлений
                        Else
                            Nst = 1
                        End If
                    End If

                    Call .Calc_PVT(Pst_atma + dPav, Tst_C + dTav)  ' делаем поправку на давление и температуру

                    ' re init gas correction for each stage mode if applicable
                    If gas_correct > 99 Then
                        gas_corr_ = GasCorrection_d(.Gas_fraction_d, gas_correct - 100)
                    End If

                    q_mix_ = .Q_mix_rc_m3day
                    q_mix__degr = q_mix_ * c_calibr_rate
                    head_mix = get_ESP_head_m(q_mix__degr, Nst, .Mu_mix_cSt) * c_calibr_head * gas_corr_
                    head_real_m_ = head_real_m_ + head_mix
                    dPStage = .Rho_mix_rc_kgm3 * head_mix * const_g * const_convert_Pa_atma ' тут когда то надо сделать коррекцию характеристики на плотность
                    Pst_atma = Pst_atma + sign_int * dPStage
                    dPav = dPStage / 2 * sign_int
                    If dPStage > 0 Then
                        ' оценим работу совершаемую насосом по перекачке жидкости
                        PowfluidWt = q_mix_ * const_convert_m3day_m3sec * dPStage * const_convert_atma_Pa   ' мощность с поправкой на плотность ГЖС
                        PowfluidTot_Wt = PowfluidTot_Wt + PowfluidWt
                        ' оценим мощность потребляемую насосом с вала
                        PowESP_Wt = get_ESP_power_W(q_mix__degr, Nst, .Mu_mix_cSt) * .Rho_mix_rc_kgm3 / 1000 * c_calibr_power                ' мощность потребляемая одной ступенью на воде
                        PowESPTot_Wt = PowESPTot_Wt + PowESP_Wt
                        ' оценим КПД ступени в данных условиях
                        If (PowESPTot_Wt > 0) Then
                            EffESP_d = PowfluidTot_Wt / PowESPTot_Wt
                        Else
                            EffESP_d = 0
                        End If

                        If (PowESP_Wt > 0) Then
                            EffStage = PowfluidWt / PowESP_Wt
                        Else
                            EffStage = 0
                        End If

                        If t_dis_C <= 0 And (Not calc_from_dis) Then ' оценка температуры по ступеням
                            If EffStage > 0 Then
                                dTpump_C = const_g * head_mix / .Cmix_JkgC * (1 - EffStage) / EffStage
                            Else
                                dTpump_C = 0
                            End If
                        Else
                            dTpump_C = (t_dis_C - t_intake_C) / stage_num * Nst
                        End If

                        If Tst_C < 299 Then
                            Tst_C = Tst_C + sign_int * dTpump_C
                            dTav = sign_int * dTpump_C / 2
                        End If

                        If Tst_C > 300 Then
                            Tst_C = 299
                            dTav = 0
                        End If

                        dTpumpSum_C = dTpumpSum_C + dTpump_C
                    Else
                        PowfluidWt = 0
                        PowESP_Wt = 0
                        EffESP_d = 0
                    End If

                    N += Nst

                    If saveCurve Then
                        If calc_from_dis Then
                            nn = stage_num - N + 1
                        Else
                            nn = N
                        End If
                        curves.Item("gas_fractionInPump").AddPoint(nn, .F_g)
                        curves.Item("PressureInPump").AddPoint(nn, Pst_atma)
                        curves.Item("TempInPump").AddPoint(nn, Tst_C)
                        curves.Item("PowerfluidInPump").AddPoint(nn, PowfluidTot_Wt)
                        curves.Item("PowerESPInPump").AddPoint(nn, PowESPTot_Wt)
                        curves.Item("EffESPInPump").AddPoint(nn, EffESP_d)
                        curves.Item("q_mix_InPump").AddPoint(nn, q_mix_)

                        curves.Item("mu_stage_cP").AddPoint(nn, .Mu_mix_cP)
                        curves.Item("corr_visc_h_").AddPoint(nn, corr_visc_h_)
                        curves.Item("corr_visc_q_").AddPoint(nn, corr_visc_q_)
                        curves.Item("corr_visc_pow_").AddPoint(nn, corr_visc_pow_)
                        curves.Item("corr_visc_eff_").AddPoint(nn, corr_visc_eff_)
                        curves.Item("gas_corr_").AddPoint(nn, gas_corr_)

                        p_curve_.AddPoint(HmesStage_m(nn), Pst_atma)
                        t_curve_.AddPoint(HmesStage_m(nn), Tst_C)
                    End If
                    i += 1
                Loop
                ' ====================== end main loop ================================================
                If calc_from_dis Then
                    nn = 0
                    curves.Item("gas_fractionInPump").AddPoint(nn, .F_g)
                    curves.Item("PressureInPump").AddPoint(nn, Pst_atma)
                    curves.Item("TempInPump").AddPoint(nn, Tst_C)
                    curves.Item("PowerfluidInPump").AddPoint(nn, PowfluidTot_Wt)
                    curves.Item("PowerESPInPump").AddPoint(nn, PowESPTot_Wt)
                    curves.Item("EffESPInPump").AddPoint(nn, EffESP_d)
                    curves.Item("q_mix_InPump").AddPoint(nn, q_mix_)

                    curves.Item("mu_stage_cP").AddPoint(nn, .Mu_mix_cP)
                    curves.Item("corr_visc_h_").AddPoint(nn, corr_visc_h_)
                    curves.Item("corr_visc_q_").AddPoint(nn, corr_visc_q_)
                    curves.Item("corr_visc_pow_").AddPoint(nn, corr_visc_pow_)
                    curves.Item("corr_visc_eff_").AddPoint(nn, corr_visc_eff_)
                    curves.Item("gas_corr_").AddPoint(nn, gas_corr_)

                    p_curve_.AddPoint(HmesStage_m(CDbl(nn)), Pst_atma)
                    t_curve_.AddPoint(HmesStage_m(CDbl(nn)), Tst_C)

                End If


                If dTpumpSum_C > 298 Then
                    AddLogMsg("Перегрев около УЭЦН, расчетная температура =" & String.Format("{0:0.#}", Tst_C) &
                                  " рост температуры на ступени =" & String.Format("{0:0.##}", dTpump_C) &
                                  " КПД ступени =" & String.Format("{0:0.##}", EffStage) &
                                  " Дебит ступени =" & String.Format("{0:0.##}", .Q_mix_rc_m3day) &
                                  " Температура исправлена на 299")
                End If
            End With
            power_ESP_Wt_ = PowESPTot_Wt + gassep_M_Nm * w_radsec
            power_fluid_Wt_ = PowfluidTot_Wt

            If calc_from_dis Then
                p_int_atma_ = Pst_atma
            Else
                p_dis_atma_ = Pst_atma
                t_dis_C_ = Tst_C
            End If

            eff_ESP_d_ = EffESP_d
            Exit Sub
        Catch ex As Exception
            Dim msg As String
            msg = "Error:CESPpump.ESP_dPIntegration: " & sDELIM

            Throw New ApplicationException(msg)
        End Try

    End Sub

    ' функция расчета деградации из за газа
    Private Function GasCorrection_d(GasFracIn As Double, Optional ByVal gas_degr_value As Double = 0) As Double
        ' переделываем логику расчета поправки на газ для ЭЦН
        ' чтобы можноб было задать руками и выбрать модель тоже
        ' gas_degr_value  - тип насоса по работе с газом:
        '           0-2 значение применяется напрямую;
        '           10 стандартный ЭЦН (предел 25%);
        '           20 ЭЦН с газостабилизирующим модулем (предел 50%);
        '           30 ЭЦН с осевым модулем (предел 75%);
        '           40 ЭЦН с модифицированным ступенями (предел 40%).
        '           110+, тогда модель n-100 применяется ко всем ступеням отдельно
        '           Предел по доле газа на входе в насос после сепарации
        '           на основе статьи SPE 117414 (с корректировкой)
        '           поправка дополнительная к деградации (суммируется).

        Dim b As Double
        b = 0
        If GasFracIn > 0 And GasFracIn < 1 Then
            b = GasFracIn
        End If
        If gas_degr_value >= 0 And gas_degr_value <= 2 Then
            GasCorrection_d = gas_degr_value
        ElseIf gas_degr_value > 9.99 And gas_degr_value < 10.01 Then

            GasCorrection_d = -9 * b ^ 2 + 0.6 * b + 1    ' SPE 117414

        ElseIf gas_degr_value > 19.99 And gas_degr_value < 20.01 Then
            GasCorrection_d = -2 * b ^ 2 + 0.05 * b + 1    ' SPE 117414  corrected rnt
        ElseIf gas_degr_value > 29.99 And gas_degr_value < 30.01 Then
            GasCorrection_d = -1.4 * b ^ 2 + 0.15 * b + 1    ' SPE 117414
        ElseIf gas_degr_value > 39.99 And gas_degr_value < 40.01 Then
            GasCorrection_d = -4 * b ^ 2 + 0.2 * b + 1    ' SPE 117414   corrected rnt
        Else
            GasCorrection_d = 1
        End If
        If GasCorrection_d < 0 Then GasCorrection_d = 0
    End Function

    Public Function get_ESP_power_W(ByVal q_m3day As Double,
                       Optional ByVal stage_num As Integer = -1,
                       Optional ByVal mu_cSt As Double = 1
                               ) As Double
        Dim b As Double
        Dim stage_num_to_calc As Integer
        Dim rate_max As Double

        Call corrections_clear()

        If q_m3day < 0 Then
            get_ESP_power_W = 0
            AddLogMsg("CPumpESP.get_ESP_power_W: расчет характеристики насоса с отрицательным дебитом  Q_m3day = " & q_m3day & "Мощность установлена = 0")
            Exit Function
        End If
        rate_max = rate_max_sm3day(mu_cSt)
        If q_m3day > rate_max Then
            ' assume that for high rate power consumption will not be less that at max rate
            q_m3day = rate_max
        End If
        ' определяем число ступеней с которым будем проводить расчет
        If stage_num > 0 Then        ' если в явном виде задан параметр то его используем
            stage_num_to_calc = stage_num
        Else                        ' иначе использует количество ступеней из характеристики насоса
            stage_num_to_calc = Me.stage_num
        End If
        If correct_visc And (mu_cSt > 0) Then   ' если большая вязкость - сделаем коррекцию
            Call calc_CorrVisc_PetrInst(q_m3day, mu_cSt)   ' метод меняет константы класса, которые влияют на характеристики насоса
        End If
        q_m3day = q_m3day / corr_visc_q_   ' делаем коррекцию по вязкости
        With db_
            b = .freq_Hz / freq_Hz
            get_ESP_power_W = 1000 * b ^ (-3) * stage_num_to_calc * crv_interpolation(.rate_points, .power_points, b * q_m3day, 2)
            If get_ESP_power_W < 0 Then
                get_ESP_power_W = 0
            End If
            get_ESP_power_W = get_ESP_power_W * corr_visc_pow_
        End With
    End Function
    Public Function get_ESP_effeciency_fr(ByVal q_m3day As Double, Optional ByVal mu_cSt As Double = 1) As Double
        Dim b As Double
        'Dim stage_num_to_calc As Integer

        Call corrections_clear()

        If q_m3day < 0 Then
            get_ESP_effeciency_fr = 0
            AddLogMsg("CPumpESP.get_ESP_effeciency_fr: расчет характеристики насоса с отрицательным дебитом  Q_m3day = " & q_m3day & "Мощность установлена = 0")
            Exit Function
        End If
        If q_m3day > rate_max_sm3day(mu_cSt) Then
            get_ESP_effeciency_fr = 0
            Exit Function
        End If
        If correct_visc And (mu_cSt > 0) Then   ' если большая вязкость - сделаем коррекцию
            Call calc_CorrVisc_PetrInst(q_m3day, mu_cSt)   ' метод меняет константы класса, которые влияют на характеристики насоса
        End If
        q_m3day = q_m3day / corr_visc_q_   ' делаем коррекцию по вязкости
        b = db_.freq_Hz / freq_Hz
        get_ESP_effeciency_fr = crv_interpolation(db_.rate_points, db_.eff_points, b * q_m3day, 2)
        If get_ESP_effeciency_fr < 0 Then
            get_ESP_effeciency_fr = 0
        End If
        get_ESP_effeciency_fr = get_ESP_effeciency_fr * corr_visc_eff_
    End Function

    'Public Function array_out(Optional ByVal num_points As Integer = 20)
    '    ' подготовка массива для вывода в Excel
    '    ' num_points - количество точек в выходных массивах для вывода
    '    '
    '    Dim arr(,)
    '    Dim M As Integer
    '    ' Dim FlowParams_out As PIPE_FLOW_PARAMS
    '    Dim offset As Integer
    '    Dim i As Integer
    '    Dim hh As Double
    '    Dim dict_pressure As New Dictionary(Of String, Object)
    '    Dim dict_power As New Dictionary(Of String, Double)
    '    Dim dict_ESP As New Dictionary(Of String, Double)
    '    Dim dict_curves As New Dictionary(Of String, Double)
    '    Dim dict_geometry As New Dictionary(Of String, Double)

    '    offset = 2

    '    Dim crv_P As CInterpolation
    '    Dim crv_T As CInterpolation

    '    ' rearrange output curves one time here - will be used later
    '    crv_P = curves.Item("PressureInPump").ClonePointsToNum(num_points)
    '    crv_T = curves.Item("TempInPump").ClonePointsToNum(num_points)

    '    M = crv_P.Num_points
    '    ReDim arr(M + offset, 10)

    '    ' в первом ряду параметров выведем результаты расчета
    '    ' которые могут пригодится в явном виде при массовых расчетах
    '    arr(0, 0) = 0
    '    arr(1, 0) = "0"

    '    arr(0, 1) = head_m
    '    arr(1, 1) = "head_m"

    '    arr(0, 2) = gassep_M_Nm
    '    arr(1, 2) = "m_Nm"

    '    With dict_geometry
    '        .Add("length_pump_m", length_m)
    '        .Add("angle", angle_deg)
    '        .Add("d_od_m", db.d_od_m)
    '        .Add("d_cas_min_m", db.d_cas_min_m)
    '        .Add("d_motor_od_m", db.d_motor_od_m)
    '    End With

    '    arr(0, 3) = JsonConvert.SerializeObject(dict_geometry)
    '    arr(1, 3) = "geometry"

    '    arr(0, 4) = gas_frac_intake_
    '    arr(1, 4) = "gas_frac_intake"

    '    arr(0, 5) = eff_ESP_d
    '    arr(1, 5) = "eff_ESP_d"

    '    With dict_pressure
    '        .Add("p_int_atma", p_int_atma)
    '        .Add("t_int_C", t_int_C)
    '        .Add("p_dis_atma", p_dis_atma)
    '        .Add("t_dis_C", t_dis_C)
    '    End With
    '    arr(0, 6) = JsonConvert.SerializeObject(dict_pressure)
    '    arr(1, 6) = "dict_pressure"

    '    With dict_power
    '        .Add("power_fluid_W", power_fluid_W)
    '        .Add("power_ESP_W", power_ESP_W)
    '        .Add("p_dis_atma", p_dis_atma)
    '        .Add("eff_ESP_d", eff_ESP_d)
    '    End With
    '    arr(0, 7) = JsonConvert.SerializeObject(dict_power)
    '    arr(1, 7) = "dict_power"

    '    arr(0, 8) = db_json_string
    '    arr(1, 8) = "db_json_string"


    '    With dict_ESP
    '        .Add("calc_from_dis", calc_from_dis_)
    '        .Add("stage_num", stage_num)
    '        .Add("freq_Hz", freq_Hz)
    '        .Add("w_obmin", w_obmin)
    '        .Add("rate_nom_sm3day", rate_nom_sm3day)
    '        .Add("rate_max_sm3day", rate_max_sm3day(-1))
    '        .Add("c_calibr_head", c_calibr_head)
    '        .Add("c_calibr_rate", c_calibr_rate)
    '        .Add("c_calibr_power", c_calibr_power)
    '        .Add("gas_correct", gas_correct)
    '        .Add("gas_corr_", gas_corr_)
    '        .Add("turb_head_factor", turb_head_factor)
    '        .Add("turb_rate_factor", turb_rate_factor)
    '    End With

    '    arr(0, 9) = JsonConvert.SerializeObject(dict_ESP)
    '    arr(1, 9) = "dict_ESP"

    '    With dict_curves
    '        .Add("gas_corr_", curves.Item("gas_corr_").getJson)
    '        .Add("corr_visc_eff_", curves.Item("corr_visc_eff_").getJson)
    '        .Add("corr_visc_pow_", curves.Item("corr_visc_pow_").getJson)
    '        .Add("corr_visc_q_", curves.Item("corr_visc_q_").getJson)
    '        .Add("corr_visc_h_", curves.Item("corr_visc_h_").getJson)
    '        .Add("mu_stage_cP", curves.Item("mu_stage_cP").getJson)
    '        .Add("EffESPInPump", curves.Item("EffESPInPump").getJson)
    '        .Add("PowerESPInPump", curves.Item("PowerESPInPump").getJson)
    '        .Add("PowerfluidInPump", curves.Item("PowerfluidInPump").getJson)
    '        .Add("q_mix_InPump", curves.Item("q_mix_InPump").getJson)
    '        .Add("gas_fractionInPump", curves.Item("gas_fractionInPump").getJson)
    '        .Add("TempInPump", curves.Item("TempInPump").getJson)
    '        .Add("PressureInPump", curves.Item("PressureInPump").getJson)
    '    End With
    '    arr(0, 9) = JsonConvert.SerializeObject(dict_curves)
    '    arr(1, 9) = "dict_curves"

    '    Dim j As Integer

    '    For i = 0 To M

    '        j = 0
    '        hh = crv_P.PointX(i)   ' fractional stage number :)
    '        If i = 0 Then
    '            arr(offset + i, j) = "i"
    '        Else
    '            arr(offset + i, 0) = i
    '        End If

    '        j = j + 1
    '        If i = 0 Then
    '            arr(offset + i, j) = "n_stage"
    '        Else
    '            arr(offset + i, j) = hh
    '        End If

    '        j = j + 1
    '        If i = 0 Then
    '            arr(offset, j) = "length_m"
    '        Else
    '            arr(offset + i, j) = HmesStage_m(hh)
    '        End If

    '        j = j + 1
    '        If i = 0 Then
    '            arr(offset + i, j) = "p_atma"
    '        Else
    '            arr(offset + i, j) = curves.Item("PressureInPump").GetPoint(hh)
    '        End If

    '        j = j + 1
    '        If i = 0 Then
    '            arr(offset + i, j) = "t_C"
    '        Else
    '            arr(offset + i, j) = curves.Item("TempInPump").GetPoint(hh)
    '        End If

    '        j = j + 1
    '        If i = 0 Then
    '            arr(offset + i, j) = "gas_fraction"
    '        Else
    '            arr(offset + i, j) = curves.Item("gas_fractionInPump").GetPoint(hh)
    '        End If

    '        j = j + 1
    '        If i = 0 Then
    '            arr(offset + i, j) = "qmix_rm3day"
    '        Else
    '            arr(offset + i, j) = curves.Item("q_mix_InPump").GetPoint(hh)
    '        End If

    '        j = j + 1
    '        If i = 0 Then
    '            arr(offset, j) = "Power_fluid_W"
    '        Else
    '            arr(offset + i, j) = curves.Item("PowerfluidInPump").GetPoint(hh)
    '        End If

    '        j = j + 1
    '        If i = 0 Then
    '            arr(offset, j) = "Power_ESP_W"
    '        Else
    '            arr(offset + i, j) = curves.Item("PowerESPInPump").GetPoint(hh)
    '        End If

    '        j = j + 1
    '        If i = 0 Then
    '            arr(offset, j) = "eff fluid"
    '        Else
    '            arr(offset + i, j) = curves.Item("EffESPInPump").GetPoint(hh)
    '        End If

    '        j = j + 1
    '        If i = 0 Then
    '            arr(offset, j) = "mu_stage_cP"
    '        Else
    '            arr(offset + i, j) = curves.Item("mu_stage_cP").GetPoint(hh)
    '        End If

    '    Next i

    '    array_out = arr

    'End Function

End Class
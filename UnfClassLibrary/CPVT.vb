'=======================================================================================
'Unifloc 7.24  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2020
'
'=======================================================================================

'class_description_to_manual_eng  - class description - for auto generated manual
' CPVT class describes the properties of reservoir fluids - oil, gas and water
' on the basis of the black oil model, as well as fluid flow parameters - flow rate q_liq and watercut f_w.
' Allows to define all the necessary parameters for the calculations for the given thermobaric conditions,
' such as: gas content in oil, free gas fraction in flow, fluid density and viscosity,
' formation volume factors of fluids and mixtures and others.
' Key function calc_PVT. Its call guarantees recalculation of all flow parameters that can be accessed
' through the appropriate properties.
'description_end_eng

'class_description_to_manual_rus      - для автогенерации описания - помещает комментарии в мануал (со след строки)
' Класс CPVT описывает свойства пластовых флюидов - нефти, газа и воды
' на основе модели нелетучей нефти (black oil), а также параметры потока флюидов - дебита q_liq и обводненности f_w.
' Позволяет для заданных термобарических условий определить все необходимые для проведения расчетов параметры,
' такие как: газосодержание в нефти, долю свободного газа в потоке, плотности и вязкости флюидов,
' объемный расход флюидов и смеси и другие.
' Ключевая функция calc_PVT. Ее вызов гарантирует пересчет всех параметров потока, к которым можно получить доступ
' через соответствующие свойства.
'description_end_rus
Option Explicit On
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq


Public Class CPVT
    Public gas_only As Boolean
    Public ksep_fr As Double
    Public p_ksep_atma As Double
    Public t_ksep_C As Double
    Private _zCorr As Z_CORRELATION
    Private _pVT_correlation As PVT_correlation       ' PVT correlation
    Public gamma_o As Double                        ' плотность нефти удельная
    Public gamma_g As Double                        ' плотность газа удельная
    Public gamma_w As Double                        ' плотность воды удельная
    ' обводненность приватная чтобы гарантировать корректность диапазона
    Private fw_fr_ As Double                        ' объемная доля воды в флюиде
    Public qliq_sm3day As Double                  ' задаем для флюида также и дебиты, это упрощает дальнейшие расчеты расходов в разных условиях
    Public q_gas_free_sm3day As Double              ' дебит газа, вернее добавка для значения Qgas, в исследовательских целях
    ' rsb и rs обладают сложным поведением при задании поэтому приватные
    Private rp_m3m3_ As Double                        ' газовый фактор добычной (приведенный к стандартным условиям)
    Private rsb_m3m3_ As Double                       ' газосодержание при давлении насыщения
    ' калибровочные параметры нефти
    Public pb_atma As Double                        ' давление насыщения  (калибровочное значение)
    Public muob_cP As Double                        ' вязкость нефти при давлении насыщения (калибровочное значение)
    Public bob_m3m3 As Double                       ' объемный коэффициент при давлении насыщения
    Public tres_C As Double                         ' пластовая температура при которой заданые значения давления насыщения и объемного коэффициента

    Private class_name_ As String                     ' имя класса для унифицированной генерации сообщений об ошибках
    Private PT_calc_ As PTtype                        ' термобарические условия при которых был проведен расчет
    ' базовые параметры флюида
    ' расчетные параметры нефти
    Private rsb_calc_m3m3_ As Double                  ' расчетное значение газосодержания при давлении насыщения - может отличаться от исходного если то недопустимо
    Private pb_calc_atma_ As Double                   ' расчетное значение давления насыщения по корреляции
    Private rs_m3m3_ As Double                        ' расчетное значение газосодержания в нефти при текущих условиях
    Private bo_m3m3_ As Double                        ' объемный коэффициент нефти при рабочих условиях
    Private mu_oil_cP_ As Double                      ' вязкость нефти при рабочих условиях
    Private mu_deadoil_cP_ As Double                  ' вязкость дегазированной нефти
    Private copmressibility_o_1atm_ As Double         ' сжимаемость нефти
    Private ST_oilgas_dyncm_ As Double                ' поверхностное натяжение нефть газ
    Private ST_watgas_dyncm_ As Double                ' поверхностное натяжение вода газ
    Private ST_liqgas_dyncm_ As Double
    ' расчетные параметры газа
    Private z_ As Double                              ' расчетное значение коэффициента сверхсжимаемости
    Private bg_m3m3_ As Double                        ' объемный коэффициент газа
    Private mu_gas_cP_ As Double                      ' вязкость газа при рабочих условиях
    ' расчетные параметры воды
    Private bw_m3m3_ As Double                        ' расчетное значение объемного коэффициента воды
    Private bw_sc_m3m3_ As Double
    Private mu_wat_cP_ As Double                      ' вязкость воды
    Private salinity_ppm_ As Double                   ' соленость воды
    ' параметры потока
    Private q_oil_rc_m3day_ As Double
    Private q_wat_rc_m3day_ As Double
    Private q_gas_rc_m3day_ As Double
    Private qliq_rc_m3day_ As Double
    Private gas_fraction_d_ As Double
    Private mu_mix_cP_ As Double
    Private rho_oil_rc_kgm3_ As Double
    Private rho_wat_rc_kgm3_ As Double
    Private rho_liq_rc_kgm3_ As Double
    Private rho_mix_rc_kgm3_ As Double
    ' набор параметров для температурных расчетов
    Private cw_JkgC_ As Double                        ' water heat capacity  теплоемкость воды

    Private heat_capacity_ratio_gas_ As Double
    Private heat_capacity_ratio_oil_ As Double
    Private heat_capacity_ratio_water_ As Double

    Private cv_gas_JkgC_ As Double
    Private cp_oil_JkgC_ As Double

    Public ReadOnly Property Heat_capacity_ratio_gas() As Double
        Get
            Heat_capacity_ratio_gas = heat_capacity_ratio_gas_
        End Get
    End Property

    Public Property Fw_fr() As Double
        Get
            Fw_fr = fw_fr_
        End Get
        Set(val As Double)
            If val >= 0 And val <= 1 Then
                fw_fr_ = val
            ElseIf val < 0 Then
                fw_fr_ = 0
            Else
                fw_fr_ = 1
            End If
        End Set
    End Property

    Public Property Fw_perc() As Double
        Get
            Fw_perc = fw_fr_ * 100
        End Get
        Set(val As Double)
            Fw_fr = val / 100
        End Set
    End Property

    Public ReadOnly Property Fm_gas_fr() As Double
        Get
            Fm_gas_fr = Mg_kgsec() / (Mg_kgsec() + Mo_kgsec() + Mw_kgsec())
        End Get
    End Property

    Public Function Fm_oil_fr() As Double
        Fm_oil_fr = Mo_kgsec() / (Mg_kgsec() + Mo_kgsec() + Mw_kgsec())
    End Function

    Public Function Fm_wat_fr() As Double
        Fm_wat_fr = Mw_kgsec() / (Mg_kgsec() + Mo_kgsec() + Mw_kgsec())
    End Function

    Public Function Polytropic_exponent() As Double
        Polytropic_exponent = (Fm_gas_fr * heat_capacity_ratio_gas_ * cv_gas_JkgC_ + Fm_oil_fr() * Cv_oil_JkgC() + Fm_wat_fr() * Cv_wat_JkgC()) /
                          (Fm_gas_fr * Cv_gas_JkgC() + Fm_oil_fr() * Cv_oil_JkgC() + Fm_wat_fr() * Cv_wat_JkgC())
    End Function

    Private Sub Calc_wc()
        If qliq_sm3day > 0 Then
            Fw_fr = Q_wat_sm3day() / qliq_sm3day
        Else
            Fw_fr = 0
        End If
    End Sub

    Public Function Q_wat_sm3day() As Double
        Q_wat_sm3day = qliq_sm3day * fw_fr_
    End Function

    Public Function Q_gas_sm3day() As Double
        Q_gas_sm3day = Q_oil_sm3day() * Rp_m3m3 + q_gas_free_sm3day    ' учитываем наличие свободного газа в потоке
    End Function

    Public Function Q_gas_insitu_sm3day() As Double
        ' расход газа в заданных термобарических условиях приведенный к стандартным условиям
        Q_gas_insitu_sm3day = (Q_gas_sm3day() - Rs_m3m3() * Q_oil_sm3day())
        If Q_gas_insitu_sm3day < 0 Then
            Q_gas_insitu_sm3day = 0
        End If
        ' при барботаже будет все как надо - свободный газ уже учтен в дебите газа в стандартных условиях
    End Function


    Public Function Q_gas_rc_m3day() As Double
        Q_gas_rc_m3day = Q_gas_insitu_sm3day() * bg_m3m3_
        If Q_gas_rc_m3day < 0 Then Q_gas_rc_m3day = 0
    End Function

    Public Function Q_oil_sm3day() As Double
        Q_oil_sm3day = qliq_sm3day * (1 - fw_fr_)
    End Function

    Public Function Q_oil_rc_m3day() As Double
        Q_oil_rc_m3day = Q_oil_sm3day() * bo_m3m3_
    End Function

    Public Function Q_wat_rc_m3day() As Double
        Q_wat_rc_m3day = Q_wat_sm3day() * bw_m3m3_
    End Function

    Public Function Qliq_rc_m3day() As Double
        Qliq_rc_m3day = Q_wat_rc_m3day() + Q_oil_rc_m3day()
    End Function

    Public Function Q_mix_rc_m3day() As Double
        Q_mix_rc_m3day = q_wat_rc_m3day_ + q_oil_rc_m3day_ + q_gas_rc_m3day_
    End Function

    Public Function Wm_kgsec() As Double
        Wm_kgsec = Mg_kgsec() + Mo_kgsec() + Mw_kgsec()
    End Function

    Public Function Compressibility_oil_1atm() As Double
        Dim t_K As Double
        Dim p_MPa As Double

        t_K = PT_calc_.t_C + const_t_K_min
        p_MPa = PT_calc_.p_atma * const_convert_atma_MPa
        Compressibility_oil_1atm = Unf_pvt_compressibility_oil_VB_1atm(rs_m3m3_, gamma_g, t_K, gamma_o, p_MPa)

    End Function

    Public Function Compressibility_wat_1atm() As Double
        Dim t_K As Double
        Dim p_MPa As Double

        t_K = PT_calc_.t_C + const_t_K_min
        p_MPa = PT_calc_.p_atma * const_convert_atma_MPa
        Compressibility_wat_1atm = Unf_pvt_compressibility_wat_1atma(p_MPa, t_K, salinity_ppm_)
        ' need to check - water compressibility strongly correlate with Bw - water formation volume factor
        ' but here two different correlations are used
        ' in hope that for water everything should be ok
    End Function

    Public Function Compressibility_gas_1atm() As Double
        Dim t_K As Double
        Dim p_MPa As Double

        t_K = PT_calc_.t_C + const_t_K_min
        p_MPa = PT_calc_.p_atma * const_convert_atma_MPa
        Compressibility_gas_1atm = 1 / p_MPa - 1 / z_ * Unf_pvt_dZdp(t_K, p_MPa, gamma_g, Z_CORRELATION.z_Kareem)
        Compressibility_gas_1atm *= const_convert_atma_MPa
    End Function

    Public Function Co_JkgC() As Double  ' oil heat capacity   теплоемкость нефти  Дж/кг/С
        Co_JkgC = cp_oil_JkgC_
    End Function

    Public Function Cw_JkgC() As Double   ' water heat capacity  теплоемкость воды
        Cw_JkgC = cw_JkgC_
    End Function

    Public Function Cg_JkgC() As Double  ' теплоемкость газа gas heat capacity
        Cg_JkgC = Cp_gas_JkgC()
    End Function

    Public Function Cp_gas_JkgC() As Double
        Cp_gas_JkgC = cv_gas_JkgC_ * heat_capacity_ratio_gas_
    End Function

    Public Function Cv_gas_JkgC() As Double
        Cv_gas_JkgC = cv_gas_JkgC_
    End Function

    Public Function Cv_oil_JkgC() As Double
        Cv_oil_JkgC = cp_oil_JkgC_ / heat_capacity_ratio_oil_
    End Function

    Public Function Cp_oil_JkgC() As Double
        Cp_oil_JkgC = cp_oil_JkgC_
    End Function

    Public Function Cv_wat_JkgC() As Double
        Cv_wat_JkgC = cw_JkgC_ / heat_capacity_ratio_water_
    End Function

    Public Function Cp_wat_JkgC() As Double
        Cp_wat_JkgC = cw_JkgC_
    End Function

    Public Function Cliq_JkgC() As Double  ' mixture heat capacity   теплоемкость жидкости  Дж/кг/С
        If Q_mix_rc_m3day() > 0 Then
            Cliq_JkgC = (Co_JkgC() * Mo_kgsec() + Cw_JkgC() * Mw_kgsec()) / (Mw_kgsec() + Mo_kgsec())
        Else
            Cliq_JkgC = Co_JkgC()
        End If
    End Function

    Public Function Cmix_JkgC() As Double  ' mixture heat capacity   теплоемкость жидкости  Дж/кг/С
        If Q_mix_rc_m3day() > 0 Then
            Cmix_JkgC = (Cliq_JkgC() * (Mw_kgsec() + Mo_kgsec()) + Cg_JkgC() * Mg_kgsec()) / (Mo_kgsec() + Mw_kgsec() + Mg_kgsec())
        Else
            Cmix_JkgC = Co_JkgC()
        End If
    End Function

    Public Function CJT_Katm() As Double
        ' коэффциент Джоуля Томсона для многофазной смеси
        Dim x As Double
        Dim wm As Double
        Dim dZdT As Double
        Dim TZdZdT As Double
        wm = (Mo_kgsec() + Mw_kgsec() + Mg_kgsec())
        dZdT = Unf_pvt_dZdt(T_calc_K, P_calc_MPaa, gamma_g, Z_CORRELATION.z_Kareem, z_)
        TZdZdT = T_calc_K() / z_ * dZdT
        If wm > 0 Then
            x = Mg_kgsec() / (Mo_kgsec() + Mw_kgsec() + Mg_kgsec())    ' массовая доля газа в потоке
        Else
            x = 0
        End If
        CJT_Katm = 1 / Cmix_JkgC() * (x / Rho_gas_rc_kgm3() * (-TZdZdT) + (1 - x) / Rho_liq_rc_kgm3()) * const_convert_atma_Pa
    End Function

    Public Function Oil_API() As Double
        Oil_API = 141.5 / gamma_o - 131.5
    End Function

    Public Function Rho_oil_rc_kgm3() As Double
        Dim msg As String
        If bo_m3m3_ > 0 Then
            Rho_oil_rc_kgm3 = 1000 * (gamma_o + rs_m3m3_ * gamma_g * const_rho_air / 1000) / bo_m3m3_
        Else
            ' странная обработка ошибок не при ввода а при расчета - потом надо будет убрать наверное
            msg = "CPVT.rho_oil_rc_kgm3: расчет плотности с неположительным значением Bo_m3m3" & Bo_m3m3() & "Значение Bo проигнорировано"
            AddLogMsg(msg)
            Rho_oil_rc_kgm3 = 1000 * (gamma_o + rs_m3m3_ * gamma_g * const_rho_air / 1000)
        End If
    End Function

    Public Function Rho_wat_rc_kgm3() As Double
        Dim msg As String
        If bw_m3m3_ > 0 Then
            Rho_wat_rc_kgm3 = 1000 * (gamma_w) / bw_m3m3_
        Else
            ' странная обработка ошибок не при ввода а при расчета - потом надо будет убрать наверное
            msg = "CPVT.rho_wat_rc_kgm3: расчет плотности с неположительным значением Bw_m3m3" & bw_m3m3_ & "Значение Bw проигнорировано"
            AddLogMsg(msg)
            Rho_wat_rc_kgm3 = 1000 * (gamma_w)
        End If
    End Function

    Public Function Rho_liq_rc_kgm3() As Double
        Rho_liq_rc_kgm3 = rho_liq_rc_kgm3_ '(1 - fw_fr) * rho_oil_rc_kgm3 + fw_fr * rho_wat_rc_kgm3
    End Function

    Public Function Rho_gas_rc_kgm3() As Double
        Dim msg As String
        If bg_m3m3_ > 0 Then
            Rho_gas_rc_kgm3 = gamma_g * const_rho_air / bg_m3m3_
        Else
            ' странная обработка ошибок не при ввода а при расчета - потом надо будет убрать наверное
            msg = "CPVT.rho_gas_rc_kgm3: расчет плотности с неположительным значением Bg_m3m3" & Bg_m3m3() & "Значение Bg проигнорировано"
            AddLogMsg(msg)
            Rho_gas_rc_kgm3 = gamma_g * const_rho_air
        End If
    End Function

    Public Function F_g() As Double
        If Q_mix_rc_m3day() > 0 Then
            F_g = Q_gas_rc_m3day() / Q_mix_rc_m3day()
        Else
            F_g = 0
        End If
    End Function

    Public Function Rho_mix_rc_kgm3() As Double
        Rho_mix_rc_kgm3 = rho_mix_rc_kgm3_  ' rho_liq_rc_kgm3 * (1 - f_g) + rho_gas_rc_kgm3 * f_g
    End Function

    Public Function Sigma_liq_Nm() As Double
        Sigma_liq_Nm = ST_liqgas_dyncm_ * 0.001
    End Function

    Public Function Sigma_oil_gas_Nm() As Double
        Sigma_oil_gas_Nm = ST_oilgas_dyncm_ * 0.001
    End Function

    Public Function Sigma_wat_gas_Nm() As Double
        'sigma_wat_gas_Nm = p_sigma_wat_gas_Nm
        Sigma_wat_gas_Nm = ST_watgas_dyncm_ * 0.001
    End Function

    <CodeAnalysis.SuppressMessage("Style", "IDE1006:Naming Styles", Justification:="<Pending>")>
    Public Function T_res_K() As Double
        T_res_K = tres_C + const_t_K_min
    End Function

    ' молярная масса газа   (используется например в штуцере)
    Public Function Mg_kgmol() As Double
        Mg_kgmol = const_m_a_kgmol * gamma_g
    End Function

    Public Function Sal_ppm() As Double
        Sal_ppm = salinity_ppm_
    End Function

    Public Function Rho_oil_sckgm3() As Double
        Rho_oil_sckgm3 = gamma_o * const_rho_ref
    End Function

    Public Function Rho_gas_sckgm3() As Double
        Rho_gas_sckgm3 = gamma_g * const_rho_air
    End Function

    Public Function Rho_wat_sckgm3() As Double
        Rho_wat_sckgm3 = gamma_w * const_rho_ref
    End Function

    Public Function Rp_full_m3m3() As Double
        If Q_oil_sm3day() > 0 Then
            Rp_full_m3m3 = rp_m3m3_ + q_gas_free_sm3day / Q_oil_sm3day()
        Else
            Rp_full_m3m3 = rp_m3m3_
        End If
    End Function

    ' ----- Rp - GOR  ----------------------------------------------------------------------------------------
    Public Property Rp_m3m3() As Double
        Get
            Rp_m3m3 = rp_m3m3_

        End Get
        Set(val As Double)
            If (val >= 0) Then
                rp_m3m3_ = val
                If rp_m3m3_ < rsb_m3m3_ Then   ' проверим, что газовый фактор должен быть больше чем газосодержание
                    'addLogMsg "Газовый фактор при вводе больше газосодержания Rp = " & Format(rp_m3m3_, "####0.00") & " < rsb = " & Format(rsb_m3m3_, "#0.00") & ". Газосодержание исправлено"
                    rsb_calc_m3m3_ = rp_m3m3_
                End If
            Else
                ' унифицированная реакция на ошибочный ввод ключевых параметров класса
                Dim msg As String, fname As String
                fname = "rp_m3m3"
                msg = class_name_ & "." & fname & ": input - wrong " & fname & " = " & CStr(val)
                AddLogMsg(msg)
                Throw New ApplicationException(msg)
                ' Err.Raise kErrPVTinput, class_name_ & "." & fname, msg
            End If

        End Set
    End Property


    ' ----- rsb -----------------------------------------------------------------------------------------
    Public Property Rsb_m3m3() As Double
        Get
            Rsb_m3m3 = rsb_m3m3_
        End Get
        Set(val As Double)
            If (val >= 0) Then
                rsb_m3m3_ = val
                rsb_calc_m3m3_ = val
                If rp_m3m3_ < rsb_m3m3_ Then   ' проверим, что газовый фактор должен быть больше чем газосодержание
                    'addLogMsg "газосодержания при вводе меньше газового фактора  Rp = " & Format(rp_m3m3_, "#0.00") & " < rsb = " & Format(rsb_m3m3_, "#0.00") & ". Газосодержание исправлено"
                    rsb_calc_m3m3_ = rp_m3m3_
                End If
            Else
                ' унифицированная реакция на ошибочный ввод ключевых параметров класса
                Dim msg As String, fname As String
                fname = "rsb_m3m3"
                msg = class_name_ & "." & fname & ": input - wrong " & fname & " = " & CStr(val)
                AddLogMsg(msg)
                Throw New ApplicationException(msg)
            End If
        End Set
    End Property


    Public Function Rsb_calc_m3m3() As Double
        Rsb_calc_m3m3 = rsb_calc_m3m3_
    End Function


    Public Function Set_rp_rsb(ByVal Rpval_m3m3 As Double, ByVal Rsbval_m3m3 As Double) As Boolean
        ' безопасный с точки зрения начисления штрафов способ установки произвольных значений газового фактора в системе
        If Rpval_m3m3 > 0 Then
            If Rpval_m3m3 >= Rsbval_m3m3 Then
                rp_m3m3_ = Rpval_m3m3
                If Rsbval_m3m3 > 0 Then
                    rsb_m3m3_ = Rsbval_m3m3
                Else
                    rsb_m3m3_ = Rpval_m3m3
                End If
                rsb_calc_m3m3_ = rsb_m3m3_
                Set_rp_rsb = True
            Else
                'addLogMsg "CPVT.set_rp_rsb: Газосодержание при вводе больше газового фактора  Rp = " & Format(Rpval_m3m3, "#0.00") & " < rsb = " & Format(Rsbval_m3m3, "#0.00") & ". Газосодержание исправлено"
                rp_m3m3_ = Rpval_m3m3
                rsb_calc_m3m3_ = Rpval_m3m3
                rsb_m3m3_ = Rsbval_m3m3
                Set_rp_rsb = True
            End If
        Else
            If Rpval_m3m3 <= 0 And Rsbval_m3m3 > 0 Then
                rp_m3m3_ = Rsbval_m3m3
                rsb_m3m3_ = Rsbval_m3m3
                rsb_calc_m3m3_ = rsb_m3m3_
                Set_rp_rsb = True
            Else
                Set_rp_rsb = False
            End If
        End If
        ' устновим все значения для зависимых флюидов
        Rp_m3m3 = Rp_m3m3
        Rsb_m3m3 = Rsb_m3m3
    End Function

    '----- Pb -----------------------------------------------------------------------------------------
    Public Function Pb_calc_atma() As Double
        ' функция выдает давление насыщения, которое было получено в ходе расчетов
        ' может отличатся от того, что было задано при инициализации, если оно было не допустимо
        ' если не было задано - то оно рассчитывается и выдается расчитанное
        If pb_calc_atma_ > 0 Then       ' ноль не допустим, это значит что значение отсутствует
            ' если известно калибровочное значение при пластовой температуре, то возвращаем его
            Pb_calc_atma = pb_calc_atma_
        Else
            ' иначе считаем что получится из расчета по корреляции по газосодержанию
            Pb_calc_atma = Calc_pb_atma(Rsb_m3m3, tres_C)
        End If
    End Function

    Public Function Rs_m3m3() As Double
        Rs_m3m3 = rs_m3m3_
    End Function

    Public Function Bo_m3m3() As Double
        Bo_m3m3 = bo_m3m3_
    End Function

    Public Function Bg_m3m3() As Double
        Bg_m3m3 = bg_m3m3_
    End Function

    Public Function Bw_m3m3() As Double
        Bw_m3m3 = bw_m3m3_
    End Function

    Public Function Mu_oil_cP() As Double
        Mu_oil_cP = mu_oil_cP_
    End Function

    Public Function Mu_wat_cP() As Double
        Mu_wat_cP = mu_wat_cP_
    End Function

    Public Function Mu_gas_cP() As Double
        Mu_gas_cP = mu_gas_cP_
    End Function

    Public Function Mu_liq_cP() As Double
        '
        ' todo надо уточнить как считать вязкость для смеси - быть может надо холдап использовать
        '
        Dim fw_rc_fr As Double
        If Qliq_rc_m3day() > 0 Then
            fw_rc_fr = Q_wat_rc_m3day() / Qliq_rc_m3day()
        Else
            fw_rc_fr = fw_fr_
        End If

        Mu_liq_cP = (Mu_oil_cP() * (1 - fw_rc_fr) +
                Mu_wat_cP() * fw_rc_fr)
    End Function

    Public Function Mu_mix_cP() As Double
        Mu_mix_cP = mu_mix_cP_
    End Function

    ' кинематическая вязкость смеси в сантистоксах
    Public Function Mu_mix_cSt() As Double
        Mu_mix_cSt = mu_mix_cP_ / (rho_mix_rc_kgm3_ / 1000)
    End Function

    ' массовый расход нефти
    Public Function Mo_kgsec() As Double
        Mo_kgsec = Q_oil_rc_m3day() * Rho_oil_rc_kgm3() / const_conver_day_sec
    End Function
    ' массовый расход воды
    Public Function Mw_kgsec() As Double
        Mw_kgsec = Q_wat_rc_m3day() * Rho_wat_rc_kgm3() / const_conver_day_sec
    End Function
    ' массовый расход газа
    Public Function Mg_kgsec() As Double
        Mg_kgsec = Q_gas_rc_m3day() * Rho_gas_rc_kgm3() / const_conver_day_sec
    End Function

    Public Function Z() As Double
        Z = z_
    End Function

    Public Function P_calc_atma() As Double
        P_calc_atma = PT_calc_.p_atma
    End Function

    Public Function P_calc_MPaa() As Double
        P_calc_MPaa = P_calc_atma() * const_convert_atma_MPa
    End Function

    <CodeAnalysis.SuppressMessage("Style", "IDE1006:Naming Styles", Justification:="<Pending>")>
    Public Function T_calc_C() As Double
        T_calc_C = PT_calc_.t_C
    End Function

    <CodeAnalysis.SuppressMessage("Style", "IDE1006:Naming Styles", Justification:="<Pending>")>
    Public Function T_calc_K() As Double
        T_calc_K = T_calc_C() + const_t_K_min
    End Function

    <CodeAnalysis.SuppressMessage("Style", "IDE1006:Naming Styles", Justification:="<Pending>")>
    Public Function T_calc_F() As Double
        T_calc_F = T_calc_C() * 1.8 + 32
    End Function

    Public Function Mu_deadoil_cP() As Double
        Mu_deadoil_cP = mu_deadoil_cP_
    End Function


    Friend Property ZCorr As Z_CORRELATION
        Get
            Return _zCorr
        End Get
        Set(value As Z_CORRELATION)
            _zCorr = value
        End Set
    End Property

    Friend Property PVT_correlation As PVT_correlation
        Get
            Return _pVT_correlation
        End Get
        Set(value As PVT_correlation)
            _pVT_correlation = value
        End Set
    End Property



    '===================================================================================
    ' функции и процедуры
    '===================================================================================

    Public Sub Class_Initialize(Optional ByVal class_name As String = "CPVT",
                                Optional ByVal gamma_gas As Double = 0.6,
                                Optional ByVal gamma_oil As Double = 0.86,
                                Optional ByVal gamma_wat As Double = 1,
                                Optional ByVal rsb_m3m3_ As Double = 100,
                                Optional ByVal pb_atma_ As Double = -1,
                                Optional ByVal bob_m3m3_ As Double = -1,
                                Optional ByVal PVTcorr As Integer = PVT_correlation.Standing_based,
                                Optional ByVal tres_C_ As Double = 90,
                                Optional ByVal rp_m3m3_ As Double = 100,
                                Optional ByVal Fw_perc_ As Double = 0,
                                Optional ByVal qliq_sm3day_ As Double = 100,
                                Optional ByVal q_gas_free_sm3day_ As Double = 0,
                                Optional ByVal cw_JkgC As Double = 4176,
                                Optional ByVal ZCorr_ As Integer = Z_CORRELATION.z_Kareem,
                                Optional ByVal gas_only_ As Boolean = False,
                                Optional ByVal pb_calc_atma As Double = 0,
                                Optional ByVal heat_capacity_ratio_gas As Double = 1.3,
                                Optional ByVal heat_capacity_ratio_oil As Double = 1.05,
                                Optional ByVal heat_capacity_ratio_water As Double = 1,
                                Optional ByVal ksep_fr_ As Double = 0,
                                Optional ByVal p_ksep_atma_ As Double = 0,
                                Optional ByVal t_ksep_C_ As Double = 0,
                                Optional ByVal muob_cP_ As Double = 0,
                                Optional ByVal rsb_calc_m3m3 As Double = 0,
                                Optional ByVal rs_m3m3 As Double = 0,
                                Optional ByVal bo_m3m3 As Double = 0,
                                Optional ByVal mu_oil_cP As Double = 0,
                                Optional ByVal mu_deadoil_cP As Double = 0,
                                Optional ByVal copmressibility_o_1atm As Double = 0,
                                Optional ByVal ST_oilgas_dyncm As Double = 0,
                                Optional ByVal ST_watgas_dyncm As Double = 0,
                                Optional ByVal ST_liqgas_dyncm As Double = 0,
                                Optional ByVal z As Double = 0,
                                Optional ByVal bg_m3m3 As Double = 0,
                                Optional ByVal mu_gas_cP As Double = 0,
                                Optional ByVal bw_m3m3 As Double = 0,
                                Optional ByVal bw_sc_m3m3 As Double = 0,
                                Optional ByVal mu_wat_cP As Double = 0,
                                Optional ByVal salinity_ppm As Double = 0,
                                Optional ByVal q_oil_rc_m3day As Double = 0,
                                Optional ByVal q_wat_rc_m3day As Double = 0,
                                Optional ByVal q_gas_rc_m3day As Double = 0,
                                Optional ByVal qliq_rc_m3day As Double = 0,
                                Optional ByVal gas_fraction_d As Double = 0,
                                Optional ByVal mu_mix_cP As Double = 0,
                                Optional ByVal rho_oil_rc_kgm3 As Double = 0,
                                Optional ByVal rho_wat_rc_kgm3 As Double = 0,
                                Optional ByVal rho_liq_rc_kgm3 As Double = 0,
                                Optional ByVal rho_mix_rc_kgm3 As Double = 0,
                                Optional ByVal cv_gas_JkgC As Double = 0,
                                Optional ByVal cp_oil_JkgC As Double = 0)

        'Optional ByVal PT_calc_ As PTtype, без понятия как объявить и стоит ли

        class_name_ = class_name
        PVT_correlation = PVTcorr
        gamma_o = gamma_oil
        gamma_g = gamma_gas
        gamma_w = gamma_wat
        Rp_m3m3 = rp_m3m3_
        Rsb_m3m3 = rsb_m3m3_
        pb_atma = pb_atma_  ' по умолчанию нет калибровок, только корреляция
        bob_m3m3 = bob_m3m3_ ' по умолчанию нет калибровок, только корреляция
        tres_C = tres_C_
        Fw_perc = Fw_perc_
        qliq_sm3day = qliq_sm3day_
        q_gas_free_sm3day = q_gas_free_sm3day_
        ' для начала для простоты инициализируем теплоемкость флюидов как константы
        ' потом можно будет добавить расчет в зависимости от условий
        cw_JkgC_ = cw_JkgC
        ZCorr = ZCorr_
        gas_only = gas_only_
        pb_calc_atma_ = pb_calc_atma
        ' heat capacity ratios estimated from perticular multiflash calcs
        ' good idea to improove adding correlations
        heat_capacity_ratio_gas_ = heat_capacity_ratio_gas
        heat_capacity_ratio_oil_ = heat_capacity_ratio_oil
        heat_capacity_ratio_water_ = heat_capacity_ratio_water

        ksep_fr = ksep_fr_
        p_ksep_atma = p_ksep_atma_
        t_ksep_C = t_ksep_C_
        muob_cP = muob_cP_
        rsb_calc_m3m3_ = rsb_calc_m3m3
        rs_m3m3_ = rs_m3m3
        bo_m3m3_ = bo_m3m3
        mu_oil_cP_ = mu_oil_cP
        mu_deadoil_cP_ = mu_deadoil_cP
        copmressibility_o_1atm_ = copmressibility_o_1atm
        ST_oilgas_dyncm_ = ST_oilgas_dyncm
        ST_watgas_dyncm_ = ST_watgas_dyncm
        ST_liqgas_dyncm_ = ST_liqgas_dyncm
        z_ = z
        bg_m3m3_ = bg_m3m3
        mu_gas_cP_ = mu_gas_cP
        bw_m3m3_ = bw_m3m3
        bw_sc_m3m3_ = bw_sc_m3m3
        mu_wat_cP_ = mu_wat_cP
        salinity_ppm_ = salinity_ppm
        q_oil_rc_m3day_ = q_oil_rc_m3day
        q_wat_rc_m3day_ = q_wat_rc_m3day
        q_gas_rc_m3day_ = q_gas_rc_m3day
        qliq_rc_m3day_ = qliq_rc_m3day
        gas_fraction_d_ = gas_fraction_d
        mu_mix_cP_ = mu_mix_cP
        rho_oil_rc_kgm3_ = rho_oil_rc_kgm3
        rho_wat_rc_kgm3_ = rho_wat_rc_kgm3
        rho_liq_rc_kgm3_ = rho_liq_rc_kgm3
        rho_mix_rc_kgm3_ = rho_mix_rc_kgm3
        cv_gas_JkgC_ = cv_gas_JkgC
        cp_oil_JkgC_ = cp_oil_JkgC
    End Sub


    Public Sub Init(Optional ByVal gamma_gas As Double = 0.6,
                    Optional ByVal gamma_oil As Double = 0.86,
                    Optional ByVal gamma_wat As Double = 1,
                    Optional ByVal rsb_m3m3 As Double = 100,
                    Optional ByVal pb_atma As Double = -1,
                    Optional ByVal bob_m3m3 As Double = -1,
                    Optional ByVal PVTcorr As Integer = PVT_correlation.Standing_based,
                    Optional ByVal tres_C As Double = 90,
                    Optional ByVal rp_m3m3 As Double = -1,
                    Optional ByVal muob_cP As Double = -1)
        gamma_g = gamma_gas
        gamma_o = gamma_oil
        gamma_w = gamma_wat
        Set_rp_rsb(rp_m3m3, rsb_m3m3)
        Me.pb_atma = pb_atma
        If tres_C > 0 Then Me.tres_C = tres_C
        Me.bob_m3m3 = bob_m3m3
        Me.muob_cP = muob_cP
        PVT_correlation = CType(PVTcorr, PVT_correlation)
    End Sub

    Public Function Clone() As CPVT
        Dim fl As New CPVT
        Call fl.Copy(Me)
        Clone = fl
    End Function

    Public Sub Copy(fl As CPVT)
        ' all params that define fluid must be copied here
        PVT_correlation = fl.PVT_correlation
        gamma_o = fl.gamma_o
        gamma_g = fl.gamma_g
        gamma_w = fl.gamma_w
        Set_rp_rsb(fl.Rp_m3m3, fl.Rsb_m3m3)
        pb_atma = fl.pb_atma
        bob_m3m3 = fl.bob_m3m3
        muob_cP = fl.muob_cP
        tres_C = fl.tres_C
        Fw_fr = fl.Fw_fr
        qliq_sm3day = fl.qliq_sm3day
        q_gas_free_sm3day = fl.q_gas_free_sm3day
        gas_only = fl.gas_only

    End Sub


    Public Sub Calc_PVT(ByVal p_atma As Double, ByVal t_C As Double)
        ' расчет свойств воды нефти и газа при заданных давлении и температуре
        Dim t_K As Double  ' internal K temp
        'PVT properties
        Dim rho_o As Double
        Dim bob_m3m3_sat As Double
        'internal buffers used to store output values
        Dim p_bi As Double
        Dim r_si As Double
        Dim rho_o_sat As Double
        Dim p_fact As Double
        Dim p_offs As Double
        Dim b_fact As Double
        Dim mu_fact As Double
        'Oil pressure in MPa
        Dim p_MPa As Double
        Dim Pb_calbr_MPa As Double
        Dim rsb_calbr_m3m3 As Double
        Dim Bo_calbr_m3m3 As Double
        Dim muo_calibr_cP As Double
        Dim ranges_good As Boolean
        Dim mu_deadoil_cP As Double
        Dim Muo_saturated_cP As Double
        Try
            t_K = t_C + const_t_K_min
            PT_calc_.p_atma = p_atma
            PT_calc_.t_C = t_C
            Call Set_rp_rsb(rp_m3m3_, rsb_m3m3_)  ' init calc variable and eliminate previous calc influence
            rsb_calbr_m3m3 = rsb_calc_m3m3_
            Bo_calbr_m3m3 = bob_m3m3
            muo_calibr_cP = muob_cP

            p_MPa = p_atma * const_convert_atma_MPa

            Pb_calbr_MPa = pb_atma * const_convert_atma_MPa 'convert user specified bubblepoint pressure
            'for saturated oil calibration is applied by application of factor p_fact to input pressure
            'for undersaturated - by shifting according to p_offs
            'calculate PVT properties
            'calculate water properties at current pressure and temperature
            bw_sc_m3m3_ = gamma_w
            heat_capacity_ratio_gas_ = Unf_pvt_gas_heat_capacity_ratio(gamma_g, t_K)

            salinity_ppm_ = Unf_pvt_Sal_BwSC_ppm(bw_sc_m3m3_)
            bw_m3m3_ = Unf_pvt_Bw_m3m3(p_MPa, t_K) '* bw_sc_m3m3_
            mu_wat_cP_ = Unf_pvt_viscosity_wat_cP(p_MPa, t_K, salinity_ppm_)
            'if no calibration gas-oil ratio specified, then set it to some very large value and
            'switch of calibration for bubblepoint and oil formation volume factor
            z_ = Unf_pvt_Zgas_d(t_K, p_MPa, gamma_g, ZCorr)
            bg_m3m3_ = Unf_pvt_Bg_z_m3m3(t_K, p_MPa, z_)
            mu_gas_cP_ = Unf_pvt_viscosity_gas_cP(t_K, p_MPa, z_, gamma_g)
            If PVT_correlation = PVT_correlation.Standing_based Then
                mu_deadoil_cP = Unf_pvt_viscosity_dead_oil_Beggs_Robinson_cP(t_K, gamma_o) 'dead oil viscosity
                Muo_saturated_cP = Unf_pvt_viscosity_saturated_oil_Beggs_Robinson_cP(rsb_calbr_m3m3, mu_deadoil_cP) 'saturated oil viscosity Beggs & Robinson
                p_bi = Unf_pvt_pb_Standing_MPa(rsb_calbr_m3m3, gamma_g, T_res_K, gamma_o)    ' считаем давление насыщения по корреляции Standing для пластовой температуры при которой заданны калибровочные значения
                ' дальше ищем калибровочные коэффициенты
                'Calculate bubble point correction factor
                If (Pb_calbr_MPa > 0) Then 'user specified
                    p_fact = p_bi / Pb_calbr_MPa
                Else ' not specified, use from correlations
                    p_fact = 1
                End If
                If (Bo_calbr_m3m3 > 0) Then 'Calculate oil formation volume factor correction factor
                    bob_m3m3_sat = Unf_pvt_FVF_Saturated_Oil_Standing_m3m3(rsb_calbr_m3m3, gamma_g, T_res_K, gamma_o)  ' значение по корреляции считаем также для пластовой температуры
                    b_fact = (Bo_calbr_m3m3 - 1) / (bob_m3m3_sat - 1)
                Else ' not specified, use from correlations
                    b_fact = 1
                End If
                If muo_calibr_cP > 0 Then           ' рассчитаем калибровочный коэффициент для вязкости при давлении насыщения
                    mu_fact = muo_calibr_cP / Muo_saturated_cP
                Else
                    mu_fact = 1
                End If
                p_bi = Unf_pvt_pb_Standing_MPa(rsb_calbr_m3m3, gamma_g, t_K, gamma_o)   ' давление насыщения по корреляции при текущей температуре
                p_MPa *= p_fact   ' растянем давление чтобы натянуть его на калиброванное значение
                If p_MPa > p_bi Then 'apply correction to undersaturated oil 'undersaturated oil
                    r_si = rsb_calbr_m3m3   ' результат такое будет
                    bob_m3m3_sat = b_fact * (Unf_pvt_FVF_Saturated_Oil_Standing_m3m3(rsb_calbr_m3m3, gamma_g, t_K, gamma_o) - 1) + 1 ' it is assumed that at pressure 1 atma bo=1
                    copmressibility_o_1atm_ = Unf_pvt_compressibility_oil_VB_1atm(rsb_calbr_m3m3, gamma_g, t_K, gamma_o, p_MPa) 'calculate compressibility at bubble point pressure
                    bo_m3m3_ = bob_m3m3_sat * Math.Exp(copmressibility_o_1atm_ * (p_bi - p_MPa))
                    mu_oil_cP_ = mu_fact * Unf_pvt_viscosity_oil_Vasquez_Beggs_cP(Muo_saturated_cP, p_MPa, p_bi)  'Vesquez&Beggs
                Else 'apply correction to saturated oil
                    r_si = Unf_pvt_GOR_Standing_m3m3(p_MPa, gamma_g, t_K, gamma_o)
                    bo_m3m3_ = b_fact * (Unf_pvt_FVF_Saturated_Oil_Standing_m3m3(r_si, gamma_g, t_K, gamma_o) - 1) + 1 ''Standing. it is assumed that at pressure 1 atma bo=1
                    mu_oil_cP_ = mu_fact * Unf_pvt_viscosity_saturated_oil_Beggs_Robinson_cP(r_si, mu_deadoil_cP)  'Beggs & Robinson
                End If
            End If
            If PVT_correlation = PVT_correlation.McCain_based Then
                ranges_good = True
                ranges_good = ranges_good And CheckRanges(t_K, "t_K", const_tMcCain_K_min, const_t_K_max, "температура потока вне диапазона для корреляции маккейна", "calc_PVT (McCain)", True)
                mu_deadoil_cP = Unf_pvt_viscosity_dead_oil_Standing_cP(t_K, gamma_o)  'dead oil viscosity
                Muo_saturated_cP = Unf_pvt_viscosity_saturated_oil_Beggs_Robinson_cP(rsb_calbr_m3m3, mu_deadoil_cP) 'saturated oil viscosity Beggs & Robinson
                p_bi = Unf_pvt_pb_Valko_McCain_MPa(rsb_calbr_m3m3, gamma_g, T_res_K, gamma_o)
                'Calculate bubble point correction factor
                If (Pb_calbr_MPa > 0) Then 'user specifie
                    p_fact = p_bi / Pb_calbr_MPa
                Else ' not specified, use from correlations
                    p_fact = 1
                End If
                p_MPa *= p_fact
                copmressibility_o_1atm_ = Unf_pvt_compressibility_oil_VB_1atm(rsb_calbr_m3m3, gamma_g, T_res_K, gamma_o, p_bi) 'calculate compressibility at bubble point pressure
                If (Bo_calbr_m3m3 > 0) Then 'Calculate oil formation volume factor correction factor
                    rho_o_sat = Unf_pvt_density_McCain_kgm3(p_bi, gamma_g, T_res_K, gamma_o, rsb_calbr_m3m3, p_bi, copmressibility_o_1atm_)    ' тут формально есть зависимость от сжимаемости но реально она не влияет (так как расчет идет при давлении насыщения)
                    bob_m3m3_sat = Unf_pvt_FVF_McCain_m3m3(rsb_calbr_m3m3, gamma_g, gamma_o * const_rho_ref, rho_o_sat)
                    b_fact = (Bo_calbr_m3m3 - 1) / (bob_m3m3_sat - 1)
                Else ' not specified, use from correlations
                    b_fact = 1
                End If
                p_bi = Unf_pvt_pb_Valko_McCain_MPa(rsb_calbr_m3m3, gamma_g, t_K, gamma_o)
                r_si = Unf_pvt_GOR_Velarde_m3m3(p_MPa, p_bi, gamma_g, t_K, gamma_o, rsb_calbr_m3m3)
                If p_MPa > p_bi Then 'apply correction to undersaturated oil
                    copmressibility_o_1atm_ = Unf_pvt_compressibility_oil_VB_1atm(rsb_calbr_m3m3, gamma_g, t_K, gamma_o, p_MPa)  'calculate compressibility at bubble point pressure
                    rho_o_sat = Unf_pvt_density_McCain_kgm3(p_bi, gamma_g, t_K, gamma_o, rsb_calbr_m3m3, p_bi, copmressibility_o_1atm_)
                    bob_m3m3_sat = Unf_pvt_FVF_McCain_m3m3(rsb_calbr_m3m3, gamma_g, gamma_o * const_rho_ref, rho_o_sat)
                    bob_m3m3_sat = b_fact * (bob_m3m3_sat - 1) + 1 ' it is assumed that at pressure 1 atma bo=1
                    bo_m3m3_ = bob_m3m3_sat * Math.Exp(copmressibility_o_1atm_ * (p_bi - p_MPa))
                Else 'apply correction to saturated oil
                    rho_o = Unf_pvt_density_McCain_kgm3(p_MPa, gamma_g, t_K, gamma_o, r_si, p_bi, copmressibility_o_1atm_)
                    bo_m3m3_ = b_fact * (Unf_pvt_FVF_McCain_m3m3(r_si, gamma_g, gamma_o * const_rho_ref, rho_o) - 1) + 1 ' it is assumed that at pressure 1 atma bo=1
                End If
                If muo_calibr_cP > 0 Then           ' рассчитаем калибровочный коэффициент для вязкости при давлении насыщения
                    If (rsb_calbr_m3m3 < 350) Then
                        mu_fact = muo_calibr_cP / Unf_pvt_viscosity_oil_Standing_cP(rsb_calbr_m3m3, mu_deadoil_cP, p_bi, p_bi)
                    Else
                        mu_fact = muo_calibr_cP / Muo_saturated_cP
                    End If
                Else
                    mu_fact = 1
                End If
                If (rsb_calbr_m3m3 < 350) Then 'Calculate oil viscosity acoording to Standing
                    mu_oil_cP_ = mu_fact * Unf_pvt_viscosity_oil_Standing_cP(r_si, mu_deadoil_cP, p_MPa, p_bi)
                Else 'Calculate according to Begs&Robinson (saturated) and Vasquez&Begs (undersaturated)
                    If p_MPa > p_bi Then 'undersaturated oil
                        mu_oil_cP_ = mu_fact * Unf_pvt_viscosity_oil_Vasquez_Beggs_cP(Muo_saturated_cP, p_MPa, p_bi)
                    Else 'saturated oil
                        'Beggs & Robinson
                        mu_oil_cP_ = mu_fact * Unf_pvt_viscosity_saturated_oil_Beggs_Robinson_cP(r_si, mu_deadoil_cP)
                    End If
                End If
            End If
            If PVT_correlation = 2 Then  'Debug mode. Linear Rs and bo vs P, Pb_calbr_atma should be specified.
                'gas properties
                z_ = 0.95 'ideal gas
                bg_m3m3_ = Unf_pvt_Bg_z_m3m3(t_K, p_MPa, z_)
                mu_gas_cP_ = 0.0000000001
                p_fact = 1         'Set to default. b_rb should be specified by user!
                p_offs = 0
                If Pb_calbr_MPa <= 0 Then
                    ' can not be estimated without calibration properties
                    Throw New ApplicationException("CPVT.calc_PVT" + " PVT correlation = 2 without Pb input not allowed")
                    'Err.Raise kErrPVTinput, "CPVT.calc_PVT", "PVT correlation = 2 without Pb input not allowed"
                End If
                p_bi = Pb_calbr_MPa
                If p_MPa > (p_bi) Then 'undersaturated oil
                    r_si = rsb_calbr_m3m3
                Else 'saturate
                    r_si = p_MPa / Pb_calbr_MPa * rsb_calbr_m3m3
                End If
                'if bob_m3m3 is not specified by the user then
                'set bob_m3m3 so, that oil density, recalculated with Rs_m3m3 would be equal to dead oil density
                If (Bo_calbr_m3m3 <= 0) Then
                    AddLogMsg("warning:CPVT.calc_PVT, PVT correlation = 2 without Bob input")
                    bo_m3m3_ = (1 + r_si * (gamma_g * const_rho_air) / (gamma_o * const_rho_ref))
                Else
                    If p_MPa > (p_bi) Then 'undersaturated oil
                        bo_m3m3_ = Bo_calbr_m3m3
                    Else 'saturate
                        bo_m3m3_ = 1 + (Bo_calbr_m3m3 - 1) * ((p_MPa - const_convert_atma_MPa) / (p_bi - const_convert_atma_MPa))
                    End If
                End If
                If muo_calibr_cP >= 0 Then
                    mu_oil_cP_ = muo_calibr_cP
                Else
                    AddLogMsg("warning:CPVT.calc_PVT, PVT correlation = 2 without Muob input")
                    mu_oil_cP_ = 1
                End If
            End If
            'Assign output variables
            pb_calc_atma_ = p_bi / p_fact / const_convert_atma_MPa
            ' pb_atma_ = pb_calc_atma_   ' corrected by issue #34
            rs_m3m3_ = r_si
            mu_deadoil_cP_ = mu_deadoil_cP
            q_oil_rc_m3day_ = qliq_sm3day * (1 - fw_fr_) * bo_m3m3_   ' для ускорения расчетов потом все что можно подсчитаем тут
            q_wat_rc_m3day_ = qliq_sm3day * fw_fr_ * bw_m3m3_
            q_gas_rc_m3day_ = (qliq_sm3day * (1 - fw_fr_) * rp_m3m3_ + q_gas_free_sm3day - rs_m3m3_ * qliq_sm3day * (1 - fw_fr_)) * bg_m3m3_
            qliq_rc_m3day_ = q_wat_rc_m3day_ + q_oil_rc_m3day_
            If (q_wat_rc_m3day_ + q_oil_rc_m3day_ + q_gas_rc_m3day_) > 0 Then
                gas_fraction_d_ = q_gas_rc_m3day_ / (q_wat_rc_m3day_ + q_oil_rc_m3day_ + q_gas_rc_m3day_)
                If qliq_rc_m3day_ > 0 Then
                    mu_mix_cP_ = (mu_oil_cP_ * q_oil_rc_m3day_ / qliq_rc_m3day_ +
                          mu_wat_cP_ * q_wat_rc_m3day_ / qliq_rc_m3day_) * (1 - gas_fraction_d_) +
                          mu_gas_cP_ * (1 - gas_fraction_d_)
                Else
                    mu_mix_cP_ = mu_gas_cP_
                End If
            Else
                gas_fraction_d_ = 0
                mu_mix_cP_ = 0
            End If
            rho_oil_rc_kgm3_ = 1000 * (gamma_o + rs_m3m3_ * gamma_g * const_rho_air / 1000) / bo_m3m3_
            rho_wat_rc_kgm3_ = 1000 * (gamma_w) / bw_m3m3_
            rho_liq_rc_kgm3_ = (1 - fw_fr_) * Rho_oil_rc_kgm3() + fw_fr_ * Rho_wat_rc_kgm3()
            rho_mix_rc_kgm3_ = Rho_liq_rc_kgm3() * (1 - F_g()) + Rho_gas_rc_kgm3() * F_g()
            Call Calc_ST(p_atma, t_C)

            cv_gas_JkgC_ = z_ * const_r / (Mg_kgmol() * (heat_capacity_ratio_gas_ - 1))
            cp_oil_JkgC_ = ((0.002 * t_C - 1.429) * gamma_o + 0.00267 * t_C + 3.49) * 1000   ' http://www.jmcampbell.com/tip-of-the-month/2014/04/simple-equations-to-approximate-changes-to-the-properties-of-crude-oil-with-changing-temperature/

        Catch ex As Exception

            Dim errmsg As String
            errmsg = "Error:CPVT.calc_PVT:" & ex.Message
            AddLogMsg(errmsg)
            Throw New ApplicationException(errmsg)

        End Try


        Exit Sub
    End Sub

    Private Sub Calc_ST(ByVal p_atma As Double, ByVal t_C As Double)
        ' calculate surface tension according Baker Sverdloff correlation
        Try
            'Расчет коэффициента поверхностного натяжения газ-нефть
            Dim ST68 As Double, ST100 As Double
            Dim STw74 As Double, STw280 As Double
            Dim Tst As Double, Tstw As Double
            Dim STo As Double, STw As Double, ST As Double
            Dim t_F As Double
            Dim P_psia As Double, p_MPa As Double
            t_F = t_C * 1.8 + 32
            P_psia = p_atma / 0.068046
            p_MPa = p_atma / 10
            ST68 = 39 - 0.2571 * Oil_API()
            ST100 = 37.5 - 0.2571 * Oil_API()

            If t_F < 68 Then
                STo = ST68
            Else
                Tst = t_F
                If t_F > 100 Then Tst = 100
                STo = (68 - (((Tst - 68) * (ST68 - ST100)) / 32)) * Math.Exp(-0.00086306 * P_psia)
                ' https://petrowiki.org/Interfacial_tension
                'If STo < 0 Then STo = ST68
            End If
            'Расчет коэффициента поверхностного натяжения газ-вода  (два способа)
            STw74 = (75 - (1.108 * (P_psia) ^ 0.349))
            STw280 = (53 - (0.1048 * (P_psia) ^ 0.637))
            If t_F < 74 Then
                STw = STw74
            Else
                Tstw = t_F
                If t_F > 280 Then Tstw = 280
                STw = STw74 - (((Tstw - 74) * (STw74 - STw280)) / 206)
            End If
            ' далее второй способ
            STw = 10 ^ (-(1.19 + 0.01 * p_MPa)) * 1000
            ' Расчет коэффициента поверхностного натяжения газ-жидкость
            ST = (STw * fw_fr_) + STo * (1 - fw_fr_)
            ST_oilgas_dyncm_ = STo
            ST_watgas_dyncm_ = STw
            ST_liqgas_dyncm_ = ST
        Catch ex As Exception
            Dim errmsg As String
            errmsg = "Error:CPVT.calc_ST:" & ex.Message
            AddLogMsg(errmsg)
            Throw New ApplicationException(errmsg)

        End Try

    End Sub

    Public Function Calc_rs_m3m3(ByVal p_atma As Double, ByVal t_C As Double) As Double
        'function calculates solution gas oil ratio
        Call Calc_PVT(p_atma, t_C)
        Calc_rs_m3m3 = rs_m3m3_
    End Function

    Public Function Calc_pb_atma(ByVal rsb_m3m3 As Double, ByVal t_C As Double) As Double
        'function calculates oil bubble point pressure
        Call Set_rp_rsb(rsb_m3m3, rsb_m3m3)
        '    rsb_m3m3_ = rsb_m3m3
        Call Calc_PVT(1, t_C)
        Calc_pb_atma = pb_calc_atma_
    End Function

    Public Function Calc_bo_m3m3(ByVal p_atma As Double, ByVal t_C As Double) As Double
        'Function calculates oil formation volume factor
        Call Calc_PVT(p_atma, t_C)
        Calc_bo_m3m3 = bo_m3m3_
    End Function

    Public Function Calc_mu_oil_cP(ByVal p_atma As Double, ByVal t_C As Double) As Double
        'function calculates oil viscosity
        Call Calc_PVT(p_atma, t_C)
        Calc_mu_oil_cP = mu_oil_cP_
    End Function

    Public Function Gas_fraction_d(Optional ByVal Ksep As Double = 0) As Double
        ' метод расчета доли газа в потоке для заданной жидкости при заданных условиях
        ' предполагается что свойства нефти газа и воды уже расчитаны и заданы при необходимых условиях

        Try
            Dim q_mix_ As Double
            Gas_fraction_d = 0
            q_mix_ = Q_mix_rc_m3day()   ' сохраним чтобы немного сэкономить на проверку нулевого значения
            If q_mix_ > 0 And Ksep >= 0 And Ksep < 1 Then
                Gas_fraction_d = Q_gas_rc_m3day() * (1 - Ksep) / (Q_wat_rc_m3day() + Q_oil_rc_m3day() + Q_gas_rc_m3day() * (1 - Ksep))
            End If
        Catch ex As Exception
            Dim errmsg As String
            errmsg = "Error:CPVT.gas_fraction_d:" & ex.Message
            AddLogMsg(errmsg)
            Throw New ApplicationException(errmsg)
        End Try
    End Function

    Public Function P_gas_fraction_atma(FreeGas As Double, t_C As Double, Optional Es As Double = 0, Optional P_init_atma As Double = 300) As Double
        'P_init     - давление инициализации, атм
        'FreeGas    - доля газ на приеме целевая
        'Es         - коэффициент сепарации насоса
        Dim p1 As Double
        Dim p2 As Double
        Dim max_iter As Integer, i As Integer
        Dim e As Double
        Dim p_gas As Double, p As Double
        max_iter = 100
        e = 0.0001
        p1 = P_init_atma
        p2 = 0
        Try
            For i = 1 To max_iter
                p = (p1 + p2) / 2
                Call Calc_PVT(p, t_C)
                p_gas = Gas_fraction_d(Es)
                If Math.Abs(p_gas - FreeGas) <= e Then Exit For
                If p_gas > FreeGas Then
                    p2 = p
                Else
                    p1 = p
                End If
            Next
            P_gas_fraction_atma = p
        Catch ex As Exception
            P_gas_fraction_atma = 0
        End Try

    End Function

    Public Function Rp_gas_fraction_m3m3(FreeGas As Double, p_atma As Double, t_C As Double, Optional Es As Double = 0, Optional Rp_init_m3m3 As Double = 500) As Double
        'P_init     - давление инициализации, атм
        'FreeGas    - доля газ на приеме целевая
        'Es         - коэффициент сепарации насоса
        Dim G1 As Double
        Dim g2 As Double
        Dim max_iter As Integer, i As Integer
        Dim e As Double
        Dim p_gas As Double, g As Double
        Dim rsb_back As Double
        max_iter = 100
        e = 0.0001
        G1 = Rp_init_m3m3
        g2 = 0
        rsb_back = rsb_m3m3_
        Try
            For i = 1 To max_iter
                g = (G1 + g2) / 2
                Call Set_rp_rsb(g, rsb_back)
                '        rsb_m3m3 = rsb_back
                '        rp_m3m3 = g
                Call Calc_PVT(p_atma, t_C)
                p_gas = Gas_fraction_d(Es)
                If Math.Abs(p_gas - FreeGas) <= e Then Exit For
                If p_gas > FreeGas Then
                    G1 = g
                Else
                    g2 = g
                End If
            Next
            Rp_gas_fraction_m3m3 = g
        Catch ex As Exception
            Rp_gas_fraction_m3m3 = 0
        End Try
    End Function

    Public Function Get_clone_mod_after_separation(p_atma As Double, t_C As Double, Ksep As Double,
                                         Optional ByVal GasGoesIntoSolution As Boolean = False) As CPVT
        Dim newfluid As CPVT
        newfluid = Me.Clone
        Call newfluid.Mod_after_separation(p_atma, t_C, Ksep, GasGoesIntoSolution)
        Get_clone_mod_after_separation = newfluid
    End Function

    Public Sub Mod_after_separation(ByVal p_atma As Double, ByVal t_C As Double, ByVal Ksep As Double,
                                         Optional ByVal GasGoesIntoSolution As Boolean = False) 'As CPVT
        ' функция модификации свойств нефти после сепарации
        ' удаление части газа меняет свойства нефти - причем добавление газа свойства не трогает
        ' на входе условия при которых проходила сепарация
        Dim Rs As Double
        Dim Bo As Double
        Dim Pb_Rs_curve As New CInterpolation ' хранилище кривой зависимости газосодержания от давления насыщения
        Dim Bo_Rs_curve As New CInterpolation
        Dim pb_atma_tab As Double, rsb_m3m3_tab As Double, Bo_m3m3_tab As Double
        Dim Delta As Double
        Dim i As Integer
        Const n = 10
        Dim Rpnew_with_Ksep As Double   ' новый ГФ с учетом сепарации газа
        Dim Rpnew_Ksep_1 As Double      ' новый ГФ без учета сепарации
        Dim Rpnew As Double
        ' найдем сколько газа осталось в растворе при условиях сепарации

        Try

            With Me
                Rs = .Calc_rs_m3m3(p_atma, t_C)
                Bo = .Calc_bo_m3m3(p_atma, t_C)

                ' оценим как изменится газовый фактор за счет ухода части газа из потока
                ' оцениваем двумя способами - буду влиять на работу опции растворения газа (фазовой неравновесности)
                '   - с учетом сепарации (показывает новый ГФ если газ может потом растворяться)
                '   - без учета сепарации (Ksep = 1 показывает ГФ если газ может только сжиматься)

                Rpnew_with_Ksep = .Rp_m3m3 - (.Rp_m3m3 - Rs) * Ksep
                Rpnew_Ksep_1 = .Rp_m3m3 - (.Rp_m3m3 - Rs)


                Delta = (.Pb_calc_atma - 1) / n    ' считать будет только в диапазоне где определены Pb за ним будет линейно экстраполировать
                ' запишем зависимость газосодержания от давления насыщения на память
                For i = 0 To n
                    pb_atma_tab = 1 + Delta * i
                    rsb_m3m3_tab = .Calc_rs_m3m3(pb_atma_tab, .tres_C)
                    Pb_Rs_curve.AddPoint(rsb_m3m3_tab, pb_atma_tab)
                    Bo_m3m3_tab = .Calc_bo_m3m3(pb_atma_tab, .tres_C)
                    Bo_Rs_curve.AddPoint(rsb_m3m3_tab, Bo_m3m3_tab)
                Next i
                ' найдем сколько всего газа осталось в потоке

                If GasGoesIntoSolution Then   ' тогда газ успеет растворится
                    Rpnew = Rpnew_with_Ksep
                Else                                   ' газ не растворяется, то же самое, что Ксеп = 1
                    Rpnew = Rpnew_Ksep_1
                End If

                If Rpnew < .Rsb_m3m3 Then
                    ' Если газовый фактор становится меньше газосодержания, тогда надо скорректировать газосодержание и давление насыщения,
                    ' которое будет от него зависеть
                    .pb_atma = Pb_Rs_curve.GetPoint(Rpnew)
                    .bob_m3m3 = Bo_Rs_curve.GetPoint(Rpnew)
                    .Rsb_m3m3 = Rpnew
                    ' иначе газа из раствора не сепарировался - свойства не менялись ничего делать не надо
                End If
                ' итоговый газовый фактор всегда с учетом сепараци
                .Rp_m3m3 = Rpnew_with_Ksep
            End With
        Catch ex As Exception
            Dim errmsg As String
            errmsg = "Error:CPVT.mod_after_separation:" & ex.Message
            AddLogMsg(errmsg)
            Throw New ApplicationException(errmsg)
        End Try

    End Sub

    Public Sub init_json(json As String)
        Dim d = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(json)
        Call init_dictionary(d)
    End Sub

    Public Sub init_dictionary(dict As Dictionary(Of String, Object))
        Dim gamma_gas As Double
        Dim gamma_oil As Double
        Dim gamma_wat As Double
        Dim rsb_m3m3 As Double
        Dim rp_m3m3 As Double
        Dim pb_atma As Double
        Dim tres_C As Double
        Dim bob_m3m3 As Double
        Dim muob_cP As Double
        Dim PVTcorr As Integer
        Dim ksep_fr As Double
        Dim p_ksep_atma As Double
        Dim t_ksep_C As Double
        Dim gas_only As Boolean

        Dim errmsg As String
        Dim key As String

        Try
            With dict
                key = "gamma_oil"
                If .ContainsKey(key) Then
                    gamma_oil = .Item(key)
                Else
                    errmsg = "CPVT.init_dictionary. error: " & key & " must be given"
                End If

                key = "gamma_gas"
                If .ContainsKey(key) Then
                    gamma_gas = .Item(key)
                Else
                    errmsg = "CPVT.init_dictionary. error: " & key & " must be given"
                End If

                key = "rsb_m3m3"
                If .ContainsKey(key) Then
                    rsb_m3m3 = .Item(key)
                Else
                    errmsg = "CPVT.init_dictionary. error: " & key & " must be given"
                End If

                If .ContainsKey("gamma_wat") Then gamma_wat = .Item("gamma_wat")
                If .ContainsKey("rp_m3m3") Then rp_m3m3 = .Item("rp_m3m3")
                If .ContainsKey("pb_atma") Then pb_atma = .Item("pb_atma")
                If .ContainsKey("t_res_C") Then tres_C = .Item("t_res_C")
                If .ContainsKey("bob_m3m3") Then bob_m3m3 = .Item("bob_m3m3")
                If .ContainsKey("muob_cP") Then muob_cP = .Item("muob_cP")
                If .ContainsKey("PVTcorr") Then PVTcorr = .Item("PVTcorr")
                If .ContainsKey("ksep_fr") Then ksep_fr = .Item("ksep_fr")
                If .ContainsKey("p_ksep_atma") Then p_ksep_atma = .Item("p_ksep_atma")
                If .ContainsKey("t_ksep_C") Then t_ksep_C = .Item("t_ksep_C")
                If .ContainsKey("gas_only") Then gas_only = .Item("gas_only")
            End With

            Call Init(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, pb_atma, bob_m3m3, PVTcorr, tres_C, rp_m3m3, muob_cP)
            Me.gas_only = gas_only
            Me.ksep_fr = ksep_fr
            Me.t_ksep_C = t_ksep_C
            If ksep_fr > 0 And ksep_fr <= 1 And p_ksep_atma > 0 And t_ksep_C > 0 Then
                Call Mod_after_separation(p_ksep_atma, t_ksep_C, ksep_fr, True)
            End If
            Exit Sub
        Catch ex As Exception
            AddLogMsg(errmsg)
            Throw New ApplicationException(errmsg)
        End Try
    End Sub

End Class

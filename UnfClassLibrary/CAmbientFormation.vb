'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
'
' класс для описания свойств пород окружающих скважину на всем протяжении скважины
' задает геотермальный градиент в скважине  (в том числе с учетом зависимости от глубины и например наличия вечной мерзлоты)
' позволяет рассчитать общие коэффициент теплопроводности для различных условий
' необходим для корректного расчета распределения температуры в скважине
'
'


' пока положим эти параметры не зависящими от глубины, хотя потом можно будет поменять
Option Explicit On
Imports System.Math
Imports Newtonsoft.Json

Public Class CAmbientFormation
    Public therm_cond_form_WmC As Double      ' теплопроводность породы Дж/сек/м/С
    Public sp_heat_capacity_form_JkgC As Double        ' удельная теплоемкость породы  specific heat capacity
    Public density_formation_kgm3 As Double       ' плотность породы вокруг скважины

    ' termal conductivity теплопроводность
    Public therm_cond_cement_WmC As Double          ' теплопроводность цемента вокруг скважины
    Public therm_cond_tubing_WmC As Double          ' теплопроводность металла НКТ
    Public therm_cond_casing_WmC As Double          ' теплопроводность металла эксплуатационной колонны

    ' convective heat transfer coeficients
    Public heat_transfer_casing_liquid_Wm2C As Double       ' конвективная теплопередача через затруб с жидкостью  Дж/м2/сек/С
    Public heat_transfer_casing_gas_Wm2C As Double       ' теплопередача через затруб с газом (радиационная)
    Public heat_transfer_fluid_convection_Wm2C As Double       ' теплопередача конвективная в потоке жидкости


    ' радиусы для проведения расчета по температуре
    ' хотя это больше относится к конструкции скважины, но тут удобнее это разместить, тем более что значительного влияния на расчет нет
    ' в перспективе можно будет перевести в массивы и формировать снаружи
    Public rti_m As Double           ' НКТ внутренний
    Public rto_m As Double           ' НКТ наружный
    Public rci_m As Double           ' Эксп колонна внутренний
    Public rco_m As Double           ' Эксп колонна наружный
    Public rcem_m As Double          ' Радиус цементного кольца вокруг скважины
    Public rwb_m As Double           ' Радиус цементного кольца вокруг скважины

    Public t_calc_hr As Double               ' время на которое вычисляется распределение давления
    'Private td_d_ As Double                ' безразмерное время расчета
    'Private TD_ As Double                  ' безразмерное температура

    Public amb_temp_curve As New CInterpolation

    Public TGeoGrad_C100m_ As Double

    Public h_vert_data_m As Double
    Public reservoir_temp_data_C As Double
    Public surf_temp_data_C As Double

    Public h_dyn_m As Double
    Public h_pump_m As Double

    Private Sub Class_Initialize(Optional ByVal therm_cond_form_WmC_ As Double = 2.4252,
                                 Optional ByVal sp_heat_capacity_form_JkgC_ As Double = 200,
                                 Optional ByVal density_formation_kgm3_ As Double = 4000,
                                 Optional ByVal therm_cond_cement_WmC_ As Double = 6.965,
                                 Optional ByVal therm_cond_tubing_WmC_ As Double = 32,
                                 Optional ByVal therm_cond_casing_WmC_ As Double = 32,
                                 Optional ByVal heat_transfer_casing_liquid_Wm2C_ As Double = 200,
                                 Optional ByVal heat_transfer_casing_gas_Wm2C_ As Double = 10,
                                 Optional ByVal heat_transfer_fluid_convection_Wm2C_ As Double = 200,
                                 Optional ByVal rti_m_ As Double = 0.06,
                                 Optional ByVal rto_m_ As Double = 0.07,
                                 Optional ByVal rci_m_ As Double = 0.124,
                                 Optional ByVal rco_m_ As Double = 0.125,
                                 Optional ByVal rcem_m_ As Double = 0.3,
                                 Optional ByVal rwb_m_ As Double = 0.3,
                                 Optional ByVal h_vert_data_m_ As Double = 2500,
                                 Optional ByVal reservoir_temp_data_C_ As Double = 95,
                                 Optional ByVal surf_temp_data_C_ As Double = 25,
                                 Optional ByVal t_calc_hr_ As Double = 24 * 10,
                                 Optional ByVal h_dyn_m_ As Double = -1,
                                 Optional ByVal h_pump_m_ As Double = -1,
                                 Optional ByVal TGeoGrad_C100m As Double = -1)

        therm_cond_form_WmC = therm_cond_form_WmC_    ' теплопроводность породы Дж/сек/м/С
        sp_heat_capacity_form_JkgC = sp_heat_capacity_form_JkgC_          ' теплоемкость породы
        density_formation_kgm3 = density_formation_kgm3_                 ' плотность породы вокруг скважины

        therm_cond_cement_WmC = therm_cond_cement_WmC_       ' теплопроводность цемента вокруг скважины
        therm_cond_tubing_WmC = therm_cond_tubing_WmC_          ' теплопроводность металла НКТ
        therm_cond_casing_WmC = therm_cond_casing_WmC_          ' теплопроводность металла эксплуатационной колонны

        ' convective heat transfer coeficients
        heat_transfer_casing_liquid_Wm2C = heat_transfer_casing_liquid_Wm2C_        ' конвективная теплопередача через затруб с жидкостью  Дж/м2/сек/С
        heat_transfer_casing_gas_Wm2C = heat_transfer_casing_gas_Wm2C_            ' теплопередача через затруб с газом (радиационная)
        heat_transfer_fluid_convection_Wm2C = heat_transfer_fluid_convection_Wm2C_     ' теплопередача конвективная в потоке жидкости

        ' радиусы для проведения расчета
        rti_m = rti_m_           ' НКТ внутренний
        rto_m = rto_m_  ' НКТ наружный
        rci_m = rci_m_          ' Эксп колонна внутренний
        rco_m = rco_m_ ' Эксп колонна наружный
        rcem_m = rcem_m_         ' Радиус цементного кольца вокруг скважины
        rwb_m = rwb_m_

        ' исходные данные по умолчанию, чтобы все считало
        h_vert_data_m = h_vert_data_m_
        reservoir_temp_data_C = reservoir_temp_data_C_
        surf_temp_data_C = surf_temp_data_C_

        t_calc_hr = t_calc_hr_   ' задаем по умолчанию время расчета распределения температуры через 10 дней

        h_dyn_m = h_dyn_m_
        h_pump_m = h_pump_m_

        amb_temp_curve.AddPoint(0, surf_temp_data_C)
        amb_temp_curve.AddPoint(h_vert_data_m, reservoir_temp_data_C)

        TGeoGrad_C100m_ = TGeoGrad_C100m
        TGeoGrad_C100m_ = h_vert_data_m / 100 / (reservoir_temp_data_C - surf_temp_data_C)


    End Sub


    ' градиент давления от температуры
    Public Function amb_temp_grad_Cm(h_vert_m As Double) As Double
        If amb_temp_curve Is Nothing Then
            amb_temp_grad_Cm = TGeoGrad_C100m_ / 100
        Else
            amb_temp_grad_Cm = (amb_temp_curve.GetPoint(h_vert_m + 1) - amb_temp_curve.GetPoint(h_vert_m)) / 1
        End If
    End Function

    ' температура на глубине
    Public Function amb_temp_C(h_vert_m As Double) As Double
        If amb_temp_curve Is Nothing Then
            amb_temp_C = reservoir_temp_data_C + (h_vert_m - h_vert_data_m) * amb_temp_grad_Cm(h_vert_m)
        Else
            amb_temp_C = amb_temp_curve.GetPoint(h_vert_m)
        End If
    End Function


    Private ReadOnly Property td_d() As Double
        Get
            td_d = therm_cond_form_WmC * t_calc_hr * const_convert_hr_sec / density_formation_kgm3 / sp_heat_capacity_form_JkgC / (rwb_m ^ 2)
        End Get
    End Property

    Private ReadOnly Property td() As Double
        Get
            td = Log(Exp(-0.2 * td_d) + (1.5 - 0.3719 * Exp(-td_d)) * (td_d ^ 0.5))
        End Get
    End Property

    Private Function Lr_1m(wt_kgsec As Double, Uto_Jm2secC As Double, Cp_JkgC As Double) As Double
        If wt_kgsec <> 0 Then
            Lr_1m = 2 * const_Pi / (Cp_JkgC * wt_kgsec) * (Uto_Jm2secC * therm_cond_form_WmC / (therm_cond_form_WmC + Uto_Jm2secC * td))
        Else
            Lr_1m = 10000
        End If
    End Function

    ' функция расчета градиента температуры
    Function calc_dtdl_Cm(ByVal h_vert_m As Double,
                       ByVal sinTheta_deg As Double,
                       ByVal T1_C As Double,
                       ByVal w_kgsec As Double,
                       ByVal Cp_JkgC As Double,
                       Optional dPdL_atmm As Double = 0,
                       Optional v_ms As Double = 0,
                       Optional dvdL_msm As Double = 0,
                       Optional Cj_Catm As Double = 0,
                       Optional FlowAlongCoord As Boolean = True) As Double
        ' h_vert_m     -  vertical depth where calculation take place
        ' sinTheta_deg - angle sin
        ' T1_C         - fluid temp at depth gien
        ' W_kgsec      - mass rate of fluid
        ' Cp_JkgC      - heat capasity
        ' dPdL_atmm    - pressure gradient at depth given (needed to account Joule Tompson effect)
        ' v_ms         - velocity of fluid mixture
        ' dvdL_msm     - acceleration of fluid mixture. acount inetria force influence (should be small but ..)
        ' Cj_Catm      - коэффициент Джоуля Томсона Joule Thomson coeficient
        ' flowUp       - flow direction
        Dim Lr As Double
        Dim Uto As Double
        Dim h As Double
        Dim sign As Integer
        ' если потока нет, то берем температуру извне
        ' if mass flow rate is zero - take ambient temp gradient
        If w_kgsec = 0 Then
            calc_dtdl_Cm = amb_temp_grad_Cm(h_vert_m)
            Exit Function
        End If
        ' set Uto - temperature emission depents on well condition
        If h_vert_m > h_pump_m Then
            Uto = Uto_cas_Jm2secC
        ElseIf h_vert_m > h_dyn_m Then
            Uto = Uto_tub_liqcas_Jm2secC
        Else
            Uto = Uto_tub_gascas_Jm2secC
        End If

        If FlowAlongCoord Then
            sign = -1
        Else
            sign = 1
        End If

        Lr = Lr_1m(w_kgsec, Uto, Cp_JkgC)
        calc_dtdl_Cm = sign * (T1_C - amb_temp_C(h_vert_m)) * Lr
        calc_dtdl_Cm = calc_dtdl_Cm - (const_g * sinTheta_deg / Cp_JkgC + v_ms / Cp_JkgC * dvdL_msm - Cj_Catm * dPdL_atmm)
    End Function

    Public ReadOnly Property Uto_cas_Jm2secC() As Double
        Get
            Uto_cas_Jm2secC = 1 / (
                             Log(rwb_m / rco_m) / therm_cond_cement_WmC +
                             Log(rco_m / rci_m) / therm_cond_casing_WmC +
                             1 / rci_m / heat_transfer_fluid_convection_Wm2C
                        )
        End Get

    End Property

    Public ReadOnly Property Uto_tub_liqcas_Jm2secC() As Double
        Get
            Uto_tub_liqcas_Jm2secC = 1 / (
                            1 * Log(rwb_m / rco_m) / therm_cond_cement_WmC +
                            1 * Log(rco_m / rci_m) / therm_cond_casing_WmC +
                            1 / rto_m / (heat_transfer_casing_gas_Wm2C + heat_transfer_casing_liquid_Wm2C) +
                            1 * Log(rto_m / rti_m) / therm_cond_tubing_WmC +
                            1 / rti_m / heat_transfer_fluid_convection_Wm2C
                        )
        End Get
    End Property

    Public ReadOnly Property Uto_tub_gascas_Jm2secC() As Double
        Get
            Uto_tub_gascas_Jm2secC = 1 / (
                            1 * Log(rwb_m / rco_m) / therm_cond_cement_WmC +
                            1 * Log(rco_m / rci_m) / therm_cond_casing_WmC +
                            1 / rto_m / (heat_transfer_casing_gas_Wm2C) +
                            1 * Log(rto_m / rti_m) / therm_cond_tubing_WmC +
                            1 / rti_m / heat_transfer_fluid_convection_Wm2C
                        )
        End Get
    End Property

    Public Sub init_amb_temp_points(ByVal h1 As Double,
                                ByVal t1 As Double,
                                ByVal h2 As Double,
                                ByVal t2 As Double)

        Dim geo_grad_curve As New CInterpolation
        geo_grad_curve.AddPoint(h1, t1)
        geo_grad_curve.AddPoint(h2, t2)
        Me.amb_temp_curve = geo_grad_curve

    End Sub

    Public Sub init_amb_temp_arr(ByVal tamb_arr_C(,) As Double,
                                 ByVal tamb_arr_hmes_m() As Double)

        Dim geo_grad_curve As New CInterpolation
        Dim t As Double
        Dim h As Double

        If tamb_arr_hmes_m.Any Then
            If tamb_arr_C.GetUpperBound(2) < 1 Then ' развернул знак
                '    Call geo_grad_curve.loadFromVertRange(tamb_arr_C) ' read deviation survey from one table
                'Else
                t = tamb_arr_C(1, 1)
                geo_grad_curve.AddPoint(0, t)
                geo_grad_curve.AddPoint(1000, t)
            End If
            'ElseIf Not tamb_arr_C.ToString.Any And Not tamb_arr_hmes_m.Any Then
            '    Call geo_grad_curve.loadFromVertRange(tamb_arr_hmes_m, tamb_arr_C) ' read deviation survey from two collumns
        End If

        ' correction if only one number pre curve have been read
        If geo_grad_curve.Num_points = 1 Then
            t = geo_grad_curve.PointY(1)
            h = geo_grad_curve.PointX(1)
            geo_grad_curve.AddPoint(h + 1000, t)
        End If

        Me.amb_temp_curve = geo_grad_curve
    End Sub

    Public Sub set_props_json(json As String)
        Dim dict As Dictionary(Of String, Double)
        If json.Length > 3 Then
            dict = JsonConvert.DeserializeObject(Of Dictionary(Of String, Double))(json) ' APrseJson(json)
            With dict
                therm_cond_form_WmC = .Item("therm_cond_form_WmC")
                sp_heat_capacity_form_JkgC = .Item("sp_heat_capacity_form_JkgC")
                therm_cond_cement_WmC = .Item("therm_cond_cement_WmC")
                therm_cond_tubing_WmC = .Item("therm_cond_tubing_WmC")
                therm_cond_casing_WmC = .Item("therm_cond_casing_WmC")
                heat_transfer_casing_liquid_Wm2C = .Item("heat_transfer_casing_liquid_Wm2C")
                heat_transfer_casing_gas_Wm2C = .Item("heat_transfer_casing_gas_Wm2C")
                heat_transfer_fluid_convection_Wm2C = .Item("heat_transfer_fluid_convection_Wm2C")
                t_calc_hr = .Item("t_calc_hr")
            End With
        End If

    End Sub

End Class

'=======================================================================================
'Unifloc 7.24  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
' types definition module
' converted 16/05/20



Module u7_types

    ' hydraulic correlations types
    Public Enum H_CORRELATION As Integer
        BeggsBrill = 0
        Ansari = 1
        Unified = 2
        Gray = 3
        HagedornBrown = 4
        SakharovMokhov = 5
        gas = 10
        Water = 11
    End Enum

    ' PVT correlations set to be used
    Public Enum PVT_correlation As Integer
        Standing_based = 0 '
        McCain_based = 1 '
        straigth_line = 2
    End Enum

    ' z factor (gas compressibility) options
    Public Enum Z_CORRELATION As Integer
        z_BB = 0
        z_Dranchuk = 1
        z_Kareem = 2
    End Enum

    ' Structure determines the method of temperature calculation in well
    Public Enum TEMP_CALC_METHOD
        StartEndTemp = 0
        GeoGradTemp = 1
        AmbientTemp = 2
    End Enum

    ' gas separation in well at pump intake. calculation method
    Public Enum SEPAR_CALC_METHOD
        fullyManual = 3
        valueManual = 2
        pressureManual = 1
        byCorrealation = 0
    End Enum

    '' standard port sizes for whetherford r1 glv
    'Public Enum GLV_R1_PORT_SIZE
    '    R1_port_1_8 = 3.18
    '    R1_port_5_32 = 3.97
    '    R1_port_3_16 = 4.76
    '    R1_port_1_4 = 6.35
    '    R1_port_5_16 = 7.94
    'End Enum

    ' multiphase flow in pipe and well calculation method
    Public Structure PARAMCALC
        Public correlation As H_CORRELATION         ' multiphase hydraulic correlation
        Public CalcAlongCoord As Boolean            ' calculation direction flag
        ' if True - pressure at lowest coordinate is given
        '           pressure at higher coordinate calculated
        '           for well (0 coord at top, hmes at bottom)
        '           equal to calc from top to bottom
        '    False - otherwise
        Public FlowAlongCoord As Boolean            ' flow direction flag
        Public temp_method As TEMP_CALC_METHOD      ' temperature caclulation method
        Public length_gas_m As Double               ' length in pipe where correlation changes.
        ' for points with cooed less then  length_gas_m
        ' gas correlation applied,
        ' other points - multiphase correlation applied
        ' allows to model easily static level in well
        Public start_length_gas_m As Double
    End Structure

    ' Structure to describe thermobaric conditions (for calculations)
    Public Structure PTtype
        Public p_atma As Double
        Public t_C As Double
    End Structure

    ' Structure for storing data about dynamic level
    Public Structure PCAS_HDYN_type
        Public p_cas_atma As Double
        Public hdyn_m As Double
        Public self_flow_condition As Boolean
        Public pump_off_condition As Boolean
        Public correct As Boolean
    End Structure

    ' Structure for describing the operation of an electric motor
    Public Structure MOTOR_DATA
        Public U_lin_V As Double       ' voltage linear (between phases)
        Public I_lin_A As Double       ' Linear current (in line)
        Public U_phase_V As Double     ' phase voltage (between phase and zero)
        Public I_phase_A As Double     ' phase current (in winding)
        Public f_Hz As Double          ' frequency synchronous (field rotation)
        Public eff_d As Double         ' Efficiency
        Public cosphi As Double        ' power factor
        Public s_d As Double           ' slippage
        Public Pshaft_kW As Double     ' mechanical power on the shaft
        Public Pelectr_kW As Double    ' power supply electric
        Public Mshaft_Nm As Double     ' torque on the shaft - mechanical
        Public load_d As Double        ' motor load
    End Structure

    ' ESP description to be loaded from DB
    ' combined in Structure to decrease a mess in CESPpump
    Public Structure ESP_PARAMS

        Public ID As String                    ' ID  из базы роспампа
        Public source As String                ' источник данных о характеристиках насоса - влият на способ расчета характеристик
        Public manufacturer As String          ' производитель насоса (справочный параметр)
        Public name As String
        Public stages_max As Integer           ' максимальной количество ступеней в насосе (из базы)
        Public rate_max_sm3day As Double        ' максимальный дебит насос (из базы) - хорошо бы для надежности определять параметр из характеристики
        Public rate_nom_sm3day As Double
        Public rate_opt_min_sm3day As Double    ' границы оптимального диапазона для насоса - минимум
        Public rate_opt_max_sm3day As Double    ' границы оптимального диапазона  - максимум
        Public freq_Hz As Double               ' частота насоса для номинальной характеристики в базе

        ' характеристика заданные по точкам
        Public head_points() As Double
        Public rate_points() As Double
        Public power_points() As Double
        Public eff_points() As Double

        Public stage_height_m As Double           ' примерная высота ступени
        Public d_od_m As Double                  ' внешний диаметр ЭЦН
        Public d_cas_min_m As Double              ' минимальный диаметр обсадной колонны, заданный производителем оборудования
        Public d_shaft_m As Double             ' диаметр вала для насоса
        Public area_shaft_m2 As Double            ' площадь поперечного сечения вала   (дублирует диаметр, но задается производителем)
        Public shaft_power_limit_W As Double       ' максимальная мощность передаваемая валом на номинальной частоте
        Public shaft_power_limit_max_W As Double    ' максимальная мощность передаваемая валом на номинальной частоте для высокопрочного вала
        Public housing_pressure_limit_atma As Double ' максимальное давление на корпус
        Public nom_slip_rpm As Double
        Public eff_max As Double
    End Structure

    ' Structure of extended description of multiphase flow parameters at a point
    Public Structure PIPE_FLOW_PARAMS
        Public md_m As Double         ' pipe measured depth (from start - top)
        Public vd_m As Double         ' pipe vertical depth from start - top
        Public diam_mm As Double      ' pipe diam
        Public p_atma As Double       ' pipe pressure at measured depth
        Public t_C As Double          ' pipe temp at measured depth

        Public dp_dl As Double
        Public dt_dl As Double

        Public dpdl_g_atmm As Double  ' gravity gradient at measured depth
        Public dpdl_f_atmm As Double  ' friction gradient at measured depth
        Public dpdl_a_atmm As Double  ' acceleration gradient at measured depth
        Public v_sl_msec As Double    ' superficial liquid velosity
        Public v_sg_msec As Double    ' superficial gas velosity
        Public h_l_d As Double        ' liquid hold up
        Public fpat As Double         ' flow pattern code
        Public thete_deg As Double
        Public roughness_m As Double

        Public rs_m3m3 As Double     ' dissolved gas in oil in the stream
        Public gasfrac As Double     ' gas flow rate

        Public mu_oil_cP As Double   ' oil viscosity in flow
        Public mu_wat_cP As Double   ' water viscosity in the flow
        Public mu_gas_cP As Double   ' gas viscosity in flow
        Public mu_mix_cP As Double   ' viscosity of the mixture in the flow

        Public Rhoo_kgm3 As Double   ' oil Density
        Public Rhow_kgm3 As Double   ' water Density
        Public rhol_kgm3 As Double   ' liquid density
        Public Rhog_kgm3 As Double   ' gas Density
        Public rhomix_kgm3 As Double ' density of the mixture in the thread

        Public q_oil_m3day As Double ' oil consumption in working conditions
        Public qw_m3day As Double    ' water consumption in working conditions
        Public Qg_m3day As Double    ' gas flow rate under operating conditions

        Public mo_kgsec As Double    ' mass flow rate of oil in working conditions
        Public mw_kgsec As Double    ' mass flow rate in working conditions
        Public mg_kgsec As Double    ' mass flow rate of gas under operating conditions

        Public vl_msec As Double     ' fluid velocity is real
        Public vg_msec As Double     ' real gas velocity

    End Structure

    ' Structure of description of free gas behavior when increasing the pressure
    ' relevant for ESPs where pressure rises
    ' The free gas can either dissolve into the stream or simply compress
    'Public Enum GAS_INTO_SOLUTION
    '    GasGoesIntoSolution = 1
    '    GasnotGoesIntoSolution = 0
    'End Enum

    ' Structure showing the way of saving the extended calculation results
    ' determines which set of calculated distribution curves will be saved
    Public Enum CALC_RESULTS
        nocurves = 0
        maincurves = 1
        allCurves = 2
    End Enum


    '=========================================================================================
    'types support functions
    '=========================================================================================

    ' flow parameter setting function in the pipe or well
    Public Function Set_calc_flow_param(
                Optional ByVal calc_along_coord As Boolean = False,
                Optional ByVal flow_along_coord As Boolean = False,
                Optional ByVal hcor As H_CORRELATION = H_CORRELATION.Ansari,
                Optional ByVal temp_method As TEMP_CALC_METHOD = TEMP_CALC_METHOD.StartEndTemp,
                Optional ByVal length_gas_m As Double = 0,
                Optional ByVal start_length_gas_m As Double = 0) As PARAMCALC
        ' calc_along_coord - calculation direction flag
        ' flow_along_coord - flow direction relative to coordinate
        ' hcor             - hydraulic correlation selector
        ' temp_method      - temperature method selector
        ' length_gas_m     - boundary of gas correlation application in flow

        Dim prm As PARAMCALC
        prm.CalcAlongCoord = calc_along_coord
        prm.FlowAlongCoord = flow_along_coord
        prm.correlation = hcor
        prm.temp_method = temp_method
        prm.length_gas_m = length_gas_m
        prm.start_length_gas_m = start_length_gas_m
        Set_calc_flow_param = prm
    End Function

    Public Function Sum_PT(PT1 As PTtype, PT2 As PTtype) As PTtype
        Sum_PT.p_atma = PT1.p_atma + PT2.p_atma
        Sum_PT.t_C = PT1.t_C + PT2.t_C
    End Function

    Public Function Subtract_PT(PT1 As PTtype, PT2 As PTtype) As PTtype
        Subtract_PT.p_atma = PT1.p_atma - PT2.p_atma
        Subtract_PT.t_C = PT1.t_C - PT2.t_C
    End Function

    Public Function Set_PT(ByVal p As Double, ByVal t As Double) As PTtype
        Set_PT.p_atma = p
        Set_PT.t_C = t
    End Function

    Public Function PT_to_array(pt As PTtype)
        PT_to_array = {pt.p_atma, pt.t_C}
    End Function

    'Public Function LBound(arr())
    '    LBound = arr.GetLowerBound(0)

    'End Function

    'Public Function UBound(arr())
    '    UBound = arr.GetUpperBound(0)

    'End Function
End Module

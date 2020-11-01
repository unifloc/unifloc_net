﻿'=======================================================================================
'Unifloc 7.24  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
' constant definition module


Public Module u7_const

    Public Const vbObjectError = 1000   ' added for compatibility with VBA
    Public Const esp_db_name = "\ESP_json.db"
    ' only database as global variable here
    ' in order to reduce db file read
    Public ESP_base_dictionary As Dictionary(Of String, ESP_dict)
    ' ----  below same cose as for VBA version
    Public Const const_unifloc_version = "7.24 net"

    Public Const const_name_qliq_m3day As String = "qliq_m3day"
    Public Const const_name_fw_perc As String = "fw_perc"
    Public Const const_name_pcas_atma As String = "pcas_atma"
    Public Const const_name_h_dyn_m As String = "h_dyn_m"
    Public Const const_name_p_line_atma As String = "p_line_atma"
    Public Const const_name_pbuf_atma As String = "pbuf_atma"
    Public Const const_name_pwf_atma As String = "pwf_atma"
    Public Const const_name_gamma_oil As String = "gamma_oil"
    Public Const const_name_gamma_water As String = "gamma_water"
    Public Const const_name_gamma_gas As String = "gamma_gas"
    Public Const const_name_rp_m3m3 As String = "rp_m3m3"
    Public Const const_name_rsb_m3m3 As String = "rsb_m3m3"
    Public Const const_name_pb_atma As String = "pb_atma"
    Public Const const_name_hResMes_m As String = "HResMes_m"
    Public Const const_name_hPumpMes_m As String = "HPumpMes_m"
    Public Const const_name_dchoke_mm As String = "Dchoke_mm, mm"
    Public Const const_name_roughness_m As String = "roughness_m, м"
    Public Const const_name_ESP_qliq_m3day As String = "ESP_qliq_m3day"
    Public Const const_name_ESP_num_stages As String = "ESP_num_stages"
    Public Const const_name_ESP_freq_Hz As String = "ESP_freq_Hz"
    Public Const const_name_ESP_p_int_atma As String = "ESP_p_int_atma"
    Public Const const_name_Pres_atma As String = "Pres_atma"
    Public Const const_name_pi_sm3dayatm As String = "pi_sm3dayatm"
    Public Const const_name_ESP_t_int_C As String = "ESP_t_int_C"
    Public Const const_name_tres_C As String = "tres_C"
    Public Const const_name_hmesCurve As String = "HmesCurve"
    Public Const const_name_dcasCurve As String = "DcasCurve"
    Public Const const_name_dtubCurve As String = "DtubCurve"
    Public Const const_name_TAmbCurve As String = "TAmbCurve"

    Public Const str_p_curve = "c_P"
    Public Const str_t_curve = "c_T"
    Public Const str_VLPcurve = "VLPcurve"                      ' кривая оттока -  зависимость забойного давления от дебита жидкости
    Public Const str_HvertCurve = "HvertCurve"                  ' кривая траектории (кривизны) скважины
    Public Const str_DcasCurve = "DcasCurve"                    ' кривая изменения диаметра эксплуатационной колонны
    Public Const str_DtubCurve = "DtubCurve"                    ' кривая изменения диаметра НКТ
    Public Const str_RoughnessCasCurve = "RoughnessCasCurve"    ' кривая изменения шероховатости по трубе эксплуатационной колонны
    Public Const str_RoughnessTubCurve = "RoughnessTubCurve"    ' кривая изменения шероховатости по трубе НКТ
    Public Const str_Hd_Depend_p_wf = "Hd_Depend_p_wf"            ' кривая - зависимость динамического уровня от забойного давления, при заданном давлении в затрубе
    Public Const str_Pan_Depend_p_wf = "Pan_Depend_p_wf"          ' кривая - зависимость затрубного давления от забойного давления
    ' зависимости лин давления и буферного давления от дин уровня логично показывать на одном графике
    Public Const str_plin_Depend_p_wf = "plin_Depend_p_wf"        ' кривая - зависимость линейного давления от дин уровня
    Public Const str_pbuf_pwf_curve = "pbuf_pwf_curve"          ' зависимость буферного давления от дин уровня
    Public Const str_ksep_natQl_curve = "ksep_natQl_curve"         ' зависимость коэффициента сепарации от дебита
    Public Const str_ksep_natRp_curve = "ksep_natRp_curve"         ' зависимость коэффициента сепарации от газового фактора
    Public Const str_ksep_totalQl_curve = "ksep_totalQl_curve"     ' кривая общего коэффициента сепарации от дебита
    Public Const str_ksep_totalRp_curve = "ksep_totalRp_curve"     ' кривая общего коэффициента сепарации от газового фактора
    Public Const str_ksep_gassepQl_curve = "ksep_gassepQl_curve"   ' кривая коэффициента сепарации газосепаратора от дебита
    Public Const str_ksep_gassepRp_curve = "ksep_gassepQl_curve"   ' кривая коэффициента сепарации газосепаратора от газового фактора
    Public Const str_Pdisc_calibr_head_curve = "Pdisc_calibr_head_curve"         ' кривая зависимости давления на устье от деградации напора УЭЦН
    Public Const str_TambHmes_curve = "TambHmes_curve"           ' профиль температуры окружающего простраства от измеренный координаты
    Public Const str_PtubHmes_curve = "PtubHmes_curve"           ' профиль давления по стволу скважины по ниже НКТ и по НКТ до устья
    Public Const str_TtubHmes_curve = "TtubHmes_curve"           ' профиль температуры по стволу скважины по НКТ
    Public Const str_PcasHmes_curve = "PcasHmes_curve"           ' профиль давления по стволу скважины по ниже НКТ и по затрубу до устья
    Public Const str_TcasHmes_curve = "TcasHmes_curve"           ' профиль температуры по стволу скважины ниже насоса и выше насоса по затрубу
    Public Const str_RstubHmes_curve = "RstubHmes_curve"         ' профиль остаточного содержания газа в нефти по потоку в НКТ
    Public Const str_RscasHmes_curve = "RscasHmes_curve"         ' профиль остаточного содержания газа в нефти по потоку по затрубу
    Public Const str_GasFracTubHmes_curve = "GasFracTubHmes_curve" ' расходное содержание газа в потоке в НКТ
    Public Const str_GasFracCasHmes_curve = "GasFracCasHmes_curve" ' расходное содержание газа в потоке по затрубу
    Public Const str_HlHmes_curve = "HlHmes_curve"         ' Liquid holdup (содержание жидкости) в потоке через НКТ
    Public Const str_HLtubHmes_curve = "HLtubHmes_curve"         ' Liquid holdup (содержание жидкости) в потоке через НКТ
    Public Const str_HLcasHmes_curve = "HLcasHmes_curve"         ' Liquid holdup (содержание жидкости) в потоке по затрубу
    Public Const str_muoTubCurve = "muoTubCurve" '
    Public Const str_muwTubCurve = "muwTubCurve" '
    Public Const str_mugTubCurve = "mugTubCurve" '
    Public Const str_mumixTubCurve = "mumixTubCurve" '
    Public Const str_rhooTubCurve = "rhooTubCurve" '
    Public Const str_rhowTubCurve = "rhowTubCurve" '
    Public Const str_rholTubCurve = "rholTubCurve" '
    Public Const str_rhogTubCurve = "rhogTubCurve" '
    Public Const str_rhomixTubCurve = "rhomixTubCurve" '
    Public Const str_qoTubCurve = "qoTubCurve" '
    Public Const str_qwTubCurve = "qwTubCurve" '
    Public Const str_qgTubCurve = "qgTubCurve" '
    Public Const str_moTubCurve = "moTubCurve" '
    Public Const str_mwTubCurve = "mwTubCurve" '
    Public Const str_mgTubCurve = "mgTubCurve" '
    Public Const str_vlTubCurve = "vlTubCurve" '
    Public Const str_vgTubCurve = "vgTubCurve" '
    Public Const str_muoCasCurve = "muoCasCurve" '
    Public Const str_muwCasCurve = "muwCasCurve" '
    Public Const str_mugCasCurve = "mugCasCurve" '
    Public Const str_mumixCasCurve = "mumixCasCurve" '
    Public Const str_rhooCasCurve = "rhooCasCurve" '
    Public Const str_rhowCasCurve = "rhowCasCurve" '
    Public Const str_rholCasCurve = "rholCasCurve" '
    Public Const str_rhogCasCurve = "rhogCasCurve" '
    Public Const str_rhomixCasCurve = "rhomixCasCurve" '
    Public Const str_qoCasCurve = "qoCasCurve" '
    Public Const str_qwCasCurve = "qwCasCurve" 'a's
    Public Const str_qgCasCurve = "qgCasCurve" '
    Public Const str_moCasCurve = "moCasCurve" '
    Public Const str_mwCasCurve = "mwCasCurve" '
    Public Const str_mgCasCurve = "mgCasCurve" '
    Public Const str_vlCasCurve = "vlCasCurve" '
    Public Const str_vgCasCurve = "vgCasCurve" '

    Public Const PVT_DEFAULT = "gamma_gas:0,900;gamma_oil:0,750;gamma_wat:1,000;rsb_m3m3:100,000;rp_m3m3:-1,000;pb_atma:-1,000;tres_C:90,000;bob_m3m3:-1,000;muob_cP:-1,000;PVTcorr:0;ksep_fr:0,000;p_ksep_atma:-1,000;t_ksep_C:-1,000;"
    Public Const ESP_DEFAULT = "ESP_ID:1006.00000;HeadNom_m:2000.00000;ESPfreq_Hz:50.00000;ESP_U_V:1000.00000;MotorPowerNom_kW:30.00000;t_int_C:85.00000;t_dis_C:25.00000;KsepGS_fr:0.00000;ESP_energy_fact_Whday:0.00000;ESP_cable_type:0;ESP_h_mes_m:0.00000;ESP_gas_correct:0;c_calibr_head:0.00000;PKV_work_min:-1,00000;PKV_stop_min:-1,00000;"
    Public Const WELL_DEFAULT = "h_perf_m:2000,00000;h_pump_m:1800,00000;udl_m:0,00000;d_cas_mm:150,00000;dtub_mm:72,00000;dchoke_mm:15,00000;roughness_m:0,00010;t_bh_C:85,00000;t_wh_C:25,00000;"
    Public Const WELL_GL_DEFAULT = "h_perf_m:2500,00000;htub_m:2000,00000;udl_m:0,00000;d_cas_mm:125,00000;dtub_mm:62,00000;dchoke_mm:15,00000;roughness_m:0,00010;t_bh_C:100,00000;t_wh_C:50,00000;GLV:1;H_glv_m:1500,000;d_glv_mm:5,000;p_glv_atma:50,000;"

    Public Const const_t_K_min = 273         ' ниже нуля ничего не считаем?
    Public Const const_tMcCain_K_min = 289         ' ниже нуля ничего не считаем?
    Public Const const_t_K_max = 573         ' выше тоже ничего не считаем?
    Public Const const_t_K_zero_C = 273
    Public Const const_t_C_min = const_t_K_min - const_t_K_zero_C
    Public Const const_t_C_max = const_t_K_max - const_t_K_min
    Public Const const_Pi As Double = 3.14159265358979
    Public Const const_tsc_C = 20
    Public Const const_tsc_K As Double = const_tsc_C + const_t_K_zero_C ' температура стандартных условиях, К
    Public Const const_psc_atma As Double = 1
    Public Const const_r As Double = 8.31 'Universal gas constant
    Public Const const_g = 9.81
    Public Const const_rho_air = 1.2217
    Public Const const_gamma_w = 1
    Public Const const_rho_ref = 1000
    Public Const const_ZNLF_rate = 0.1
    Public Const const_m_a_kgmol As Double = 0.029 'Air molar mass
    Public Const const_sigma_wat_gas_Nm = 0.01 ' поверхностное натяжение на границе с воздухом (газом) - типовые значения для дефолтных параметров  Н/м
    Public Const const_sigma_oil_Nm = 0.025
    Public Const const_mu_w = 0.36
    Public Const const_mu_g = 0.0122
    Public Const const_mu_o = 0.7
    Public Const const_gg_ = 0.6
    Public Const const_gw_ = 1
    Public Const const_go_ = 0.86
    Public Const const_rsb_default = 100
    Public Const const_Bob_default = 1.2
    Public Const const_tres_default = 90
    Public Const const_Roughness_default = 0.0001

    ' набор констант для общих ограничений значений переменных
    Public Const const_gamma_gas_min = 0.5   ' плотность метана 0.59 - предпологаем легче газов не будет
    Public Const const_gamma_gas_max = 2     ' плотность углеводородных газов (гексан) может доходить до 4, но мы считаем что в смеси таких не много должно быть
    Public Const const_gamma_water_min = 0.9 ' плотность воды от 0.9 до 1.5
    Public Const const_gamma_water_max = 1.5
    Public Const const_gamma_oil_min = 0.5   ' плотность нефти
    Public Const const_gamma_oil_max = 1.5

    Public Const const_P_MPa_min = 0
    Public Const const_P_MPa_max = 50
    Public Const const_Salinity_ppm_min = 0
    Public Const const_Salinity_ppm_max = 265000  ' equal to weigh percent salinity 26.5%.  Ограничение по границам применимости корреляций МакКейна
    Public Const const_rsb_m3m3_min = 0
    Public Const const_rsb_m3m3_max = 100000 ' rsb more that 100 000 not allowed
    Public Const const_Ppr_min = 0.002
    Public Const const_Ppr_max = 30
    Public Const const_Tpr_min = 0.7
    Public Const const_Tpr_max = 3
    Public Const const_Z_min = 0.05
    Public Const const_Z_max = 5
    Public Const const_TGeoGrad_C100m = 3   ' геотермальный градиент в градусах на 100 м
    Public Const const_Heps_m = 0.001       ' дельта для корретировки кривой трубы, - примерно соответствует длине сочленения труб
    Public Const const_ESP_length = 1      ' длина УЭЦН по умолчанию
    Public Const const_pipe_diam_default_mm = 62

    ' набор констант для перевода единиц измерений в различных размерностях
    Public Const const_convert_atma_Pa = 101325
    Public Const const_convert_Pa_atma = 1 / const_convert_atma_Pa
    Public Const const_convert_kgfcm2_Pa = 98066.5
    Public Const const_convert_m3day_bbl = 6.289810569
    Public Const const_convert_gpm_m3day = 5.450992992     ' (US) gallon per minute
    Public Const const_convert_m3day_gpm = 1 / const_convert_gpm_m3day
    Public Const const_convert_m3m3_scfbbl = 5.614583544
    Public Const const_convert_scfbbl_m3m3 = 1 / const_convert_m3m3_scfbbl
    Public Const const_convert_bbl_m3day = 1 / const_convert_m3day_bbl
    Public Const const_conver_day_sec = 86400   ' updated for test  rnt21
    Public Const const_convert_hr_sec = 3600
    Public Const const_convert_m3day_m3sec = 1 / const_conver_day_sec
    Public Const const_conver_sec_day = 1 / const_conver_day_sec
    Public Const const_convert_atma_psi = 14.7
    Public Const const_convert_psi_atma = 1 / const_convert_atma_psi
    Public Const const_convert_ft_m = 0.3048
    Public Const const_convert_m_ft = 1 / const_convert_ft_m
    Public Const const_convert_m_mm = 1000
    Public Const const_convert_mm_m = 1 / const_convert_m_mm
    Public Const const_convert_cP_Pasec = 1 / 1000
    Public Const const_convert_HP_W = 745.69987  ' 735.49875  ' метрическая лошадиная сила. следует учесть, что иногда может применяться механическя лошадиная сила (1.013 метрической)
    Public Const const_convert_W_HP = 1 / const_convert_HP_W
    Public Const const_convert_Nm_dynescm = 1000
    Public Const const_convert_lbmft3_kgm3 = 16.01846
    Public Const const_convert_kgm3_lbmft3 = 1 / const_convert_lbmft3_kgm3
    Public Const const_convert_psift_atmm = 1 / const_convert_atma_psi / const_convert_ft_m ' pressure gradient conversion factor
    Public Const const_convert_MPa_atma = 1000000 / const_convert_atma_Pa  ' 9.8692
    Public Const const_convert_atma_MPa = 1 / const_convert_MPa_atma ' 0.101325' константа для конверсии единиц давления из Мпа в atma
    Public Const const_p_atma_min = const_P_MPa_min * const_convert_MPa_atma
    Public Const const_p_atma_max = const_P_MPa_max * const_convert_MPa_atma
    Public Const MAXIT = 100
    ' константы для расчета многофазного потока
    Public Const const_MaxSegmLen = 100
    Public Const const_n_n = 20
    Public Const const_MaxdP = 10
    Public Const const_minPpipe_atma = 0.9
    Public Const const_pressure_tolerance = 0.001
    Public Const const_well_P_tolerance = 0.05     ' допустимая погрешность при расчете забойного давления в скважине
    Public Const const_P_difference = 0.0001       ' допустимая погрешность при сравнении (в основном) давлений
    Public Const ang_max = 5
    Public Const const_OutputCurveNumPoints = 50
    Public Const DEFAULT_PAN_STEP = 15

    Public Const kErrWellConstruction = 513 + vbObjectError
    Public Const kErrPVTinput = 514 + vbObjectError
    Public Const kErrNodalCalc = 515 + vbObjectError
    Public Const kErrInitCalc = 516 + vbObjectError
    Public Const kErrESPbase = 517 + vbObjectError
    Public Const kErrPVTcalc = 518 + vbObjectError
    Public Const kErrESPcalc = 519 + vbObjectError
    Public Const kErrGradcalc = 520 + vbObjectError
    Public Const kErrArraySize = 701 + vbObjectError
    Public Const kErrBuildCurve = 702 + vbObjectError
    Public Const kErrcurvestablePointIndex = 703 + vbObjectError
    Public Const kErrCurvePointIndex = 704 + vbObjectError
    Public Const kErrReadDataFromWorksheet = 705 + vbObjectError
    Public Const kErrWriteDataFromWorksheet = 706 + vbObjectError
    Public Const kStrConversion = 707 + vbObjectError
    Public Const kErrDegradationNotFound = 708 + vbObjectError
    Public Const kErrDegradationError = 709 + vbObjectError
    Public Const kreadRangeError = 710 + vbObjectError
    Public Const kErrCInterpolation = 711 + vbObjectError
    Public Const kErrTester = 712 + vbObjectError
    Public Const kErrBisection = 713 + vbObjectError

    Public Const sDELIM As String = vbLf & vbNewLine
    Public Const MinCountPoints_calc_pwf_pcas_hdyn_atma = 5




End Module

'=======================================================================================
'Unifloc 7.24  coronav                                     khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
'PVT UDF  (user defined functions for PVT calculation)
Option Explicit On
Imports ExcelDna.Integration

Public Module u7PVT


    <ExcelFunction(Description:="My first .NET function")>
    Public Function HelloDna(name As String, arr As Double(,)) As Double(,)
        arr(0, 0) = 23
        Return arr
    End Function

    ' функция вывода номера версии унифлок для расчетных модулей
    'Public Function getUFVersion()
    '    getUFVersion = const_unifloc_version
    'End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' функция расчета объемного коэффициента газа
    Public Function PVT_bg_m3m3(
                ByVal p_atma As Integer,
                ByVal t_C As Integer,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal tres_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре tres_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' tres_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' Возвращает значение объемного коэффициента газа, м3/м3
        ' для заданных термобарических условий.
        ' В основе расчета корреляция для z факотора
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = ReadPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, tres_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_bg_m3m3 = PVT.Bg_m3m3

        Catch ex As Exception
            PVT_bg_m3m3 = -1
            '"Error:PVT_bg_m3m3:"
        End Try
    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет объемного коэффициента нефти
    Public Function PVT_bo_m3m3(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число
        ' Возвращает значение объемного коэффициента нефти, м3/м3
        ' для заданных термобарических условий.
        ' В основе расчета корреляции PVT
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_bo_m3m3 = PVT.Bo_m3m3
            Exit Function
        Catch ex As Exception
            PVT_bo_m3m3 = -1
            '"Error:PVT_bo_m3m3:"
        End Try
    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет объемного коэффициента воды
    Public Function PVT_bw_m3m3(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число
        ' Возвращает значение объемного коэффициента воды, м3/м3
        ' для заданных термобарических условий.
        ' В основе расчета корреляции PVT
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_bw_m3m3 = PVT.Bw_m3m3
            Exit Function
        Catch ex As Exception
            PVT_bw_m3m3 = -1
            '"Error:PVT_bw_m3m3:" 
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет солености воды
    Public Function PVT_salinity_ppm(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число
        ' Возвращает соленость воды, ppm
        ' для заданных термобарических условий.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_salinity_ppm = PVT.Sal_ppm
            Exit Function
        Catch ex As Exception
            PVT_salinity_ppm = -1
            '"Error:PVT_salinity_ppm:"
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет вязкости нефти
    Public Function PVT_mu_oil_cP(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число - вязкость нефти
        '           при заданных термобарических условиях, сП
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_mu_oil_cP = PVT.Mu_oil_cP
            Exit Function
        Catch ex As Exception
            PVT_mu_oil_cP = -1
            ' "Error:PVT_mu_oil_cP:" 
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет вязкости газа
    Public Function PVT_mu_gas_cP(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число - вязкость газа
        '           при заданных термобарических условиях, сП
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_mu_gas_cP = PVT.Mu_gas_cP
            Exit Function
        Catch ex As Exception
            PVT_mu_gas_cP = -1
            '"Error:PVT_mu_gas_cP:"
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет вязкости воды
    Public Function PVT_mu_wat_cP(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число - вязкость воды
        '           при заданных термобарических условиях, сП
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_mu_wat_cP = PVT.Mu_wat_cP
            Exit Function
        Catch ex As Exception
            PVT_mu_wat_cP = -1
            ' "Error:PVT_mu_wat_cP:"
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет газосодержания
    Public Function PVT_rs_m3m3(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число - газосодержание при
        '           заданных термобарических условиях, м3/м3.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_rs_m3m3 = PVT.Rs_m3m3
            Exit Function
        Catch ex As Exception
            PVT_rs_m3m3 = -1
            '"Error:PVT_Rs_m3m3:"
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет коэффициента сверхсжимаемости газа
    Public Function PVT_z(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число - z фактор газа.
        '           коэффициент сверхсжимаемости газа,
        '           безразмерная величина
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_z = PVT.Z
            Exit Function
        Catch ex As Exception
            PVT_z = -1
            '"Error:PVT_Z:" 
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет плотности нефти в рабочих условиях
    Public Function PVT_rho_oil_kgm3(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число - плотность нефти
        '           при заданных термобарических условиях, кг/м3.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_rho_oil_kgm3 = PVT.Rho_oil_rc_kgm3
            Exit Function
        Catch ex As Exception
            PVT_rho_oil_kgm3 = -1
            ' "Error:PVT_rho_oil_kgm3:"
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет плотности газа в рабочих условиях
    Public Function PVT_rho_gas_kgm3(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число - плотность газа
        '           при заданных термобарических условиях, кг/м3.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_rho_gas_kgm3 = PVT.Rho_gas_rc_kgm3
            Exit Function
        Catch ex As Exception
            PVT_rho_gas_kgm3 = -1
            '"Error:PVT_rho_gas_kgm3:"
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет плотности воды в рабочих условиях
    Public Function PVT_rho_wat_kgm3(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число - плотность воды
        '           при заданных термобарических условиях, кг/м3.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_rho_wat_kgm3 = PVT.Rho_wat_rc_kgm3
            Exit Function
        Catch ex As Exception
            PVT_rho_wat_kgm3 = -1
            '"Error:PVT_rho_wat_kgm3:"
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' Расчет давления насыщения
    Public Function PVT_pb_atma(
                     ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число - давление насыщения.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            PVT_pb_atma = PVT.Calc_pb_atma(rsb_m3m3, t_C)
            Exit Function
        Catch ex As Exception
            PVT_pb_atma = -1
            '"Error:PVT_pb_atma:"
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет коэффициента поверхностного натяжения нефть - газ
    Public Function PVT_ST_oilgas_Nm(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число
        ' Возвращает коэффициента поверхностного натяжения нефть - газ, Нм
        ' для заданных термобарических условий.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_ST_oilgas_Nm = PVT.Sigma_oil_gas_Nm
            Exit Function
        Catch ex As Exception
            PVT_ST_oilgas_Nm = -1
            '"Error:PVT_ST_oilgas_Nm:"
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет коэффициента поверхностного натяжения вода - газ
    Public Function PVT_ST_watgas_Nm(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число
        ' Возвращает коэффициента поверхностного натяжения вода - газ, Нм
        ' для заданных термобарических условий.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_ST_watgas_Nm = PVT.Sigma_wat_gas_Nm
            Exit Function
        Catch ex As Exception
            PVT_ST_watgas_Nm = -1
            ' "Error:PVT_ST_watgas_Nm:"
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет коэффициента поверхностного натяжения жидкость - газ
    Public Function PVT_ST_liqgas_Nm(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число
        ' Возвращает коэффициента поверхностного натяжения жидкость - газ, Нм
        ' для заданных термобарических условий.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_ST_liqgas_Nm = PVT.Sigma_liq_Nm
            Exit Function
        Catch ex As Exception
            PVT_ST_liqgas_Nm = -1
            ' "Error:PVT_ST_liqgas_Nm:"
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет теплоемкости нефти при постоянном давлении cp
    Public Function PVT_cp_oil_JkgC(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число
        ' Возвращает коэффициента поверхностного натяжения жидкость - газ, Нм
        ' для заданных термобарических условий.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_cp_oil_JkgC = PVT.Cp_oil_JkgC
            Exit Function
        Catch ex As Exception
            PVT_cp_oil_JkgC = -1
            '"Error:PVT_cp_oil_JkgC:" 
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет теплоемкости газа при постоянном давлении cp
    Public Function PVT_cp_gas_JkgC(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число
        ' Возвращает коэффициента поверхностного натяжения жидкость - газ, Нм
        ' для заданных термобарических условий.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_cp_gas_JkgC = PVT.Cp_gas_JkgC
            Exit Function
        Catch ex As Exception
            PVT_cp_gas_JkgC = -1
            '"Error:PVT_cp_gas_JkgC:"
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет теплоемкости газа при постоянном давлении cp
    Public Function PVT_cv_gas_JkgC(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число
        ' Возвращает коэффициента поверхностного натяжения жидкость - газ, Нм
        ' для заданных термобарических условий.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_cv_gas_JkgC = PVT.Cv_gas_JkgC
            Exit Function
        Catch ex As Exception
            PVT_cv_gas_JkgC = -1
            '"Error:PVT_cv_gas_JkgC:"
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет теплоемкости воды при постоянном давлении cp
    Public Function PVT_cp_wat_JkgC(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число
        ' Возвращает коэффициента поверхностного натяжения жидкость - газ, Нм
        ' для заданных термобарических условий.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_cp_wat_JkgC = PVT.Cp_wat_JkgC
            Exit Function
        Catch ex As Exception
            PVT_cp_wat_JkgC = -1
            '"Error:PVT_cp_wat_JkgC:"
        End Try

    End Function


    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет сжимаемости воды
    Public Function PVT_compressibility_wat_1atm(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число
        ' Возвращает коэффициента поверхностного натяжения жидкость - газ, Нм
        ' для заданных термобарических условий.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_compressibility_wat_1atm = PVT.Compressibility_wat_1atm
            Exit Function
        Catch ex As Exception
            PVT_compressibility_wat_1atm = -1
            ' "Error:PVT_compressibility_wat_1atm:"
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет сжимаемости нефти
    Public Function PVT_compressibility_oil_1atm(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число
        ' Возвращает коэффициента поверхностного натяжения жидкость - газ, Нм
        ' для заданных термобарических условий.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_compressibility_oil_1atm = PVT.Compressibility_oil_1atm
            Exit Function
        Catch ex As Exception
            PVT_compressibility_oil_1atm = -1
            ' "Error:PVT_compressibility_oil_1atm:" 
        End Try

    End Function


    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет сжимаемости нефти
    Public Function PVT_compressibility_gas_1atm(
                ByVal p_atma As Double,
                ByVal t_C As Double,
                Optional ByVal gamma_gas As Double = 0.7,
                Optional ByVal gamma_oil As Double = 0.86,
                Optional ByVal gamma_wat As Double = 1.02,
                Optional ByVal rsb_m3m3 As Double = 100,
                Optional ByVal rp_m3m3 As Double = -1,
                Optional ByVal pb_atma As Double = -1,
                Optional ByVal t_res_C As Double = 30,
                Optional ByVal bob_m3m3 As Double = -1,
                Optional ByVal muob_cP As Double = -1,
                Optional ByVal PVTcorr As Integer = 0,
                Optional ByVal ksep_fr As Double = 0,
                Optional ByVal p_ksep_atma As Double = -1,
                Optional ByVal t_ksep_C As Double = -1,
                Optional ByVal str_PVT As String = "")
        ' p_atma  - давление, атм
        ' t_C     - температура, С.
        ' gamma_gas - удельная плотность газа, по воздуху.
        '           const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '           const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '           const_gw_ = 1
        ' rsb_m3m3 -  газосодержание при давлении насыщения, м3/м3.
        '           const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           имеет приоритет перед rsb если Rp < rsb
        ' pb_atma - Давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0 то рассчитается по корреляции
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти, м3/м3.
        ' muob_cP - вязкость нефти при давлении насыщения
        '           По умолчанию рассчитывается по корреляции
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           Standing_based = 0 - на основе кор-ии Стендинга
        '           McCain_based = 1 - на основе кор-ии Маккейна
        '           straigth_line = 2 - на основе упрощенных зависимостей
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации доли свободного газа.
        '           изменение свойств нефти зависит от условий
        '           сепарации газа, которые должны быть явно заданы
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' str_PVT - закодированная строка с параметрами PVT.
        '           Если задана - перекрывает другие значения
        '
        ' результат - число
        ' Возвращает коэффициента поверхностного натяжения жидкость - газ, Нм
        ' для заданных термобарических условий.
        'description_end

        Try
            Dim PVT As UnfClassLibrary.CPVT
            PVT = readPVT(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, rp_m3m3, pb_atma, t_res_C, bob_m3m3,
                        muob_cP, PVTcorr, ksep_fr, p_ksep_atma, t_ksep_C, str_PVT)
            Call PVT.Calc_PVT(p_atma, t_C)
            PVT_compressibility_gas_1atm = PVT.Compressibility_gas_1atm
            Exit Function
        Catch ex As Exception
            PVT_compressibility_gas_1atm = -1
            '"Error:PVT_compressibility_gas_1atm:"
        End Try

    End Function


    ' вспомогательная функция для создания PVT объекта из исходных данных
    Private Function ReadPVT(Optional ByVal gamma_gas As Double = 0.7,
                    Optional ByVal gamma_oil As Double = 0.86,
                    Optional ByVal gamma_wat As Double = 1.02,
                    Optional ByVal rsb_m3m3 As Double = 100,
                    Optional ByVal rp_m3m3 As Double = -1,
                    Optional ByVal pb_atma As Double = -1,
                    Optional ByVal tres_C As Double = 50,
                    Optional ByVal bob_m3m3 As Double = -1,
                    Optional ByVal muob_cP As Double = -1,
                    Optional ByVal PVTcorr As Integer = 0,
                    Optional ByVal ksep_fr As Double = 0,
                    Optional ByVal p_ksep_atma As Double = -1,
                    Optional ByVal t_ksep_C As Double = -1,
                    Optional ByVal str_PVT As String = ""
                    ) As UnfClassLibrary.CPVT

        Dim PVT As UnfClassLibrary.CPVT
        If str_PVT <> "" Then
            PVT = PVT_decode_string(str_PVT)
        Else
            PVT = New UnfClassLibrary.CPVT
            PVT.Class_Initialize()
            PVT.Init(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, pb_atma, bob_m3m3, PVTcorr, tres_C, rp_m3m3, muob_cP)
            If ksep_fr > 0 And ksep_fr <= 1 And p_ksep_atma > 0 And t_ksep_C > 0 Then
                Call PVT.Mod_after_separation(p_ksep_atma, t_ksep_C, ksep_fr, True)
            End If
        End If

        ReadPVT = PVT
    End Function

End Module

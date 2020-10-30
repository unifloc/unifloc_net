'=======================================================================================
'Unifloc 7.24  coronav                                     khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
'PVT UDF  (user defined functions for PVT calculation)

Imports ExcelDna.Integration




Public Module u7PVT


    <ExcelFunction(Description:="My first .NET function")>
    Public Function HelloDna(name As String, arr As Double(,)) As Double(,)
        arr(0, 0) = 23
        Return arr
    End Function

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
                Optional ByVal t_ksep_C As Double = -1)
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
            PVT_bg_m3m3 = {PVT.Bg_m3m3, PVT.Bo_m3m3, gamma_oil}

        Catch ex As Exception
            PVT_bg_m3m3 = -1


        End Try
    End Function


    ' функция вывода номера версии унифлок для расчетных модулей
    'Public Function getUFVersion()
    '    getUFVersion = const_unifloc_version
    'End Function

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
                    Optional ByVal t_ksep_C As Double = -1
                    ) As UnfClassLibrary.CPVT

        Dim PVT As UnfClassLibrary.CPVT

        PVT = New UnfClassLibrary.CPVT
        PVT.Init(gamma_gas, gamma_oil, gamma_wat, rsb_m3m3, pb_atma, bob_m3m3, PVTcorr, tres_C, rp_m3m3, muob_cP)
        If ksep_fr > 0 And ksep_fr <= 1 And p_ksep_atma > 0 And t_ksep_C > 0 Then
            Call PVT.Mod_after_separation(p_ksep_atma, t_ksep_C, ksep_fr, True)
        End If

        ReadPVT = PVT
    End Function

End Module

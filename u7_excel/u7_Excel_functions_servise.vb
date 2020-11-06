'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
' вспомогательные функции для проведения расчетов из рабочих книг Excel

Option Explicit On
Imports Newtonsoft.Json
Public Module u7_Excel_function_servise
    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' функция кодирования параметров PVT в строку,
    ' для передачи PVT свойств в прикладные функции Унифлок.
    Public Function PVT_encode_string(
                    Optional ByVal gamma_gas As Double = UnfClassLibrary.u7_const.const_gg_,
                    Optional ByVal gamma_oil As Double = UnfClassLibrary.u7_const.const_go_,
                    Optional ByVal gamma_wat As Double = UnfClassLibrary.u7_const.const_gw_,
                    Optional ByVal rsb_m3m3 As Double = UnfClassLibrary.u7_const.const_rsb_default,
                    Optional ByVal rp_m3m3 As Double = -1,
                    Optional ByVal pb_atma As Double = -1,
                    Optional ByVal t_res_C As Double = 90,
                    Optional ByVal bob_m3m3 As Double = -1,
                    Optional ByVal muob_cP As Double = -1,
                    Optional ByVal PVTcorr As Integer = 0,
                    Optional ByVal ksep_fr As Double = -1,
                    Optional ByVal p_ksep_atma As Double = -1,
                    Optional ByVal t_ksep_C As Double = -1,
                    Optional ByVal gas_only As Boolean = False
                    )
        ' gamma_gas - удельная плотность газа, по воздуху.
        '             По умолчанию const_gg_ = 0.6
        ' gamma_oil - удельная плотность нефти, по воде.
        '             По умолчанию const_go_ = 0.86
        ' gamma_wat - удельная плотность воды, по воде.
        '             По умолчанию const_gw_ = 1
        ' rsb_m3m3  - газосодержание при давлении насыщения, м3/м3.
        '             По умолчанию const_rsb_default = 100
        ' rp_m3m3 - замерной газовый фактор, м3/м3.
        '           Имеет приоритет перед rsb если rp < rsb
        ' pb_atma - давление насыщения при  температуре t_res_C, атма.
        '           Опциональный калибровочный параметр,
        '           если не задан или = 0, то рассчитается по корреляции.
        ' t_res_C  - пластовая температура, С.
        '           Учитывается при расчете давления насыщения.
        '           По умолчанию  const_tres_default = 90
        ' bob_m3m3 - объемный коэффициент нефти при давлении насыщения
        '            и пластовой температуре, м3/м3.
        '            По умолчанию рассчитывается по корреляции.
        ' muob_cP  - вязкость нефти при давлении насыщения.
        '            и пластовой температуре, сП.
        '            По умолчанию рассчитывается по корреляции.
        ' PVTcorr - номер набора PVT корреляций для расчета:
        '           0 - на основе корреляции Стендинга;
        '           1 - на основе кор-ии Маккейна;
        '           2 - на основе упрощенных зависимостей.
        ' ksep_fr - коэффициент сепарации - определяет изменение свойств
        '           нефти после сепарации части свободного газа.
        '           Зависит от давления и температуры
        '           сепарации газа, которые должны быть явно заданы.
        ' p_ksep_atma - давление при которой была сепарация
        ' t_ksep_C    - температура при которой была сепарация
        ' gas_only   - флаг - в потоке только газ
        '              по умолчанию False (нефть вода и газ)
        ' результат - закодированная строка
        'description_end

        Dim str As String
        Dim frmt As String
        Dim frmt_int As String

        Dim pvt_dict As New Dictionary(Of String, Double)

        pvt_dict.Add("gamma_gas", gamma_gas)
        pvt_dict.Add("gamma_oil", gamma_oil)
        pvt_dict.Add("gamma_wat", gamma_wat)
        pvt_dict.Add("rsb_m3m3", rsb_m3m3)

        If rp_m3m3 Then pvt_dict.Add("rp_m3m3", rp_m3m3)
        If pb_atma Then pvt_dict.Add("pb_atma", pb_atma)
        If t_res_C Then pvt_dict.Add("t_res_C", t_res_C)
        If bob_m3m3 Then pvt_dict.Add("bob_m3m3", bob_m3m3)
        If muob_cP Then pvt_dict.Add("muob_cP", muob_cP)
        If PVTcorr Then pvt_dict.Add("PVTcorr", PVTcorr)
        If ksep_fr Then pvt_dict.Add("ksep_fr", ksep_fr)
        If p_ksep_atma Then pvt_dict.Add("p_ksep_atma", p_ksep_atma)
        If t_ksep_C Then pvt_dict.Add("t_ksep_C", t_ksep_C)
        If gas_only Then pvt_dict.Add("gas_only", gas_only)

        Dim new_json_ = JsonConvert.SerializeObject(pvt_dict)
        PVT_encode_string = new_json_
        'Debug.Print(PVT_encode_string)
    End Function

    Public Function PVT_decode_string(
                    Optional ByVal str_PVT As String = UnfClassLibrary.u7_const.PVT_DEFAULT,
                    Optional ByVal getStr As Boolean = False)
        Dim PVT As New UnfClassLibrary.CPVT
        Try
            PVT.Class_Initialize()
            If Len(str_PVT) < 3 Then
                PVT_decode_string = Nothing
                Exit Function
            End If

            Call PVT.init_json(str_PVT)
            If getStr Then
                With PVT
                    PVT_decode_string = PVT_encode_string(.gamma_g, .gamma_o,
                                                            .gamma_w, .Rsb_m3m3, .Rsb_m3m3,
                                                            .pb_atma, .tres_C, .bob_m3m3, .muob_cP, 0, .ksep_fr, .p_ksep_atma,
                                                            .t_ksep_C, PVT.gas_only)
                End With
            Else
                PVT_decode_string = PVT
            End If
            Exit Function
        Catch ex As Exception
            Dim errmsg As String
            errmsg = "Error:PVT_decode_string"
            Throw New ApplicationException(errmsg)
        End Try
    End Function


End Module

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
        ' gamma_gas - óäåëüíàÿ ïëîòíîñòü ãàçà, ïî âîçäóõó.
        '             Ïî óìîë÷àíèþ const_gg_ = 0.6
        ' gamma_oil - óäåëüíàÿ ïëîòíîñòü íåôòè, ïî âîäå.
        '             Ïî óìîë÷àíèþ const_go_ = 0.86
        ' gamma_wat - óäåëüíàÿ ïëîòíîñòü âîäû, ïî âîäå.
        '             Ïî óìîë÷àíèþ const_gw_ = 1
        ' rsb_m3m3  - ãàçîñîäåðæàíèå ïðè äàâëåíèè íàñûùåíèÿ, ì3/ì3.
        '             Ïî óìîë÷àíèþ const_rsb_default = 100
        ' rp_m3m3 - çàìåðíîé ãàçîâûé ôàêòîð, ì3/ì3.
        '           Èìååò ïðèîðèòåò ïåðåä rsb åñëè rp < rsb
        ' pb_atma - äàâëåíèå íàñûùåíèÿ ïðè  òåìïåðàòóðå t_res_C, àòìà.
        '           Îïöèîíàëüíûé êàëèáðîâî÷íûé ïàðàìåòð,
        '           åñëè íå çàäàí èëè = 0, òî ðàññ÷èòàåòñÿ ïî êîððåëÿöèè.
        ' t_res_C  - ïëàñòîâàÿ òåìïåðàòóðà, Ñ.
        '           Ó÷èòûâàåòñÿ ïðè ðàñ÷åòå äàâëåíèÿ íàñûùåíèÿ.
        '           Ïî óìîë÷àíèþ  const_tres_default = 90
        ' bob_m3m3 - îáúåìíûé êîýôôèöèåíò íåôòè ïðè äàâëåíèè íàñûùåíèÿ
        '            è ïëàñòîâîé òåìïåðàòóðå, ì3/ì3.
        '            Ïî óìîë÷àíèþ ðàññ÷èòûâàåòñÿ ïî êîððåëÿöèè.
        ' muob_cP  - âÿçêîñòü íåôòè ïðè äàâëåíèè íàñûùåíèÿ.
        '            è ïëàñòîâîé òåìïåðàòóðå, ñÏ.
        '            Ïî óìîë÷àíèþ ðàññ÷èòûâàåòñÿ ïî êîððåëÿöèè.
        ' PVTcorr - íîìåð íàáîðà PVT êîððåëÿöèé äëÿ ðàñ÷åòà:
        '           0 - íà îñíîâå êîððåëÿöèè Ñòåíäèíãà;
        '           1 - íà îñíîâå êîð-èè Ìàêêåéíà;
        '           2 - íà îñíîâå óïðîùåííûõ çàâèñèìîñòåé.
        ' ksep_fr - êîýôôèöèåíò ñåïàðàöèè - îïðåäåëÿåò èçìåíåíèå ñâîéñòâ
        '           íåôòè ïîñëå ñåïàðàöèè ÷àñòè ñâîáîäíîãî ãàçà.
        '           Çàâèñèò îò äàâëåíèÿ è òåìïåðàòóðû
        '           ñåïàðàöèè ãàçà, êîòîðûå äîëæíû áûòü ÿâíî çàäàíû.
        ' p_ksep_atma - äàâëåíèå ïðè êîòîðîé áûëà ñåïàðàöèÿ
        ' t_ksep_C    - òåìïåðàòóðà ïðè êîòîðîé áûëà ñåïàðàöèÿ
        ' gas_only   - ôëàã - â ïîòîêå òîëüêî ãàç
        '              ïî óìîë÷àíèþ False (íåôòü âîäà è ãàç)
        ' ðåçóëüòàò - çàêîäèðîâàííàÿ ñòðîêà
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

'=======================================================================================
'Unifloc 7.25  coronav                                     khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
' базовые функции для проведения расчетов из интерфейса Excel


Option Explicit On

Public Module u7_Excel_functions_IPR
    ' ==============  функции для расчета пласта ==========================
    ' =====================================================================

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет дебита по давлению и продуктивности
    Public Function IPR_qliq_sm3day(
                 ByVal pi_sm3dayatm As Double,
                 ByVal pres_atma As Double,
                 ByVal pwf_atma As Double,
        Optional ByVal fw_perc As Double = 0,
        Optional ByVal pb_atma As Double = -1)
        ' pi_sm3dayatm   - коэффициент продуктивности, ст.м3/сут/атм
        ' Pres_atma      - пластовое давление, абс. атм
        ' pwf_atma       - забойное давление, абс. атм
        ' fw_perc        - обводненность, %
        ' pb_atma        - давление насыщения, абс. атм
        ' результат      - значение дебита жидкости, ст.м3/сут
        'description_end

        Try
            Dim res As New UnfClassLibrary.CReservoirVogel
            If pb_atma <= 0 Then pb_atma = 0   ' поставим ноль иначе флюид додсчитает по корреляции значение
            res.InitProp(pres_atma, pb_atma, fw_perc)
            res.pi_sm3dayatm = pi_sm3dayatm

            IPR_qliq_sm3day = res.calc_qliq_sm3day(pwf_atma)
            res = Nothing


            Exit Function
        Catch ex As Exception
            IPR_qliq_sm3day = -1
            Dim msg As String
            msg = "Error:IPR_qliq_sm3day:" & Err.Description

            Throw New ApplicationException(msg)
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет забойного давления по дебиту и продуктивности
    Public Function IPR_pwf_atma(
                 ByVal pi_sm3dayatm As Double,
                 ByVal pres_atma As Double,
                 ByVal qliq_sm3day As Double,
        Optional ByVal fw_perc As Double = 0,
        Optional ByVal pb_atma As Double = -1)
        ' pi_sm3dayatm   - коэффициент продуктивности, ст.м3/сут/атм
        ' Pres_atma      - пластовое давление, абс. атм
        ' qliq_sm3day    - дебит жидкости скважины на поверхности, ст.м3/сут
        ' fw_perc        - обводненность, %
        ' pb_atma        - давление насыщения, абс. атм
        ' результат      - значение забойного давления, абс. атм
        'description_end

        Try
            Dim res As New UnfClassLibrary.CReservoirVogel
            If pb_atma <= 0 Then pb_atma = 0   ' поставим ноль иначе флюид подсчитает по корреляции значение
            res.InitProp(pres_atma, pb_atma, fw_perc)
            res.pi_sm3dayatm = pi_sm3dayatm
            IPR_pwf_atma = res.calc_pwf_atma(qliq_sm3day)
            res = Nothing

            Exit Function
        Catch ex As Exception
            IPR_pwf_atma = -1
            Dim msg As String
            msg = "Error:IPR_pwf_atma:" & Err.Description

            Throw New ApplicationException(msg)
        End Try

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет коэффициента продуктивности пласта
    ' по данным тестовой эксплуатации
    Public Function IPR_pi_sm3dayatm(
                 ByVal Qtest_sm3day As Double,
                 ByVal pwf_test_atma As Double,
                 ByVal pres_atma As Double,
        Optional ByVal fw_perc As Double = 0,
        Optional ByVal pb_atma As Double = -1)
        ' Qtest_sm3day   - тестовый дебит скважины, ст.м3/сут
        ' pwf_test_atma  - тестовое забойное давление, абс. атм
        ' Pres_atma      - пластовое давление, абс. атм
        ' fw_perc        - обводненность, %
        ' pb_atma        - давление насыщения, абс. атм
        ' результат      - значение коэффициента продуктивности, ст.м3/сут/атм
        'description_end

        Try
            Dim res As New UnfClassLibrary.CReservoirVogel
            If pb_atma <= 0 Then pb_atma = 0   ' поставим ноль иначе флюид подсчитает по корреляции значение
            res.InitProp(pres_atma, pb_atma, fw_perc)
            IPR_pi_sm3dayatm = res.calc_pi_sm3dayatm(Qtest_sm3day, pwf_test_atma)
            res = Nothing

            Exit Function
        Catch ex As Exception
            IPR_pi_sm3dayatm = -1
            Dim msg As String
            msg = "error in function :IPR_pi_sm3dayatm:" & Err.Description

            Throw New ApplicationException(msg)
        End Try

    End Function

End Module

'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
'функции для расчета гидродинамических исследований
'
Option Explicit On
Imports System.Math

Public Module u7_Excel_functions_transient

    Function Stehfest(ByVal func_name As String,
                  ByVal td As Double,
                  ByVal CoeffA() As Double) As Double
        Dim SumR As Double, DlogTW As Double, z As Double
        Dim j As Integer, N As Integer
        Dim v() As Integer ', M As Integer
        Dim plapl As Double



        SumR = 0#
        N = 12
        DlogTW = Log(2.0#)
        v = coef_stehfest(N)

        For j = 1 To N
            z = j * DlogTW / td
            ' через if Добавить вызовы
            'plapl = Application.Run(func_name, z, CoeffA)
            SumR = SumR + v(j) * plapl * z / j
        Next j
        Stehfest = SumR

    End Function
    Private Function coef_stehfest(N As Integer)
        Dim g(20) As Double, h(10) As Double
        Dim NH As Integer, SN As Double
        Dim K As Integer, k1 As Integer, K2 As Integer
        Dim i As Integer, FI As Double
        Dim v(20) As Double
        Dim M As Integer

        If M <> N Then
            M = N
            g(1) = 1.0#
            NH = N / 2
            For i = 2 To N
                g(i) = g(i - 1) * i
            Next i
            h(1) = 2.0# / g(NH - 1)
            For i = 2 To NH
                FI = i
                If i <> NH Then
                    h(i) = (FI ^ NH) * g(2 * i) / (g(NH - i) * g(i) * g(i - 1))
                Else
                    h(i) = (FI ^ NH) * g(2 * i) / (g(i) * g(i - 1))
                End If
            Next i
            SN = 2 * (NH - (NH \ 2) * 2) - 1
            For i = 1 To N
                v(i) = 0#
                k1 = (i + 1) \ 2
                K2 = i
                If K2 > NH Then K2 = NH
                For K = k1 To K2
                    If 2 * K - i = 0 Then
                        v(i) = v(i) + h(K) / (g(i - K))
                        GoTo nxtIt
                    End If
                    If i = K Then
                        v(i) = v(i) + h(K) / g(2 * K - i)
                        GoTo nxtIt
                    End If
                    v(i) = v(i) + h(K) / (g(i - K) * g(2 * K - i))
nxtIt:          Next K
                v(i) = SN * v(i)
                SN = -SN
            Next i
        End If
        coef_stehfest = v

    End Function


    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' Расчет интегральной показательной функции Ei(x)
    Function Ei(ByVal X As Double)
        ' x  - агрумент функции, может быть и положительным и отрицательным
        ' результат - значение функции
        'description_end
        If X > 0 Then
            'Ei = UnfClassLibrary.exponentialintegralei(X)
        Else
            'Ei = -exponentialintegralei(-X, 1)
        End If
    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' Расчет интегральной показательной функции $E_1(x)$
    ' для вещественных положительных x, x>0 верно E_1(x)=- Ei(-x)
    Function E_1(ByVal X As Double)
        ' x  - агрумент функции, может быть и положительным и отрицательным
        ' результат - значение функции
        'description_end

        'E_1 = ExponentialIntegralEN(X, 1)
    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' Расчет неустановившегося решения уравнения фильтрации
    ' для различных моделей радиального притока к вертикльной скважине
    ' основано не решениях в пространстве Лапласа и преобразовании Стефеста
    Function transient_pd_radial(ByVal td As Double,
                    Optional ByVal cd As Double = 0,
                    Optional ByVal skin As Double = 0,
                    Optional ByVal rd As Double = 1,
                    Optional model As Integer = 0)
        ' td         - безразмерное время для которого проводится расчет
        ' сd         - безразмерный коэффициент влияния ствола скважины
        ' skin       - скин-фактор, безразмерный skin>0.
        '              для skin<0 используйте эффективный радиус скважины
        ' rd         - безразмерное расстояние для которого проводится расчет
        '              rd=1 соответвует забою скважины
        ' model      - модель проведения расчета. 0 - модель линейного стока Ei
        '              1 - модель линейного стока через преобразование Стефеста
        '              2 - конечный радиус скважины
        '              3 - линейный сток со скином и послепритоком
        '              4 - конечный радиус скважины со скином и послепритоком
        ' результат - безразмерное давление pd
        'description_end


        Try
            Dim CoeffA(4) As Double
            If rd < 1 Then rd = 1
            If skin < 0 Then skin = 0
            CoeffA(0) = rd
            CoeffA(1) = cd
            CoeffA(2) = skin
            CoeffA(3) = model
            Select Case model

                Case 0
                    transient_pd_radial = 0.5 * E_1(rd ^ 2 / 4 / td)
                Case 1
                    transient_pd_radial = Abs(Stehfest("pd_lalp_Ei", td, CoeffA))
                Case 2
                    transient_pd_radial = Abs(Stehfest("pd_lalp_rw", td, CoeffA))
                Case 3
                    CoeffA(3) = 0
                    transient_pd_radial = Abs(Stehfest("pd_lalp_cd_skin", td, CoeffA))
                Case 4
                    CoeffA(3) = 1
                    transient_pd_radial = Abs(Stehfest("pd_lalp_cd_skin", td, CoeffA))
                Case 5
                    transient_pd_radial = Abs(Stehfest("pd_lalp_wbs", td, CoeffA))
            End Select
            ' здесь abs чтобы при маленьких значениях pd около нуля все оставалось положительным
            Exit Function
        Catch ex As Exception
            transient_pd_radial = -1
            Dim msg As String
            msg = "Error:transient_pd_radial:" & Err.Description

            Throw New ApplicationException(msg)
        End Try

    End Function


    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет изменения забойного давления после запуска скважины
    ' с постоянным дебитом (terminal rate solution)
    Function transient_pwf_radial_atma(ByVal t_hr As Double,
                                   ByVal qliq_sm3day As Double,
                          Optional ByVal pi_atma As Double = 250,
                          Optional ByVal skin As Double = 0,
                          Optional ByVal cs_1atm As Double = 0,
                          Optional ByVal r_m As Double = 0.1,
                          Optional ByVal rw_m As Double = 0.1,
                          Optional ByVal k_mD As Double = 100,
                          Optional ByVal h_m As Double = 10,
                          Optional ByVal porosity As Double = 0.2,
                          Optional ByVal mu_cP As Double = 1,
                          Optional ByVal b_m3m3 As Double = 1.2,
                          Optional ByVal ct_1atm As Double = 0.00001,
                          Optional ByVal model As Integer = 0) As Double
        ' t_hr        - время для которого проводится расчет, час
        ' qliq_sm3day - дебит запуска скважины, м3/сут в стандартных условиях
        ' pi_atma     - начальное пластовое давление, атма
        ' skin        - скин - фактор, может быть отрицательным
        ' cs_1atm     - коэффициент влияния ствола скважины, 1/атм
        ' r_m         - расстояние от скважины для которого проводится расчет, м
        ' rw_m        - радиус скважины, м
        ' k_mD        - проницаемость пласта, мД
        ' h_m         - толщина пласта, м
        ' porosity    - пористость
        ' mu_cP       - вязкость флюида в пласте, сП
        ' b_m3m3      - объемный коэффициент нефти, м3/м3
        ' ct_1atm     - общая сжимаемость системы в пласте, 1/атм
        ' model      - модель проведения расчета. 0 - модель линейного стока Ei
        '              1 - модель линейного стока через преобразование Стефеста
        '              2 - конечный радиус скважины
        '              3 - линейный сток со скином и послепритоком
        '              4 - конечный радиус скважины со скином и послепритоком
        ' результат -  давление pwf
        'description_end

        Dim td As Double, cd As Double, rd As Double
        Dim pd As Double
        Dim delta_p_atm As Double

        cd = 0.159 / h_m / porosity / ct_1atm / (rw_m * rw_m) * cs_1atm
        If skin < 0 Then
            rw_m = rw_m * Exp(-skin)
            If r_m < rw_m Then r_m = rw_m
            skin = 0
        End If

        td = 0.00036 * k_mD / porosity / mu_cP / ct_1atm / (rw_m * rw_m) * t_hr
        rd = r_m / rw_m

        pd = transient_pd_radial(td, cd, skin, rd, model)

        delta_p_atm = 18.41 * qliq_sm3day * b_m3m3 * mu_cP / k_mD / h_m * pd
        transient_pwf_radial_atma = pi_atma - delta_p_atm

    End Function


    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет безразмерного коэффициента влияния ствола скважины (определение)
    Function transient_def_cd(ByVal cs_1atm As Double,
             Optional ByVal rw_m As Double = 0.1,
             Optional ByVal h_m As Double = 10,
             Optional ByVal porosity As Double = 0.2,
             Optional ByVal ct_1atm As Double = 0.00001
             ) As Double
        ' cs_1atm     - коэффициент влияния ствола скважины, 1/атм
        ' rw_m        - радиус скважины, м
        ' h_m         - толщина пласта, м
        ' porosity    - пористость
        ' ct_1atm     - общая сжимаемость системы в пласте, 1/атм
        ' результат   - безразмерный коэффициент влияния ствола скважины  cd
        'description_end

        transient_def_cd = 0.159 / h_m / porosity / ct_1atm / (rw_m * rw_m) * cs_1atm

    End Function


    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет коэффициента влияния ствола скважины (определение)
    Function transient_def_cs_1atm(ByVal cd As Double,
             Optional ByVal rw_m As Double = 0.1,
             Optional ByVal h_m As Double = 10,
             Optional ByVal porosity As Double = 0.2,
             Optional ByVal ct_1atm As Double = 0.00001
             ) As Double
        ' cs_1atm     - коэффициент влияния ствола скважины, 1/атм
        ' rw_m        - радиус скважины, м
        ' h_m         - толщина пласта, м
        ' porosity    - пористость
        ' ct_1atm     - общая сжимаемость системы в пласте, 1/атм
        ' результат   - коэффициент влияния ствола скважины  cs
        'description_end

        transient_def_cs_1atm = 1 / 0.159 * h_m * porosity * ct_1atm * (rw_m * rw_m) * cd

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет безразмерного времени (определение)
    Function transient_def_td(ByVal t_day As Double,
             Optional ByVal rw_m As Double = 0.1,
             Optional ByVal k_mD As Double = 100,
             Optional ByVal porosity As Double = 0.2,
             Optional ByVal mu_cP As Double = 1,
             Optional ByVal ct_1atm As Double = 0.00001
             ) As Double
        ' t_day       - время для которого проводится расчет, сут
        ' rw_m        - радиус скважины, м
        ' k_mD        - проницаемость пласта, мД
        ' porosity    - пористость
        ' mu_cP       - вязкость флюида в пласте, сП
        ' ct_1atm     - общая сжимаемость системы в пласте, 1/атм
        ' результат   - безразмерное время td
        'description_end

        transient_def_td = 0.00036 * k_mD / porosity / mu_cP / ct_1atm / (rw_m * rw_m) * t_day

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет времени по безразмерному времени (определение)
    Function transient_def_t_day(ByVal td As Double,
             Optional ByVal rw_m As Double = 0.1,
             Optional ByVal k_mD As Double = 100,
             Optional ByVal porosity As Double = 0.2,
             Optional ByVal mu_cP As Double = 1,
             Optional ByVal ct_1atm As Double = 0.00001
             ) As Double
        ' t_day       - время для которого проводится расчет, сут
        ' rw_m        - радиус скважины, м
        ' k_mD        - проницаемость пласта, мД
        ' porosity    - пористость
        ' mu_cP       - вязкость флюида в пласте, сП
        ' ct_1atm     - общая сжимаемость системы в пласте, 1/атм
        ' результат   - время t
        'description_end

        transient_def_t_day = 1 / 0.00036 / k_mD * porosity * mu_cP * ct_1atm * (rw_m * rw_m) * td

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет безразмерного давления (определение)
    Function transient_def_pd(ByVal pwf_atma As Double,
                          ByVal qliq_sm3day As Double,
                 Optional ByVal pi_atma As Double = 250,
                 Optional ByVal k_mD As Double = 100,
                 Optional ByVal h_m As Double = 10,
                 Optional ByVal mu_cP As Double = 1,
                 Optional ByVal b_m3m3 As Double = 1.2
             ) As Double
        ' pwf_atma    - забойное давление, атма
        ' qliq_sm3day - дебит запуска скважины, м3/сут в стандартных условиях
        ' pi_atma     - начальное пластовое давление, атма
        ' k_mD        - проницаемость пласта, мД
        ' h_m         - толщина пласта, м
        ' mu_cP       - вязкость флюида в пласте, сП
        ' b_m3m3      - объемный коэффициент нефти, м3/м3
        ' результат   - безразмерное время td
        'description_end

        transient_def_pd = k_mD * h_m / 18.41 / qliq_sm3day / mu_cP / b_m3m3 * (pi_atma - pwf_atma)

    End Function

    'description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' расчет безразмерного давления (определение)
    Function transient_def_pwf_atma(ByVal pd As Double,
                          ByVal qliq_sm3day As Double,
                 Optional ByVal pi_atma As Double = 250,
                 Optional ByVal k_mD As Double = 100,
                 Optional ByVal h_m As Double = 10,
                 Optional ByVal mu_cP As Double = 1,
                 Optional ByVal b_m3m3 As Double = 1.2
             ) As Double
        ' pwf_atma    - забойное давление, атма
        ' qliq_sm3day - дебит запуска скважины, м3/сут в стандартных условиях
        ' pi_atma     - начальное пластовое давление, атма
        ' k_mD        - проницаемость пласта, мД
        ' h_m         - толщина пласта, м
        ' mu_cP       - вязкость флюида в пласте, сП
        ' b_m3m3      - объемный коэффициент нефти, м3/м3
        ' результат   - безразмерное время td
        'description_end

        transient_def_pwf_atma = pi_atma - 18.41 / k_mD / h_m * qliq_sm3day * mu_cP * b_m3m3 * pd

    End Function

End Module

Public Module u7_solve
    Public Function solve_equation_bisection(func_name As String,
                                           ByVal x1 As Double,
                                           ByVal x2 As Double,
                                           CoeffA As Object,
                                           prm As CSolveParam) As Boolean
        ' func_name             - название функции для которой ищем решение
        ' x1                    - левая граница аргумента для поиска решения
        ' x2                    - правая граница аргумента для поиска решения
        ' coeffA                - параметры функции для которой ищем решение
        ' prm                   - объект с настройками поиска решения
        '                         через этот же объект возвращаются решение и его параметры

        Dim y1 As Double
        Dim y2 As Double
        Dim y_temp As Double
        Dim x_temp As Double
        Dim i As Long
        i = 0
        Try
            ' определим значения параметров на границе
            If func_name = "calc_dq_gas_pd_valve" Then
                y1 = calc_dq_gas_pd_valve(x1, CoeffA)
                y2 = calc_dq_gas_pd_valve(x2, CoeffA)
            ElseIf func_name = "calc_choke_dp_error_calibr_grav_atm" Then
                y1 = calc_choke_dp_error_calibr_grav_atm(x1, CoeffA)
                y2 = calc_choke_dp_error_calibr_grav_atm(x2, CoeffA)
            Else
                y1 = calc_dq_gas_pu_valve(x1, CoeffA)
                y2 = calc_dq_gas_pu_valve(x2, CoeffA)
            End If
            With prm
                If y1 * y2 > 0 Then
                    ' если значения на границе одного знака - то метод поиска решения не работает
                    ' возможно решения нет и найти его не получится
                    .iterations = 0
                    .found_solution = False
                    .msg = "solve_equation_bisection: values at segment's ends must have a different sign"
                    solve_equation_bisection = False
                    Exit Function
                End If
                ' начинаем цикл поиска решений (итерации)
                Do
                    i += 1
                    ' делим отрезок пополам
                    x_temp = (x1 + x2) / 2
                    If func_name = "calc_dq_gas_pd_valve" Then
                        y_temp = calc_dq_gas_pd_valve(x_temp, CoeffA)
                    Else
                        y_temp = calc_dq_gas_pu_valve(x_temp, CoeffA)
                    End If
                    If Math.Abs(y_temp) < .y_tolerance Then
                        solve_equation_bisection = True
                        .x_solution = x_temp
                        .y_solution = y_temp
                        .iterations = i
                        .found_solution = True
                        .msg = "solve_equation_bisection: done by  " + CStr(i) + " iterations, tolerance " + CStr(.y_tolerance)
                        Exit Function
                    Else
                        If y_temp * y1 > 0 Then
                            x1 = x_temp
                            y1 = y_temp
                        Else
                            x2 = x_temp
                            y2 = y_temp
                        End If
                    End If
                Loop Until i >= 100

                solve_equation_bisection = False
                .x_solution = x_temp
                .y_solution = y_temp
                .iterations = i
                .found_solution = False
                .msg = "solve_equation_bisection: too many iterations " + CStr(i)
            End With
            Exit Function

        Catch ex As Exception
            With prm
                .iterations = i
                .found_solution = False
                .msg = "solve_equation_bisection error "
            End With
        End Try

    End Function

    ' функция для поиска решения по расчету давления в клапане
    Public Function calc_dq_gas_pu_valve(Pu As Double, CoeffA() As Object) As Double
        Dim q_gas As Double, d_mm As Double, pd As Double, gg As Double, t As Double
        Dim c_calibr As Double
        q_gas = CDbl(CoeffA(0))
        d_mm = CDbl(CoeffA(1))
        pd = CDbl(CoeffA(2))
        gg = CDbl(CoeffA(3))
        t = CDbl(CoeffA(4))
        c_calibr = CDbl(CoeffA(5))

        calc_dq_gas_pu_valve = q_gas - CDbl(GLV_q_gas_sm3day(d_mm, Pu, pd, gg, t, c_calibr)(0))

    End Function

    Public Function calc_dq_gas_pd_valve(pd As Double, CoeffA() As Object) As Double
        Dim q_gas As Double, d_mm As Double, Pu As Double, gg As Double, t As Double
        Dim c_calibr As Double
        q_gas = CDbl(CoeffA(0))
        d_mm = CDbl(CoeffA(1))
        Pu = CDbl(CoeffA(2))
        gg = CDbl(CoeffA(3))
        t = CDbl(CoeffA(4))
        c_calibr = CDbl(CoeffA(5))

        calc_dq_gas_pd_valve = q_gas - CDbl(GLV_q_gas_sm3day(d_mm, Pu, pd, gg, t, c_calibr)(0))

    End Function

    ' функция расчета ошибки в оценке давления для штуцера
    ' в зависимости от поправки на калибровочный параметр
    Public Function calc_choke_dp_error_calibr_grav_atm(ByVal c_calibr As Double,
                                                        CoeffA() As Object) As Double
        Dim pt As PTtype
        Dim pt0 As PTtype
        Dim choke As CChoke
        Dim p_in_atma As Double
        Dim p_out_atma As Double

        ' read coeffA - parameters
        choke = CoeffA(0)
        p_in_atma = CDbl(CoeffA(1))
        p_out_atma = CDbl(CoeffA(2))

        With choke
            .c_calibr_fr = c_calibr
            pt0.p_atma = p_out_atma
            pt0.t_C = .t_choke_C
            pt = .calc_choke_p(pt0, calc_p_down:=0)
            calc_choke_dp_error_calibr_grav_atm = (pt.p_atma - p_in_atma)
        End With

    End Function
End Module

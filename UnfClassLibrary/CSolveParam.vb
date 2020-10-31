' Параметры и результаты решения уравнений
' вида y = f(x) = 0
'
Option Explicit On

Public Class CSolveParam
    Public max_iterations As Integer ' допустимое максимальное количество итераций
    Public y_tolerance As Double ' допустимая погрешность решения
    Public x_tolerance As Double ' допустимая погрешность аргумента
    Public iterations As Long  ' количество итераций
    Public found_solution As Boolean
    Public msg As String
    Public x_solution As Double
    Public y_solution As Double
    Public obj As Object   ' ссылка на объект расчета
    Public Sub Class_Initialize(Optional ByVal max_iterations_ As Double = 100,
                               Optional ByVal msg_ As String = "",
                                Optional ByVal y_tolerance_ As Double = 0.001,
                                Optional ByVal x_tolerance_ As Double = 0.001,
                                Optional ByVal iterations_ As Double = 0,
                                Optional ByVal x_solution_ As Double = 0,
                                Optional ByVal y_solution_ As Double = 0,
                                Optional ByVal found_solution_ As Boolean = False)

        max_iterations = max_iterations_
        msg = msg_
        y_tolerance = y_tolerance_
        x_tolerance = x_tolerance_
        iterations = iterations_
        x_solution = x_solution_
        y_solution = y_solution_
        found_solution = found_solution_
    End Sub

    ' функция ищет корни уравнения вида
    ' f(x) = 0 на отрезке [x1..x2]
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
        Dim err_msg As String
        i = 0
        On Error GoTo err1
        ' определим значения параметров на границе
        y1 = Application.Run(func_name, x1, CoeffA)
        y2 = Application.Run(func_name, x2, CoeffA)
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
                i = i + 1
                ' делим отрезок пополам
                x_temp = (x1 + x2) / 2
                y_temp = Application.Run(func_name, x_temp, CoeffA)
                If Abs(y_temp) < .y_tolerance Then
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
err1:
        On Error GoTo 0
        With prm
            .iterations = i
            .found_solution = False
            .msg = "solve_equation_bisection error " & Err.Description
        End With
        solve_equation_bisection = False
    End Function
End Class

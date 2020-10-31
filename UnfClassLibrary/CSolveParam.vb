Option Explicit On

Public Class CSolveParam
    Public max_iterations As Integer ' допустимое максимальное количество итераций
    Public y_tolerance As Double ' допустимая погрешность решения
    Public x_tolerance As Double ' допустимая погрешность аргумента
    Public iterations As Double   ' количество итераций
    Public found_solution As Boolean
    Public msg As String
    Public x_solution As Double
    Public y_solution As Double
    Public obj As Object   ' ссылка на объект расчета
    Public Sub Class_Initialize(Optional ByVal max_iterations_ As Integer = 100,
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
End Class
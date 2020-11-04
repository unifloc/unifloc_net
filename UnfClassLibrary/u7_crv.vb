'=======================================================================================
'Unifloc 7.25  coronav                                     khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
' функции для работы с кривыми из интерфейса Excel

Module u7_crv
    ' description_to_manual      - для автогенерации описания - помещает заголовок функции и окружающие комментарии в мануал (со след строки)
    ' функция поиска значения функции по заданным табличным данным (интерполяция)
    Public Function crv_interpolation(x_points() As Double, y_points() As Double, x_val As Double,
                                Optional ByVal type_interpolation As Integer = 0) As Double
        ' x_points  - таблица аргументов функции
        ' y_points  - таблица значений функции
        '             количество агрументов и значений функции должно совпадать
        '             для табличной функции одному аргументу соответствует
        '             строго одно значение функции (последнее)
        ' x_val     - аргумент для которого надо найти значение
        '             одно значение в ячейке или диапазон значений
        '             для диапазона аргументов будет найден диапазон значений
        '             диапазоны могут быть заданы как в строках,
        '             так и в столбцах
        ' type_interpolation - тип интерполяции
        '             0 - линейная интерполяция
        '             1 - кубическая интерполяция
        '             2 - интерполяция Акима (выбросы)
        '                 https://en.wikipedia.org/wiki/Akima_spline
        '             3 - кубический сплай Катмулла Рома
        '                 https://en.wikipedia.org/wiki/Cubic_Hermite_spline
        ' результат
        '             значение функции для заданного x_val
        'description_end

        'Dim x_arr(), y_arr(), x_val_arr(), y_out(,) As Double
        Dim y_val_temp As Double
        'Dim X_Range(,) As Double
        'Dim Y_Range(,) As Double
        Dim i As Integer
        Dim crv As CInterpolation
        crv = New CInterpolation
        Dim interp_type As String
        Dim y_out(,) As Double
        ReDim y_out(0 To 1, 1)

        Try
            ' прочитаем все исходные вектора с листа и подготовим выходной массив
            'Call read_xy_vectors(x_points, y_points, x_val, x_arr, y_arr, x_val_arr, y_out)
            ' формируем объект функции для работы с ним
            For i = 0 To x_points.GetUpperBound(0)
                crv.AddPoint(x_points(i), y_points(i))
            Next i

            ' готовим интерполяцию
            Select Case type_interpolation
                Case 0
                    interp_type = "Linear"
                Case 1
                    interp_type = "Cubic"
                Case 2
                    interp_type = "Akima"
                Case 3
                    interp_type = "CatmullRom"
            End Select

            crv.Init_interpolation(interp_type)

            ' интерполируем требуемые данные и готовим для вывода массива значений
            'For i = x_val.GetLowerBound(0) To x_val.GetUpperBound(0)
            y_val_temp = crv.Get_interpolation_point(x_val)
            'If y_out.GetUpperBound(1) > 1 Then                                         ' (0)?!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            '    y_out(i, y_out.GetLowerBound(0)) = y_val_temp
            'Else
            '    y_out(y_out.GetLowerBound(0), i) = y_val_temp
            'End If
            'Next i
            crv_interpolation = y_val_temp  'y_out 

            Exit Function
        Catch ex As Exception
            Dim msg As String
            msg = "Error:crv_interpolation:"

            Throw New ApplicationException(msg)
        End Try

    End Function

    '' рабочая функция для чтения данных кривых из range
    'Private Sub read_xy_vectors(x_points As Double, y_points As Double, x_val As Double,
    '                                ByRef x_arr() As Double,
    '                                ByRef y_arr() As Double,
    '                                ByRef x_val_arr() As Double,
    '                                ByRef y_val_arr(,) As Double)

    '    'Dim X_Range(,) As Double
    '    'Dim Y_Range(,) As Double
    '    'Dim x_val_range(,) As Double

    '    'Dim check_x As Boolean
    '    'Dim check_y As Boolean
    '    'Dim i As Integer

    '    Try
    '        Call convert_to_array(x_points, x_arr)
    '        Call convert_to_array(y_points, y_arr)
    '        Call convert_to_array(x_val, x_val_arr)
    '        ReDim y_val_arr(0 To x_val_arr.GetUpperBound(0), 1)
    '        Exit Sub
    '    Catch ex As Exception
    '        Dim msg As String
    '        msg = "Error:read_xy_vectors:"

    '        Throw New ApplicationException(msg)
    '    End Try

    'End Sub


End Module

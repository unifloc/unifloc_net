'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
'
Option Explicit On
Imports System.Math

' Класс описывающий траекторию трубы. Содержит методы для работы с инклинометрией и траекторией
' На вход подаются данные по траектории или инклинометрии в исходном виде (из базы данных)
' На выходе объект пригодный для проведения расчетов с использованием класса скважины и трубы
' У трубы могут меняться наклон по участкам и диаметр - шероховатость и толщина стенки постоянны для
' всего сегмента.
'
' История
' 2016.01.18    Хабибуллин Ринат
' 2019.10.25    рефакторинг для упрощения и оптимизации логики (убраны лишние сущности из скважины h_perf)


Public Class CPipeTrajectory

    ' тип описывающий полную конструкцию скважины в заданной точке
    Private Structure WELL_POINT_FULL
        Public h_mes_m As Double           ' измеренная глубина
        Public ang_deg As Double           ' угол  от вертикали
        Public h_abs_m As Double           ' абсолютная глубина
        Public diam_in_m As Double         ' диаметр трубы, внутренний
        Public diam_out_m As Double        ' диаметр трубы, внешний
        Public roughness_m As Double       ' шероховатость
    End Structure

    Private h_abs_init_m_ As New CInterpolation      ' исходный массив абсолютных глубин
    Private angle_init_deg_ As New CInterpolation    ' исходный массив углов
    Private diam_init_m_ As New CInterpolation       ' исходный массив значений диаметров НКТ
    Private wall_thickness_mm_ As Double                    ' пока считаем что НКТ всегда имеет одинаковую толщину - потом можно будет учесть
    Private wall_roughness_m_ As Double                    ' пока также считаем, что шероховатость везде тоже одинакова
    Private pipe_trajectory_() As WELL_POINT_FULL     ' полная конструкция скважины пригодная для расчетов - итоговое свойство класса
    Private num_points_out_ As Integer               ' количество точек в выходном массиве
    Private length_between_points_m_ As Double       ' мин растояние между точками при генерации исходного массива
    Private construction_points_curve_ As New CInterpolation  ' точки которые должны быть добавлены (измеренная глубина - абсолютная глубина)
    Private h_points_curve_ As New CInterpolation    ' результирующие точки для заполнения массивов

    Private Sub Class_Initialize(Optional ByVal wall_thickness_mm As Double = 10,
                                 Optional ByVal wall_roughness_m As Double = 0.0001,
                                 Optional ByVal length_between_points_m As Double = 100,
                                 Optional ByVal h_abs_init_m As Boolean = False,
                                 Optional ByVal angle_init_deg As Boolean = True,
                                 Optional ByVal diam_init_m As Boolean = True,
                                 Optional ByVal num_points_out As Integer = 0)

        ' установка значений по умолчанию
        wall_thickness_mm_ = wall_thickness_mm
        wall_roughness_m_ = wall_roughness_m
        length_between_points_m_ = length_between_points_m    ' по умолчанию ставим расстояние между точками инклинометрии 100 м
        h_abs_init_m_.isStepFunction = h_abs_init_m   ' абсолютные глубины линейно интерполируются
        angle_init_deg_.isStepFunction = angle_init_deg   ' углы - ступенчатая функция
        diam_init_m_.isStepFunction = diam_init_m   ' диаметры - ступенчатая функция
        num_points_out_ = num_points_out
    End Sub

    Public Sub init_from_curves(ByVal habs_curve_m As CInterpolation,
                            ByVal diam_curve_mm As CInterpolation)
        Dim i As Integer
        Dim ang
        Dim sina As Double, cosa As Double

        angle_init_deg_.ClearPoints()

        h_abs_init_m_ = habs_curve_m
        diam_curve_mm.isStepFunction = True
        diam_init_m_ = diam_curve_mm.Transform(multY:=const_convert_mm_m)
        diam_init_m_.isStepFunction = True   ' диаметры - ступенчатая функция

        For i = 2 To habs_curve_m.Num_points
            sina = (habs_curve_m.PointY(i) - habs_curve_m.PointY(i - 1)) / (habs_curve_m.PointX(i) - habs_curve_m.PointX(i - 1))
            cosa = Sqrt(MaxReal(1 - sina ^ 2, 0))
            If cosa = 0 Then
                ang = 90 * sina
            Else
                ang = Atan(sina / cosa) * 180 / const_Pi
            End If
            angle_init_deg_.AddPoint(habs_curve_m.PointX(i - 1), ang)
        Next i
        calc_trajectory()
    End Sub

    Private Function calc_trajectory() As Boolean
        ' функция расчета траектории скважины - из исходных данных считает нормализованные выходные данные и готовит данные для скважины
        Dim i As Integer
        Dim h As Double
        Dim allDone As Boolean
        Dim i_constrPoint As Integer
        Dim Hmes As Double, HmesNext As Double

        construction_points_curve_.ClearPoints()
        construction_points_curve_.AddPoint(h_abs_init_m_.Minx, h_abs_init_m_.GetPoint(h_abs_init_m_.Minx))   ' на всякий случай добавим в конструкцию нулевую точку из которой стартуем

        For i = 1 To diam_init_m_.Num_points
            h = diam_init_m_.PointX(i)
            If h < h_abs_init_m_.Minx Then h = h_abs_init_m_.Minx
            construction_points_curve_.AddPoint(h, h_abs_init_m_.GetPoint(h))
        Next i

        Hmes = h_abs_init_m_.Minx
        i = 0
        i_constrPoint = 1
        allDone = False

        ' начинаем цикл, в котором формируем набор точек из которых должна состоять траетория скважины
        ' учитывая минимальное расстояние между точками и присутствие в списке всех обязательных точек
        Do
            HmesNext = Hmes + length_between_points_m_    ' смотрим куда должна попасть след точка
            If construction_points_curve_.PointX(i_constrPoint) < HmesNext Then
                Hmes = construction_points_curve_.PointX(i_constrPoint)
                i_constrPoint = i_constrPoint + 1
            Else
                Hmes = HmesNext
            End If

            If Hmes >= construction_points_curve_.Maxx Then
                Hmes = construction_points_curve_.Maxx
                allDone = True
            End If
            h_points_curve_.AddPoint(Hmes, h_abs_init_m_.GetPoint(Hmes))   ' сохраняем измеренную и абсолютную глубины тут
        Loop Until allDone
        ' набор точек для траектории сформирован

        ' дальше надо по данному набору заполнить все элементы массива конструкции

        ReDim pipe_trajectory_(h_points_curve_.Num_points - 1)
        For i = 0 To h_points_curve_.Num_points - 1
            Hmes = h_points_curve_.PointX(i + 1)
            With pipe_trajectory_(i)
                .h_mes_m = Hmes
                .h_abs_m = h_points_curve_.PointY(i + 1)
                .ang_deg = angle_init_deg_.GetPoint(Hmes)
                .diam_in_m = diam_init_m_.GetPoint(Hmes)
                .diam_out_m = diam_init_m_.GetPoint(Hmes) + wall_thickness_mm_ * const_convert_mm_m
                .roughness_m = wall_roughness_m_
            End With
        Next i

        ' траекторию сформировали
        ' надо теперь подготовить массивы для класса скважина

    End Function


    ' функция для подготовки траектории с данных листа
    ' должна использоваться для чтения исходных данных с листа
    ' ожидаем, что на входе либо range - тогда их конвертируем в массивы
    ' либо массивы, либо числа - тогда генерируем простые массивы
    Public Sub init_from_vert_range(ByRef h_data_m(,) As Double,
                                    ByRef diam_data_mm(,) As Double,
                                    Optional ByVal h_limit_top_m As Double = 1.0E+20,
                                    Optional ByVal h_limit_bottom_m As Double = -1.0E+20)
        ' h_data_m - инклинометрия - range или двухмерный массив или число
        '            зависимость значений вертикальной глубины от измеренной,
        '            первый столбец - измеренная глубина, м
        '            второй столбец - вертикальная глубина, м
        '            если передано одно число - то будет задана вертикальная траектория заданной глубины
        ' diam_data_mm - значения диаметров от измеренной глубины - range или двухмерный массив или число
        '            первый столбец - измеренная глубина, м
        '            второй столбец - диаметр трубы, мм - применяется от текущего значения глубины и до следующего
        '            если передано одно число - то будет задан постоянный диаметр

        Dim i As Integer
        Dim habs_curve_m As New CInterpolation
        Dim diam_curve_mm As New CInterpolation
        Dim diam_val_mm As Double, h_val As Double
        Dim diam_number As Boolean, h_number As Boolean
        'Dim fix_index As Integer

        diam_number = False
        diam_val_mm = -1
        h_number = False
        h_val = -1
        ' проверим

        For i = h_data_m.GetLowerBound(1) To h_data_m.GetUpperBound(1)
            If h_data_m.GetUpperBound(2) = 1 Then
                habs_curve_m.AddPoint(h_data_m(i, 1), h_data_m(i, 1))
            Else
                habs_curve_m.AddPoint(h_data_m(i, 1), h_data_m(i, 2))
            End If
        Next
        If habs_curve_m.Num_points = 1 Then
            habs_curve_m.AddPoint(0, 0)
        End If

        If diam_data_mm.GetUpperBound(2) = 1 Then
            diam_curve_mm.AddPoint(0, diam_data_mm(1, 1))
            diam_curve_mm.AddPoint(diam_data_mm(diam_data_mm.GetUpperBound(1), 1), diam_data_mm(1, 1))
        ElseIf diam_data_mm.GetUpperBound(2) > 1 Then
            For i = diam_data_mm.GetLowerBound(1) To diam_data_mm.GetUpperBound(1)
                If diam_data_mm(i, 1) < diam_data_mm(diam_data_mm.GetUpperBound(1), 1) Then
                    diam_curve_mm.AddPoint(diam_data_mm(i, 1), diam_data_mm(i, 2))
                End If
            Next
            diam_curve_mm.AddPoint(diam_data_mm(diam_data_mm.GetUpperBound(1), 1), diam_data_mm(i - 1, 2))
        End If

        Call habs_curve_m.CutByValueTrajectory(h_limit_top_m, h_limit_bottom_m)
        Call diam_curve_mm.CutByValueTrajectory(h_limit_top_m, h_limit_bottom_m)

        Call init_from_curves(habs_curve_m, diam_curve_mm)
    End Sub


    ' функции для обеспечения клонирования траекторий
    Public Function get_habs_curve_m() As CInterpolation
        Dim i As Integer
        Dim crv As New CInterpolation
        For i = pipe_trajectory_.GetLowerBound(0) To pipe_trajectory_.GetUpperBound(0)
            With pipe_trajectory_(i)
                Call crv.AddPoint(.h_mes_m, .h_abs_m)
            End With
        Next i
        get_habs_curve_m = crv
    End Function

    Public Function get_diam_curve_mm() As CInterpolation
        Dim i As Integer
        Dim crv As New CInterpolation
        For i = pipe_trajectory_.GetLowerBound(0) To pipe_trajectory_.GetUpperBound(0)
            With pipe_trajectory_(i)
                Call crv.AddPoint(.h_mes_m, .diam_in_m * 1000)
            End With
        Next i
        get_diam_curve_mm = crv
    End Function



    ' ========================================================================
    ' свойства и методы задания параметров
    ' ========================================================================
    Public ReadOnly Property num_points() As Integer
        Get
            num_points = h_points_curve_.Num_points

        End Get
    End Property

    Public ReadOnly Property ang_deg(i As Integer) As Double
        Get
            ang_deg = pipe_trajectory_(i).ang_deg
        End Get
    End Property

    Public ReadOnly Property ang_hmes_deg(h_mes_m As Double) As Double
        Get
            ang_hmes_deg = angle_init_deg_.GetPoint(h_mes_m)
        End Get
    End Property

    Public ReadOnly Property h_mes_m(i As Integer) As Double
        Get
            h_mes_m = pipe_trajectory_(i).h_mes_m
        End Get
    End Property

    Public ReadOnly Property h_abs_m(i As Integer) As Double
        Get
            h_abs_m = pipe_trajectory_(i).h_abs_m
        End Get
    End Property

    Public ReadOnly Property h_abs_hmes_m(ByVal h_mes_m As Double) As Double
        Get
            h_abs_hmes_m = h_points_curve_.GetPoint(h_mes_m)
        End Get
    End Property

    Public ReadOnly Property diam_in_m(i As Integer) As Double
        Get
            diam_in_m = pipe_trajectory_(i).diam_in_m
        End Get
    End Property

    Public ReadOnly Property diam_hmes_m(h_mes_m As Double) As Double
        Get
            diam_hmes_m = diam_init_m_.GetPoint(h_mes_m)
        End Get
    End Property

    'Public ReadOnly Property roughness_m() As Double
    '    Get
    '        roughness_m = wall_roughness_m_
    '    End Get
    'End Property

    Public Property roughness_m() As Double
        Get
            roughness_m = wall_roughness_m_
        End Get
        Set(val As Double)
            Dim i As Integer
            If val > 0 Then wall_roughness_m_ = val
            For i = 0 To h_points_curve_.Num_points - 1
                pipe_trajectory_(i).roughness_m = wall_roughness_m_
            Next i
        End Set
    End Property

    Public ReadOnly Property wall_thickness_m() As Double
        Get
            wall_thickness_m = wall_roughness_m_
        End Get
    End Property

    Public Property wall_thickness_mm() As Double
        Get
            wall_thickness_mm = wall_thickness_mm_
        End Get
        Set(val As Double)
            Dim i As Integer
            Dim Hmes As Double
            If val > 0 Then wall_thickness_mm_ = val
            For i = 0 To h_points_curve_.Num_points - 1
                Hmes = h_points_curve_.PointX(i + 1)
                With pipe_trajectory_(i)
                    .diam_out_m = diam_init_m_.GetPoint(Hmes) + wall_thickness_mm_ * const_convert_mm_m
                End With
            Next i
        End Set
    End Property

    Public ReadOnly Property top_m() As Double
        Get
            top_m = h_points_curve_.PointX(1)
        End Get
    End Property

    Public ReadOnly Property bottom_m() As Double
        Get
            bottom_m = h_points_curve_.PointX(num_points)
        End Get
    End Property

End Class

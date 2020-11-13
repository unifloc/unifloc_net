'=======================================================================================
'Unifloc 7.24  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
'
' класс для хранения и работы с графиками функций y=f(x) заданных в табличном виде
'
' Ver 1.3
' 2019/07/13
' добавлены функции для интерполяциями сплайнами на основе alglib
'
'
' Ver 1.2
' rnt
' обновление для более полного манипулирования графиками
'


Public Class CInterpolation

    Private Structure TDPoint    ' тип для хранения точек
        Public x As Double
        Public y As Double
        Public stable As Boolean    ' специальный признак точки - показывает должна ли она сохраняться при преобразовании
    End Structure

    ' Поиск решения x при известном y
    Public Enum CInterpolation_SOLUTION
        TS_EXTRPOLATION = 0                 ' осуществлять экстраполирование решение
        TS_NO_EXTRPOLATION = 1              ' без экстрополяции
    End Enum

    Private ReadOnly class_name_ As String              ' имя класса для унифицированной генерации сообщений об ошибках
    Private FPoints() As TDPoint            ' исходный массив точек
    Private FSolutionPoints() As TDPoint    ' массив точек решений (поиск x при известном y)
    Private FStablePoints() As Double       ' массив особых (стабильных) точек, которые сохраняются при трансформации кривой
    Private FkPoint As Integer              ' количество точек в массиве
    Private FkSolPoints As Integer          ' количество точек решений
    Private FkStablePoints As Integer       ' количество стабильных точек
    Private FMinY As Double                 ' минимальное значение функции
    Private FMaxY As Double                 ' максимальное значение функции
    'Public Z As Double                      ' неизвестная переменная - не используется ?
    ' флаг определяющий является ли функция линейно интерпорлированной или ступенчатой
    Public isStepFunction As Boolean
    ' доп параметры для описания графиков
    Public Title As String
    Public xName As String
    Public yName As String
    Public note As String
    Public special As Boolean
    Private spline_interpolant As spline1dinterpolant

    Public Sub New()
        class_name_ = "CInterpolation"
        special = False
        Call ClearPoints()
        isStepFunction = False  ' по умолчанию - линейно интерполированная
    End Sub

    Public Function NumStablePoints() As Integer
        NumStablePoints = FkStablePoints
    End Function

    ' свойство возвращает значение стабильной точки по ее номеру, если такая точка есть
    Public Function StablePoint(i As Integer) As Double
        If i > 0 And i <= FkStablePoints Then
            StablePoint = FStablePoints(i - 1)
        Else
            Throw New ApplicationException("Неверный индекс при считывании стабильных точек кривой CInterpolation")
            'Err.Raise kErrcurvestablePointIndex, , "Неверный индекс при считывании стабильных точек кривой CInterpolation"
        End If
    End Function

    Public Function Num_points() As Integer
        Num_points = FkPoint
    End Function

    Public Function PointStable(i As Integer) As Boolean
        If i > 0 And i <= FkPoint Then
            PointStable = FPoints(i - 1).stable
        Else
            Throw New ApplicationException("Неверный индекс при считывании стабильных точек кривой CInterpolation")
            'Err.Raise kErrCurvePointIndex, , "Неверный индекс при считывании точек Х кривой CInterpolation"
        End If
    End Function

    Public Function PointX(i As Integer) As Double
        If i > 0 And i <= FkPoint Then
            PointX = FPoints(i - 1).x
        Else
            Throw New ApplicationException("Неверный индекс при считывании точек Х кривой CInterpolation")
            'Err.Raise kErrCurvePointIndex, , "Неверный индекс при считывании точек Х кривой CInterpolation"
        End If
    End Function

    Public Function PointY(i As Integer) As Double
        If i > 0 And i <= FkPoint Then
            PointY = FPoints(i - 1).y
        Else
            Throw New ApplicationException("Неверный индекс при считывании точек Y кривой CInterpolation")
            'Err.Raise kErrCurvePointIndex, , "Неверный индекс при считывании точек Y кривой CInterpolation"
        End If
    End Function

    Public Function SolutionPointX(i As Integer) As Double
        If i > 0 And i <= FkSolPoints Then
            SolutionPointX = FSolutionPoints(i - 1).x
        Else
            Throw New ApplicationException("Неверный индекс при считывании точек X решений кривой CInterpolation")
            'Err.Raise kErrCurvePointIndex, , "Неверный индекс при считывании точек X решений кривой CInterpolation"
        End If
    End Function

    Public Function SolutionPointY(i As Integer) As Double
        If i > 0 And i <= FkSolPoints Then
            SolutionPointY = FSolutionPoints(i - 1).y
        Else
            Throw New ApplicationException("Неверный индекс при считывании точек Y решений кривой CInterpolation")
            'Err.Raise kErrCurvePointIndex, , "Неверный индекс при считывании точек Y решений кривой CInterpolation"
        End If
    End Function

    Public Function Miny() As Double
        Miny = FMinY
    End Function

    Public Function Maxy() As Double
        Maxy = FMaxY
    End Function

    Public Function Minx() As Double
        If FkPoint = 0 Then Minx = 0 Else Minx = FPoints(0).x
    End Function

    Public Function Maxx() As Double
        If FkPoint = 0 Then Maxx = 0 Else Maxx = FPoints(FkPoint - 1).x
    End Function

    Private Sub FindMinMaxY()
        'находит минимальное и максимальное значение функции
        Dim i As Integer
        If FkPoint > 1 Then
            FMinY = FPoints(FPoints.GetLowerBound(0)).y
            FMaxY = FPoints(FPoints.GetLowerBound(0)).y
            For i = FPoints.GetLowerBound(0) To FPoints.GetUpperBound(0)
                If FPoints(i).y > FMaxY Then
                    FMaxY = FPoints(i).y
                End If
                If FPoints(i).y < FMinY Then
                    FMinY = FPoints(i).y
                End If
            Next i
        End If
    End Sub

    Private Function GetFirstPointNo(ByVal x As Double) As Integer
        Dim i As Integer
        Dim F As Boolean

        i = 0
        F = True
        While F
            F = False
            If i < FkPoint - 1 Then
                If x > FPoints(i).x Then
                    i += 1
                    F = True
                End If
            End If
        End While
        If i = 0 Then i = 1
        GetFirstPointNo = i - 1
    End Function

    Public Function FindSolutions(Yvalue As Double, Optional ByVal with_extrapolation As CInterpolation_SOLUTION = CInterpolation_SOLUTION.TS_EXTRPOLATION) As Integer
        ' FindSolutions Функция поиска решений X по известному Y. По умолчанию расчет ведется с линейной экстраполяцией на краях
        ' @param Yvalue - значение Y
        ' @param with_extrapolation - производить ли экстраполяцию при расчете
        ' @return Количество найденных точек
        Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double
        Dim x As Double
        Dim i As Integer

        Try
            FkSolPoints = 0  ' assume no soutions
            If FPoints.GetUpperBound(0) = FPoints.GetLowerBound(0) Then  ' если только одна точка то ничего нельзя сделать
                FindSolutions = 0
                Exit Function
            End If

            ReDim Preserve FSolutionPoints(FkSolPoints) ' удаляем хранилище точек пересечений
            For i = FPoints.GetLowerBound(0) To FPoints.GetUpperBound(0)
                If i < FPoints.GetUpperBound(0) Then
                    If (FPoints(i).y <= Yvalue) And (FPoints(i + 1).y >= Yvalue) Or (FPoints(i).y >= Yvalue) And (FPoints(i + 1).y <= Yvalue) Then    ' must be solution here
                        If (FPoints(i).y = Yvalue) And (FPoints(i + 1).y = Yvalue) Then   ' infinite solutions of line segment
                            If FkSolPoints = 0 Then
                                AddPointToSolPoints(FPoints(i).x, FPoints(i).y)
                            ElseIf FSolutionPoints(FkSolPoints - 1).x <> FPoints(i).x Then ' особенности VBA, чтобы при FkSolPoints = 0 не падало
                                ' особенности VBA,
                                AddPointToSolPoints(FPoints(i).x, FPoints(i).y)
                            End If

                        Else ' one solution
                            x1 = FPoints(i).x
                            x2 = FPoints(i + 1).x
                            y1 = FPoints(i).y
                            y2 = FPoints(i + 1).y
                            x = (x2 - x1) / (y2 - y1) * (Yvalue - y1) + x1
                            AddPointToSolPoints(x, Yvalue)
                        End If
                    End If
                Else
                    If FPoints(i).y = Yvalue Then
                        AddPointToSolPoints(FPoints(i).x, FPoints(i).y)
                    End If
                End If
            Next i

            If FkSolPoints = 0 And with_extrapolation = CInterpolation_SOLUTION.TS_EXTRPOLATION Then
                ' проверяем существование y на левом крае
                i = 0
                y1 = FPoints(i).y
                y2 = FPoints(i + 1).y
                If ((Yvalue - y1) * (y1 - y2) > 0) Then
                    x1 = FPoints(i).x
                    x2 = FPoints(i + 1).x
                    x = (x2 - x1) / (y2 - y1) * (Yvalue - y1) + x1
                    AddPointToSolPoints(x, Yvalue)
                End If
                ' проверяем существование y на правом крае
                i = FPoints.GetUpperBound(0)
                y1 = FPoints(i).y
                y2 = FPoints(i - 1).y
                If ((Yvalue - y1) * (y1 - y2) > 0) Then
                    x1 = FPoints(i).x
                    x2 = FPoints(i - 1).x
                    x = (x2 - x1) / (y2 - y1) * (Yvalue - y1) + x1
                    AddPointToSolPoints(x, Yvalue)
                End If
            End If
            FindSolutions = FkSolPoints
        Catch ex As Exception
            ' унифицированная реакция на ошибочный ввод ключевых параметров класса
            Dim msg As String, fname As String
            fname = "FindSolutions"
            msg = class_name_ & "." & fname & ": error finding solution for" & Yvalue & " = " & CStr(Yvalue)
            AddLogMsg(msg)
            Throw New ApplicationException(msg)
            'Err.Raise kErrCInterpolation, class_name_ & "." & fname, msg

        End Try

    End Function

    Public Function FindMinOneSolution(Yvalue As Double, Optional ByVal with_extrapolation As CInterpolation_SOLUTION = CInterpolation_SOLUTION.TS_EXTRPOLATION) As Double
        ' FindSolutions Функция поиска решений X по известному Y. По умолчанию расчет ведется с линейной экстраполяцией на краях
        ' @param Yvalue - значение Y
        ' @param with_extrapolation - производить ли экстраполяцию при расчете
        ' @return Возвращает искомое решение, если решение одно,возвращает минимальное значения для решения, если значений несколько,
        ' Вызывает исключение, если решений нет
        Dim points_solve_size As Integer

        points_solve_size = FindSolutions(Yvalue, with_extrapolation)
        If (points_solve_size = 1) Then
            FindMinOneSolution = Me.SolutionPointX(1)
        ElseIf (points_solve_size > 1) Then
            FindMinOneSolution = Me.SolutionPointX(1) ' тут надо проверить - что возвращается, сделать возвращение минимального
        Else

            Throw New ApplicationException("FindMinOneSolution no solutioin")
            FindMinOneSolution = 0
            'Err.Raise kErrArraySize, , "FindMinOneSolution завершился неудачно, решений нет"
        End If
    End Function

    Public Function FindMaxOneSolution(Yvalue As Double, Optional ByVal with_extrapolation As CInterpolation_SOLUTION = CInterpolation_SOLUTION.TS_EXTRPOLATION) As Double
        ' FindSolutions Функция поиска решений X по известному Y. По умолчанию расчет ведется с линейной экстраполяцией на краях
        ' @param Yvalue - значение Y
        ' @param with_extrapolation - производить ли экстраполяцию при расчете
        ' @return Возвращает искомое решение, если решение одно,возвращает максимальное значения для решения, если значений несколько,
        ' Вызывает исключение, если решений нет
        Dim points_solve_size As Integer

        points_solve_size = FindSolutions(Yvalue, with_extrapolation)
        If (points_solve_size = 1) Then
            FindMaxOneSolution = Me.SolutionPointX(1)
        ElseIf (points_solve_size > 1) Then
            FindMaxOneSolution = Me.SolutionPointX(points_solve_size) ' возвращаем послежнюю точку
        Else
            Throw New ApplicationException("FindMaxOneSolution no solutioin")
            FindMaxOneSolution = 0
            'Err.Raise kErrArraySize, , "FindMaxOneSolution завершился неудачно, решений нет"
        End If
    End Function

    Private Sub AddPointToSolPoints(ByVal x As Double, ByVal y As Double)
        Dim i As Integer
        If FkSolPoints > 0 Then
            For i = 0 To FkSolPoints - 1
                If FSolutionPoints(i).x = x Then
                    ' если точка решения уже есть - перезапишем
                    FSolutionPoints(i).y = y
                    Exit Sub
                End If
            Next i
        End If

        ReDim Preserve FSolutionPoints(FkSolPoints)
        FSolutionPoints(FkSolPoints).x = x
        FSolutionPoints(FkSolPoints).y = y
        FkSolPoints += 1
    End Sub

    Public Function GetPoint(ByVal x As Double) As Double
        Dim n As Integer
        Dim x1, x2, y1, y2 As Double
        ' интерполирует или экстраполирует значения по кривой - линейно
        GetPoint = 0
        If FkPoint < 2 And Not isStepFunction Then
            AddLogMsg("CInterpolation.getPoint   error - trying to find intersection with one point line")
            Exit Function
        End If
        ' если ступенчатая функция - то достаточно только одной точки чтобы получить значение где угодно
        If FkPoint < 1 Then
            AddLogMsg("CInterpolation.getPoint   error - trying to find intersection with line without points")
            Exit Function
        End If

        n = GetFirstPointNo(x)
        x1 = FPoints(n).x
        y1 = FPoints(n).y

        If FkPoint > 1 Then
            x2 = FPoints(n + 1).x
            y2 = FPoints(n + 1).y
        Else
            x2 = x1
            y2 = y1
        End If

        ' делаем проверку - если функция ступенчатая то выдаем не интерполированное значение, а значение в предущей точке
        If isStepFunction Then
            If x >= x2 Then
                GetPoint = y2
            Else
                GetPoint = y1
            End If
        Else
            GetPoint = (y2 - y1) / (x2 - x1) * (x - x1) + y1
        End If
    End Function

    Public Function TestPoint(ByVal x As Double) As Integer
        ' проверяет если точка с заданным аргументом
        '
        Dim i, n As Integer

        n = -1
        For i = 0 To FkPoint - 1
            If FPoints(i).x = x Then
                n = i
                Exit For
            End If
        Next i
        TestPoint = n
    End Function

    Public Sub ClearPoints()
        ReDim FPoints(0)
        ReDim FSolutionPoints(0)
        ReDim FStablePoints(0)
        FkPoint = 0
        FkSolPoints = 0
        FkStablePoints = 0
    End Sub

    Public Sub AddPointsCurve(ParamArray crv() As CInterpolation)
        ' добавляет в кривую все точки из другой кривой
        Dim i As Integer, j As Integer
        Dim crv_local As CInterpolation
        ' If crv <> Nothing Then
        For j = crv.GetLowerBound(0) To crv.GetUpperBound(0)
            crv_local = crv(j)
            For i = 1 To crv_local.Num_points
                Me.AddPoint(crv_local.PointX(i), crv_local.PointY(i), crv_local.PointStable(i))
            Next i
        Next j
        ' End If
    End Sub

    Public Sub AddPoint(ByVal x As Double, ByVal y As Double, Optional isStable As Boolean = False)
        ' добавление точки с сортировкой и обеспечением возрастания аргументов
        Dim i, n As Integer
        Dim CheckMinMaxY As Boolean
        Dim tp As TDPoint
        Dim F As Boolean

        Try
            n = TestPoint(x)
            If n >= 0 Then ' если аргумент уже есть в массиве
                FPoints(n).x = x
                If (FPoints(n).y = FMinY) Or (FPoints(n).y = FMinY) Then
                    CheckMinMaxY = True
                Else
                    CheckMinMaxY = False
                    If y > FMaxY Then FMaxY = y
                    If y < FMinY Then FMinY = y
                End If
                FPoints(n).y = y
                FPoints(n).stable = isStable
                If CheckMinMaxY Then Call FindMinMaxY()
                Exit Sub
            End If

            ReDim Preserve FPoints(FkPoint)

            FPoints(FkPoint).x = x
            FPoints(FkPoint).y = y
            FPoints(FkPoint).stable = isStable

            ' дальше сортируем точки, чтобы получилось все хорошо
            If (y > FMaxY) Or (FkPoint = FPoints.GetLowerBound(0)) Then FMaxY = y
            If (y < FMinY) Or (FkPoint = FPoints.GetLowerBound(0)) Then FMinY = y
            FkPoint += 1
            If FkPoint > 1 Then
                i = FkPoint - 1
                F = True
                While F
                    F = False
                    If i > 0 Then
                        If FPoints(i - 1).x > FPoints(i).x Then
                            tp = FPoints(i)
                            FPoints(i) = FPoints(i - 1)
                            FPoints(i - 1) = tp
                            i -= 1
                            F = True
                        End If
                    End If
                End While
            End If
            ' в конце перечитаем массив специальных стабильных точек
            Call UpdateStablePointsList()
        Catch ex As Exception
            ' унифицированная реакция на ошибочный ввод ключевых параметров класса
            Dim msg As String, fname As String
            fname = "AddPoint"
            msg = class_name_ & "." & fname & ": add error, x = " & CStr(x) & ": , y = " & CStr(y)
            AddLogMsg(msg)
            Throw New ApplicationException(msg)
        End Try
    End Sub

    ' функция которая по признакам точек обновляет массив стабильных точек
    Private Sub UpdateStablePointsList()
        Dim i As Integer
        ReDim FStablePoints(0)
        FkStablePoints = 0

        ' заполняем массив - первая и последние точки там всегда есть по умолчанию
        For i = 0 To FkPoint - 1
            If FPoints(i).stable Or (i = 0) Or i = (FkPoint - 1) Then
                ReDim Preserve FStablePoints(FkStablePoints)
                FStablePoints(FkStablePoints) = FPoints(i).x
                FkStablePoints += 1
            End If
        Next i
    End Sub

    'Public Sub PrintPoints()
    '    Dim i As Integer
    '    For i = 0 To FkPoint - 1
    '        'Debug.Print "i" = i; "x = "; FPoints(i).x; " "; "y = "; FPoints(i).y
    '        Debug.Print FPoints(i).x & " " & FPoints(i).y
    'Next i
    'End Sub

    '    Public Sub PrintValXY(ByVal x As Double)
    '        Dim y As Double

    '        y = getPoint(x)
    '        Debug.Print "F(" + CStr(x) + ") = " + CStr(y)
    'End Sub

    '    Public Sub PrintInterval(ByVal x As Double)
    '        Dim S As String
    '        Dim n As Integer

    '        n = getFirstPointNo(x)
    '        S = CStr(FPoints(n).x) + " (" + CStr(x) + ") " + CStr(FPoints(n + 1).x)
    '        Debug.Print S
    'End Sub

    ' метод который позволяет получить кривую с заданным количеством точек
    Public Function ClonePointsToNum(num_points As Integer) As CInterpolation
        Dim outCurve As New CInterpolation  ' определили новую кривую
        Dim i As Integer
        Dim xPoint As Double, DX As Double
        Dim NumToAdd As Integer
        Dim AddedStablePoints As Integer

        Const eps = 0.01

        outCurve.xName = xName
        outCurve.yName = yName

        If Me.Num_points <= 1 Then
            AddLogMsg("CInterpolation.ClonePointsToNum: error - trying to populate one point curve. curve name: " & note)
            outCurve.AddPoint(0, 0)
            ClonePointsToNum = outCurve
            Exit Function
        End If
        ' добавим все стабильные точки в результирующую кривую
        For i = 1 To FkStablePoints
            xPoint = FStablePoints(i - 1)
            outCurve.AddPoint(xPoint, GetPoint(xPoint))
            If isStepFunction And xPoint > 0 Then outCurve.AddPoint(xPoint - eps, GetPoint(xPoint - eps))
        Next i

        AddedStablePoints = outCurve.Num_points
        ' найдем точки равномерного распределения
        NumToAdd = num_points - AddedStablePoints   ' количество точек, которые надо добавить  концы отрезков уже добавлены
        If NumToAdd <= 0 Then
            ClonePointsToNum = outCurve
            Exit Function
        End If
        DX = (Maxx() - Minx()) / (NumToAdd + 1)      ' приращение - ориетировочное расстояние между точками которые добавляем
        ' добавим недостающие точки
        For i = 1 To NumToAdd
            xPoint = Minx() + DX * i
            outCurve.AddPoint(xPoint, GetPoint(xPoint))  ' добавляем точку в выходной массив
        Next i
        ' может так получится, что стабильные точки совпадают
        While outCurve.Num_points < num_points And outCurve.Num_points > 1
            Call outCurve.DivMaxL()
        End While
        ClonePointsToNum = outCurve
    End Function

    ' функция разделяет максимальный отрезок пополам
    Public Sub DivMaxL()
        Dim xNew, yNew As Double
        Dim maxL As Double
        Dim i As Integer, MaxI As Integer

        MaxI = 0
        maxL = 0
        For i = FPoints.GetLowerBound(0) + 1 To FPoints.GetUpperBound(0)
            If maxL < (FPoints(i).x - FPoints(i - 1).x) Then
                maxL = (FPoints(i).x - FPoints(i - 1).x)
                MaxI = i
            End If
        Next i

        If MaxI > 0 Then
            xNew = FPoints(MaxI - 1).x + (FPoints(MaxI).x - FPoints(MaxI - 1).x) / 2
            yNew = GetPoint(xNew)
            AddPoint(xNew, yNew)
        End If
    End Sub

    Public Function ConvertPointsToNum(num_points As Integer) As Boolean
        ' функция преобразует кривую к кривой такой же с заданным количеством точек (пока линейная интерполяция)
        Dim i As Integer
        Dim MaxL As Double
        Dim MaxI As Integer
        Dim xNew As Double, yNew As Double
        MaxI = 0

        If FkPoint < num_points Then  ' тут надо добавлять точки
            Do
                MaxL = 0
                For i = FPoints.GetLowerBound(0) + 1 To FPoints.GetUpperBound(0)
                    If MaxL < (FPoints(i).x - FPoints(i - 1).x) Then
                        MaxL = (FPoints(i).x - FPoints(i - 1).x)
                        MaxI = i
                    End If
                Next i

                xNew = FPoints(MaxI - 1).x + (FPoints(MaxI).x - FPoints(MaxI - 1).x) / 2
                yNew = GetPoint(xNew)
                AddPoint(xNew, yNew)

            Loop Until FkPoint = num_points
            ConvertPointsToNum = True
        Else                        ' тут надо удалять точки
            ConvertPointsToNum = False
        End If
    End Function

    Public Function Transform(Optional ByVal multY As Double = 1, Optional ByVal sumY As Double = 0,
                          Optional ByVal multX As Double = 1, Optional ByVal sumX As Double = 0) As CInterpolation
        ' преобразует кривую с использованием линейного преобразования на плоскости
        Dim i As Integer
        Dim crv As New CInterpolation

        For i = 0 To FkPoint - 1
            crv.AddPoint(FPoints(i).x * multX + sumX, FPoints(i).y * multY + sumY)
            'FPoints(i).y = FPoints(i).y * multY + sumY
            'FPoints(i).x = FPoints(i).x * multX + sumX
        Next i

        Transform = crv
    End Function

    '    Public Sub loadFromVertRange(ByVal RangX As Object,
    '                    Optional ByVal RangY As Object )
    '        ' функция для чтения range в кривую значений. range читаются по вертикали - значения должны быть в строках - столбец должен быть только один
    '        ' должна использоваться для чтения исходных данных с листа
    '        '
    '        Dim i As Integer
    '        Dim NumVal As Integer
    '        Dim x As Double, y As Double
    '        Dim data_in_2_col As Boolean
    '        Dim arrx, arry

    '        Try
    '            Call ClearPoints()
    '            data_in_2_col = IsMissing(RangY)

    '            If data_in_2_col Then
    '                If TypeName(RangX) = "Range" Then
    '                    NumVal = RangX.Rows.Count
    '                    arrx = RangX.Value2
    '                ElseIf IsArray(RangX) Then
    '                    NumVal = UBound(RangX)
    '                    arrx = RangX
    '                End If
    '            Else
    '                If TypeName(RangX) = "Range" And TypeName(RangY) = "Range" Then
    '                    NumVal = MinReal(RangX.Rows.Count, RangY.Rows.Count)
    '                    arrx = RangX.Value2
    '                    arry = RangY.Value2
    '                ElseIf IsArray(RangX) And IsArray(RangY) Then
    '                    NumVal = MinReal(UBound(RangX), UBound(RangY))
    '                    arrx = RangX
    '                    arry = RangY
    '                End If
    '            End If

    '            If NumVal < 0 Then Err.Raise 1, , "Не удалось прочитать кривую"
    '    ' читаем поэлементно, чтобы отсеять пустые ячейки по пути
    '            On Error Resume Next
    '            For i = 1 To NumVal
    '                x = arrx(i, 1)
    '                If data_in_2_col Then
    '                    y = arrx(i, 2)
    '                Else
    '                    y = arry(i, 1)
    '                End If
    '                If (i = 1) Or (x > 0) Then
    '                    If isStepFunction Then
    '                        Me.AddPoint x, y, isStable:=True
    '            Else
    '                        Me.AddPoint x, y, isStable:=False
    '            End If
    '                End If
    '            Next i
    '            Exit Sub
    'err1:
    '            Err.Raise 1, , "loadFromVertRange: Не удалось прочитать кривую"
    'End Sub

    '    Public Sub load_from_range(range As Variant)
    '        ' функция для чтения range [0..N,0..1] в кривую значений.
    '        ' должна использоваться для чтения исходных данных с листа

    '        Dim i As Integer
    '        Dim NumVal As Integer
    '        Dim x As Double, y As Double
    '        Dim arr
    '        Dim C2

    '        Call ClearPoints()

    '        arr = array_num_from_range(range, True)

    '        ' If TypeName(range) = "Range" Then range = range.Value2

    '        ' читаем поэлементно, чтобы отсеять пустые ячейки по пути
    '        On Error Resume Next
    '        C2 = UBound(arr, 2)
    '        If C2 > 2 Then C2 = 2
    '        For i = LBound(arr, 1) To UBound(arr, 1)
    '            x = arr(i, 1)
    '            y = arr(i, C2)
    '            If isStepFunction Then
    '                Me.AddPoint x, y, isStable:=True
    '        Else
    '                Me.AddPoint x, y, isStable:=False
    '        End If

    '        Next i
    '    End Sub

    Public Function CutByValue(val As Double) As CInterpolation
        Dim i As Integer
        Dim pcur As New CInterpolation
        For i = 1 To Num_points()

            If PointX(i) > val Then
                pcur.AddPoint(PointX(i), PointY(i))
            End If
        Next i
        pcur.AddPoint(val, GetPoint(val))
        pcur.AddPoint(0, GetPoint(val))
        CutByValue = pcur
    End Function
    Public Function CutByValueTrajectory(Optional cut_top_value As Double = 1.0E+20,
                           Optional cut_bottom_value As Double = -1.0E-20) As CInterpolation

        Dim i As Integer
        Dim j As Integer
        Dim FPts() As TDPoint

        j = 1
        For i = 1 To Num_points()

            If PointX(i) < cut_top_value And PointX(i) > cut_bottom_value Then
                If j = 1 And i > 1 And cut_bottom_value < FPoints(i - 1).x Then
                    ReDim Preserve FPts(j)
                    FPts(j - 1).x = cut_bottom_value
                    FPts(j - 1).y = GetPoint(cut_bottom_value)
                    FPts(j - 1).stable = False
                    j = j + 1
                End If

                ReDim Preserve FPts(j)
                FPts(j - 1) = FPoints(i - 1)
                j = j + 1

            End If
        Next i

        If cut_top_value < FPoints(i - 2).x Then
            ReDim Preserve FPts(j - 1)
            FPts(j - 1).x = cut_top_value
            FPts(j - 1).y = GetPoint(cut_top_value)
            FPts(j - 1).stable = False
            j = j + 1
        End If
        If j < 3 Then
            AddLogMsg("CInterpolation.CutByValue: too little points after cut = " & CStr(j - 1))
        End If
        FPoints = FPts
        FkPoint = j - 1
        Call UpdateStablePointsList()
    End Function

    Public Function CutByCurve(crv As CInterpolation) As CInterpolation
        ' обрезание кривой с использованием другой кривой
        Dim i As Integer
        Dim J1, J2 As Integer
        J1 = 0
        J2 = 0
        Dim pcur As New CInterpolation
        Dim crv_min As CInterpolation
        Dim crv_val As Double
        Dim val As Double
        For i = 1 To Num_points()
            crv_val = crv.GetPoint(PointX(i))
            If PointY(i) > crv_val Then
                pcur.AddPoint(PointX(i), PointY(i))
                J1 += 1
            Else
                pcur.AddPoint(PointX(i), crv_val)
                J2 += 1
            End If
        Next i
        If J1 > 0 And J2 > 0 Then
            ' for sure there is an intersection - need to find and add it
            crv_min = SubtractCurve(crv)
            i = crv_min.FindSolutions(0)
            If i = 1 Then
                val = crv_min.SolutionPointX(1)
            Else
            End If
            ' adding Hdyn point as stable - to make sure to have pretty charts later
            pcur.AddPoint(val, GetPoint(val), isStable:=True)
        End If
        CutByCurve = pcur
    End Function



    'Public Sub WriteToRange(RangX As range, Optional RangY As range, Optional ByVal numpt As Integer = 0)
    '    If numpt > 0 And num_points() > 1 Then
    '        Me.ClonePointsToNum(numpt).WriteToRange RangX, RangY, -1
    'ElseIf numpt = -1 Then
    '        WriteToVertRange RangX, RangY
    'Else
    '        numpt = RangX.Rows.Count
    '        Me.ClonePointsToNum(numpt).WriteToRange RangX, RangY, -1
    'End If
    'End Sub

    '    Private Sub WriteToVertRange(RangX As range, RangY As range)
    '        ' позволим кривой записывать себя за заранее данный диапазон ячеек (тут хорошо бы сообразить уместится ли запись или нет - может надо кривую масштабировать?)
    '        ' функция записи кривой на лист excel
    '        On Error GoTo er1
    '        Dim ValX As Double, ValY As Double
    '        Dim NumStr As Integer
    '        Dim i As Integer

    '        RangX.Clear
    '        If Not RangY Is Nothing Then RangY.Clear
    '        NumStr = MinReal(num_points, RangX.Rows.Count)   ' определяем количество элементов в списке. Оно равно числу значений по оси X
    '        RangX.Cells(0, 1) = xName   ' XDescription
    '        If RangY Is Nothing Then
    '            RangX.Cells(0, 2) = yName  'YDescription
    '        Else
    '            RangY.Cells(0, 1) = yName
    '        End If
    '        For i = 1 To NumStr
    '            ValX = pointX(i)
    '            ValY = PointY(i)
    '            RangX.Cells(i, 1) = ValX
    '            If RangY Is Nothing Then
    '                RangX.Cells(i, 2) = ValY
    '            Else
    '                RangY.Cells(i, 1) = ValY
    '            End If
    '        Next i

    '        Exit Sub
    'er1:
    '        Err.Raise kErrWriteDataFromWorksheet, "CInterpolation.WriteToVertRange", "Ошибка, при записи кривой. Точек " & NumStr & " в диапазон ."
    'End Sub

    Public Function SubtractCurve(curv As CInterpolation) As CInterpolation
        ' находит разность двух кривых
        Dim i As Integer
        Dim curve As New CInterpolation

        For i = 1 To Num_points()
            curve.AddPoint(PointX(i), PointY(i) - curv.GetPoint(PointX(i)))
        Next i

        For i = 1 To curv.Num_points
            curve.AddPoint(curv.PointX(i), GetPoint(curv.PointX(i)) - curv.PointY(i))
        Next i
        SubtractCurve = curve
    End Function

    ' инициализация интерполяции данных
    Public Sub Init_interpolation(Optional interpolation_type As String = "Linear",
                              Optional ByVal BoundLType As Integer = 0,
                              Optional ByVal BoundL As Double = 0,
                              Optional ByVal BoundRType As Integer = 0,
                              Optional ByVal BoundR As Double = 0,
                              Optional ByVal CRBoundType As Integer = 0,
                              Optional ByVal CRTension As Double = 0)



        Dim xval() As Double
        Dim yval() As Double
        Dim dval() As Double
        Dim nval As Integer

        Dim i As Integer
        Try
            nval = FkPoint
            ReDim xval(nval)
            ReDim yval(nval)
            ReDim dval(nval)

            For i = 0 To FkPoint - 1
                xval(i) = FPoints(i).x
                yval(i) = FPoints(i).y
                If i = 0 Then
                    yval(i) = FMaxY   ' обнуляет нулевой эллемент массива, поэтому временно сделал так
                End If
                dval(i) = 0 ' todo - need find a way to specify derivatives
            Next i

            Select Case interpolation_type
                Case "Linear"
                    If nval > 2 Then
                        spline1dbuildlinear(xval, yval, nval, spline_interpolant)
                    Else
                    End If
                Case "Cubic"
                    If nval > 2 Then
                        spline1dbuildcubic(xval, yval, nval, BoundLType, BoundL, BoundRType, BoundR, spline_interpolant)
                    Else
                    End If
                Case "Akima"
                    If nval > 5 Then
                        spline1dbuildakima(xval, yval, nval, spline_interpolant)
                    Else
                    End If
                Case "CatmullRom"
                    If nval > 2 Then
                        spline1dbuildcatmullrom(xval, yval, nval, CRBoundType, CRTension, spline_interpolant)
                    Else
                    End If
                Case "Hermite"
                    If nval > 2 Then
                        spline1dbuildhermite(xval, yval, dval, nval, spline_interpolant)
                    Else
                    End If
            End Select
        Catch ex As Exception
            ' унифицированная реакция на ошибочный ввод ключевых параметров класса
            Dim msg As String, fname As String
            fname = "init_interpolation"
            msg = class_name_ & "." & fname & ": spline error, spline type = " & interpolation_type
            AddLogMsg(msg)
            Throw New ApplicationException(msg)
            'Err.Raise kErrPVTinput, class_name_ & "." & fname, msg

        End Try
    End Sub

    ' функция для возврата значения интерполированного сплайнами
    Public Function Get_interpolation_point(ByVal x As Double) As Double

        Try
            Get_interpolation_point = spline1dcalc(spline_interpolant, x)

        Catch ex As Exception
            ' унифицированная реакция на ошибочный ввод ключевых параметров класса
            Dim msg As String, fname As String
            fname = "get_interpolation_point"
            msg = class_name_ & "." & fname & ": spline error, x = " & CStr(x)
            AddLogMsg(msg)
            Throw New ApplicationException(msg)
            'Err.Raise kErrPVTinput, class_name_ & "." & fname, msg

        End Try
    End Function

    Public Sub load_from_range(range(,) As Double)
        ' функция для чтения range [0..N,0..1] в кривую значений.
        ' должна использоваться для чтения исходных данных с листа

        Dim i As Integer
        Dim NumVal As Integer
        Dim X As Double, y As Double
        Dim C2 As Integer

        Call ClearPoints()


        ' If TypeName(range) = "Range" Then range = range.Value2

        ' читаем поэлементно, чтобы отсеять пустые ячейки по пути
        C2 = range.GetUpperBound(2)
        If C2 > 2 Then C2 = 2
        For i = range.GetLowerBound(1) To range.GetUpperBound(1)
            X = range(i, 1)
            y = range(i, C2)
            If isStepFunction Then
                AddPoint(X, y, isStable:=True)
            Else
                AddPoint(X, y, isStable:=False)
            End If

        Next i
    End Sub
End Class

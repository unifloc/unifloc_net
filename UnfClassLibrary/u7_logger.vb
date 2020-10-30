Module u7_logger
    Public Sub AddLogMsg(msg As String)
        msg = "remove later"
    End Sub

    Public Function CheckRanges(ByRef var_value As Double, ByVal var_name As String, ByVal var_min As Double, ByVal var_max As Double,
                                  Optional ByVal out_comment As String = "", Optional ByVal func_name As String = "",
                                  Optional ByVal var_set_default As Boolean = False) As Boolean
        ' функция проверки диапазонов входных параметров для физ мат функци oppump
        CheckRanges = False

        If var_min > var_max Then
            AddLogMsg("CheckRanges:" & func_name & ": wrong check range for " & var_name & ". no check perpformed")
            Exit Function
        End If

        If var_value < var_min Then
            AddLogMsg("CheckRanges:" & func_name & ": variable " & var_name & " = " & var_value & " less than min value = " & var_min & ". " & out_comment)
            If var_set_default Then
                AddLogMsg("CheckRanges:" & func_name & ":for variable " & var_name & " default value set = " & var_min)
                var_value = var_min
            End If

            Exit Function
        End If

        If var_value > var_max Then
            AddLogMsg("CheckRanges:" & func_name & ": variable " & var_name & " = " & var_value & " greater than max value = " & var_max & ". " & out_comment)
            If var_set_default Then
                AddLogMsg("CheckRanges:" & func_name & ":для переменной " & var_name & " default value set = " & var_max)
                var_value = var_max
            End If

            Exit Function

        End If

        CheckRanges = True

    End Function


End Module

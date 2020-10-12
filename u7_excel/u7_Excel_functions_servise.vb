'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
' вспомогательные функции для проведения расчетов из рабочих книг Excel

Option Explicit On


Public Function decode_json_string(json,
                          Optional transpose As Boolean = False,
                          Optional keys_filter,
                          Optional only_values As Boolean = False)
    ' json   - строка содержащая результаты расчета
    ' transpose - выбор вывода в строки или в столбцы
    ' keys_filter - строка с ключами, которые надо вывести
    ' only_values - если = 1 подписи выводиться не будут
    ' результат - закодированная строка
    'description_end

    Dim d As IDictionary
    Dim c As ICollection
    Dim p
    Dim i As Integer
    Dim outarr
    Dim v
    Dim K

    Dim keylist As IDictionary
    Dim arrkeys As Object

    Try


        JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True

        p = ParseJson(json)
        If TypeName(p) = "Dictionary" Then

            d = p

            If keys_filter Then
                arrkeys = Split(keys_filter, ",")
                keylist.Clear()
                For i = LBound(arrkeys) To UBound(arrkeys)
                    If d.Contains(arrkeys(i)) Then
                        keylist.Add(arrkeys(i), d.Item(arrkeys(i)))
                    End If
                Next i
                If keylist.Count > 0 Then
                    d = keylist
                End If
            End If


            If transpose Then
            If only_values Then
                ReDim outarr(1 To 1, 1 To d.Count)
            Else
                ReDim outarr(1 To 2, 1 To d.Count)
            End If
            For i = 1 To d.Count
                K = d.keys(i - 1)
                outarr(1, i) = K
                    If TypeOf d.Values(i - 1) Is Object Then
                        v = d.Values(i - 1)
                    Else
                        v = d.Values(i - 1)
                    End If
                If TypeName(v) = "Collection" Then
                    If only_values Then
                        outarr(1, i) = ConvertToJson(v)
                    Else
                        outarr(2, i) = ConvertToJson(v)
                    End If
                Else

                    If only_values Then
                        outarr(1, i) = v
                    Else
                        outarr(2, i) = v
                    End If
                End If
            Next
        Else
            If only_values Then
                ReDim outarr(1 To d.Count, 1 To 1)
            Else
                ReDim outarr(1 To d.Count, 1 To 2)
            End If
            For i = 1 To d.Count
                K = d.keys(i - 1)

                outarr(i, 1) = d.keys(i - 1)
                    If TypeOf d.Values(i - 1) Is Object Then
                        v = d.Values(i - 1)
                    Else
                        v = d.Values(i - 1)
                    End If
                If TypeName(v) = "Collection" Then
                    If only_values Then
                        outarr(i, 1) = ConvertToJson(v)
                    Else
                        outarr(i, 2) = ConvertToJson(v)
                    End If
                Else
                    If only_values Then
                        outarr(i, 1) = v
                    Else
                        outarr(i, 2) = v
                    End If
                End If
            Next i
        End If
    Else  ' expect collection here

            c = p
            If c.Count = 1 Then
            If TypeName(c.Item(1)) = "Collection" Then
                    outarr = CollectionToArray2D(c.Item(1))
                End If
        Else
            i = 1
            If transpose Then
                ReDim outarr(1 To 1, 1 To c.Count)
                For Each v In c
                    outarr(1, i) = v
                    i = i + 1
                Next
            Else
                ReDim outarr(1 To c.Count, 1 To 1)
                For Each v In c
                    outarr(i, 1) = v
                    i = i + 1
                Next
            End If
        End If
    End If
    decode_json_string = outarr
    Exit Function
err1:
    decode_json_string = "error"
End Function
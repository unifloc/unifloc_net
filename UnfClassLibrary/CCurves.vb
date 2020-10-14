'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
' class for calculated curves managements
' ---------------------------------------------------------
' ���������� �������
' ������ - ������ ���� CInterpolation
'          ������� �� ������ ����� (x,y) ������ ������ �������� x ������������� ���� y
'          ������ ���������� ������� (� �������������� �������)
' ---------------------------------------------------------
Option Explicit On

Public Class CCurves
    Private curves As IDictionary '������� ������ � ������������ �������� 

    Public ReadOnly Property Item(key As String) As CInterpolation
        Get
            If curves.Contains(key) Then
                Item = curves.Item(key)
            Else
                Item = New CInterpolation
                curves.Item(key) = Item
            End If
        End Get
    End Property

    Public WriteOnly Property Item(key As String, valNew As CInterpolation)
        Set
            curves.Item(key) = valNew
            ' for dictionary if key exist it will be overwritten
        End Set
    End Property

    Public Sub ClearPoints()
        Dim crv As CCurves
        For Each crv In curves.Values
            Call crv.ClearPoints
        Next crv
    End Sub

    'Public Sub ClearPoints_unprotected()
    'Dim crv As CCurves
    'For Each crv In curves.Values
    'If Not crv.special Then Call crv.ClearPoints
    'Next crv
    'End Sub
End Class

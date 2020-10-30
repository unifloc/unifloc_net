﻿'=======================================================================================
'Unifloc 7.24  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================


' Функции для упрощения перевода единиц в Excel
'



Module u7_UnitsConversion
    Public Function UC_Temperature_C_to_F(ByVal t_C As Double) As Double

        UC_Temperature_C_to_F = t_C * 9 / 5 + 32

    End Function

    Public Function UC_Temperature_F_to_C(ByVal t_F As Double) As Double

        UC_Temperature_F_to_C = (t_F - 32) * 5 / 9

    End Function

    Public Function UC_Rs_m3m3_to_scfbbl(ByVal rs_m3m3 As Double) As Double

        UC_Rs_m3m3_to_scfbbl = rs_m3m3 * const_convert_m3m3_scfbbl

    End Function

    Public Function UC_Rs_scfbbl_to_m3m3(ByVal Rs_scfbbl As Double) As Double

        UC_Rs_scfbbl_to_m3m3 = Rs_scfbbl * const_convert_scfbbl_m3m3

    End Function

    Public Function UC_pressure_atma_to_psi(ByVal p_atma As Double) As Double

        UC_pressure_atma_to_psi = p_atma * const_convert_atma_psi

    End Function

    Public Function UC_pressure_psi_to_atma(ByVal p_psi As Double) As Double

        UC_pressure_psi_to_atma = p_psi * const_convert_psi_atma

    End Function

    Public Function UC_pressure_atma_to_MPa(ByVal p_atma As Double) As Double

        UC_pressure_atma_to_MPa = p_atma * const_convert_atma_MPa

    End Function

    Public Function UC_pressure_MPa_to_atma(ByVal p_MPa As Double) As Double

        UC_pressure_MPa_to_atma = p_MPa * const_convert_MPa_atma

    End Function

End Module

'=======================================================================================
'Unifloc 7.25  coronav                                     khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
'функции для проведения расчетов из интерфейса Excel
'многофазный поток в трубах и элементах инфраструктуры
Option Explicit On

Public Module u7_Excel_functions_MF
    Public Function MF_calibr_choke_fast(
                                        ByVal qliq_sm3day As Double,
                                        ByVal fw_perc As Double,
                                        ByVal d_choke_mm As Double,
                                        Optional ByVal p_in_atma As Double = -1,
                                        Optional ByVal p_out_atma As Double = -1,
                                        Optional ByVal d_pipe_mm As Double = 70,
                                        Optional ByVal t_choke_C As Double = 20,
                                        Optional ByVal str_PVT As String = UnfClassLibrary.u7_const.PVT_DEFAULT,
                                        Optional ByVal q_gas_sm3day As Double = 0)
        Try
            Dim Choke As UnfClassLibrary.CChoke
            Choke = New UnfClassLibrary.CChoke
            Dim pt As UnfClassLibrary.u7_types.PTtype
            Dim PVT As UnfClassLibrary.CPVT
            Dim out, out_desc

            PVT = PVT_decode_string(str_PVT)
            If PVT.gas_only Then
                MF_calibr_choke_fast = "not implemented yet"
                Exit Function
            End If

            PVT.q_gas_free_sm3day = q_gas_sm3day
            Choke.fluid = PVT
            Choke.Class_Initialize()
            Choke.fluid.qliq_sm3day = qliq_sm3day
            Choke.fluid.Fw_fr = fw_perc / 100
            Choke.d_down_m = d_pipe_mm / 1000
            Choke.d_up_m = d_pipe_mm / 1000
            Choke.d_choke_m = d_choke_mm / 1000

            If p_in_atma > p_out_atma And p_out_atma >= 1 Then
                Call Choke.calc_choke_calibration(p_in_atma, p_out_atma, t_choke_C)
                out = Choke.c_calibr_fr
                out_desc = "c_calibr_fr"
            End If

            'Dim new_array(1) As Object
            'new_array(0) = (out, p_in_atma, p_out_atma, t_choke_C, Choke.c_calibr_fr)
            'new_array(1) = (out_desc, "p_intake_atma", "p_out_atma", "t_choke_C", "c_calibr_fr")
            'MF_calibr_choke_fast = new_array
            MF_calibr_choke_fast = {out, p_in_atma, p_out_atma, t_choke_C, Choke.c_calibr_fr}
            'MF_calibr_choke_fast = Join(MF_calibr_choke_fast)
            Exit Function

        Catch ex As Exception
            MF_calibr_choke_fast = -1
            Dim errmsg As String
            errmsg = "Error:MF_calibr_choke_fast:"
            Throw New ApplicationException(errmsg)

        End Try
    End Function

    Public Function MF_q_choke_sm3day(ByVal fw_perc As Double,
                                      ByVal d_choke_mm As Double,
                                      ByVal p_in_atma As Double,
                                      ByVal p_out_atma As Double,
                                      Optional ByVal d_pipe_mm As Double = 70,
                                      Optional ByVal t_choke_C As Double = 20,
                                      Optional ByVal c_calibr_fr As Double = 1,
                                      Optional ByVal str_PVT As String = UnfClassLibrary.u7_const.PVT_DEFAULT,
                                      Optional ByVal q_gas_sm3day As Double = 0)
        Try
            Dim choke As New UnfClassLibrary.CChoke
            choke = New UnfClassLibrary.CChoke
            Dim PVT As UnfClassLibrary.CPVT
            Dim q As Double

            PVT = PVT_decode_string(str_PVT)
            choke.fluid = PVT
            choke.Class_Initialize()
            choke.d_down_m = d_pipe_mm / 1000
            choke.d_up_m = d_pipe_mm / 1000
            choke.d_choke_m = d_choke_mm / 1000
            choke.fluid.Fw_perc = fw_perc
            choke.fluid.q_gas_free_sm3day = q_gas_sm3day
            choke.c_calibr_fr = c_calibr_fr

            If PVT.gas_only Then
                MF_q_choke_sm3day = -1
            Else
                q = choke.calc_choke_qliq_sm3day(p_in_atma, p_out_atma, t_choke_C)
                MF_q_choke_sm3day = {q, p_in_atma, p_out_atma, t_choke_C}
            End If

        Catch ex As Exception
            MF_q_choke_sm3day = -1
            Dim errmsg As String
            errmsg = "Error:MF_q_choke_sm3day:"
            Throw New ApplicationException(errmsg)
        End Try
    End Function
End Module
﻿'=======================================================================================
'Unifloc 7.24  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
'Модуль расчетов физико-химических свойств нефтей
'
'префикс unf_ символизиует что это функции Unifloc для внутреннего использования при кодировании
' converted 16/05/2020


Module u7_PVT

    Function Unf_pvt_viscosity_dead_oil_Standing_cP(ByVal Temperature_K As Double,
                                                    ByVal gamma_oil As Double) As Double

        Unf_pvt_viscosity_dead_oil_Standing_cP = (0.32 + 1.8 * 10 ^ 7 / (141.5 / gamma_oil - 131.5) ^ 4.53) * (360 / (1.8 * Temperature_K - 260)) ^ (10 ^ (0.43 + 8.33 / (141.5 / gamma_oil - 131.5)))
    End Function

    Function Unf_pvt_compressibility_oil_VB_1atm(ByVal rs_m3m3 As Double,
                                                 ByVal gamma_gas As Double,
                                                 ByVal t_K As Double,
                                                 ByVal gamma_oil As Double,
                                                 ByVal p_MPa As Double) As Double

        Unf_pvt_compressibility_oil_VB_1atm = (28.1 * rs_m3m3 + 30.6 * t_K - 1180 * gamma_gas + 1784 / gamma_oil - 10910) / (100000 * p_MPa)
    End Function

    Function Unf_pvt_pb_Standing_MPa(ByVal rsb_m3m3 As Double,
                                     ByVal gamma_gas As Double,
                                     ByVal Temperature_K As Double,
                                     ByVal gamma_oil As Double) As Double

        Const Min_rsb As Double = 1.8
        Dim rsb_old As Double

        rsb_old = rsb_m3m3
        If (rsb_m3m3 < Min_rsb) Then
            rsb_m3m3 = Min_rsb
        End If

        Dim yg As Double

        yg = 1.225 + 0.001648 * Temperature_K - 1.769 / gamma_oil
        Unf_pvt_pb_Standing_MPa = 0.5197 * (rsb_m3m3 / gamma_gas) ^ 0.83 * 10 ^ yg
        If (rsb_old < Min_rsb) Then
            Unf_pvt_pb_Standing_MPa = (Unf_pvt_pb_Standing_MPa - 0.1013) * rsb_old / Min_rsb + 0.1013
        End If
    End Function

    Function Unf_pvt_FVF_Saturated_Oil_Standing_m3m3(ByVal rs_m3m3 As Double,
                                                     ByVal gamma_gas As Double,
                                                     ByVal Temperature_K As Double,
                                                     ByVal gamma_oil As Double) As Double
        Dim F As Double

        F = 5.615 * rs_m3m3 * (gamma_gas / gamma_oil) ^ 0.5 + 2.25 * Temperature_K - 575
        Unf_pvt_FVF_Saturated_Oil_Standing_m3m3 = 0.972 + 0.000147 * F ^ 1.175
    End Function

    Function Unf_pvt_FVF_above_pb_Standing_m3m3(ByVal p_MPa As Double,
                                                ByVal pb_MPa As Double,
                                                ByVal Oil_Compressibility As Double,
                                                ByVal FVF_Saturated_Oil As Double) As Double
        If p_MPa <= pb_MPa Then
            Unf_pvt_FVF_above_pb_Standing_m3m3 = FVF_Saturated_Oil
        Else
            Unf_pvt_FVF_above_pb_Standing_m3m3 = FVF_Saturated_Oil * Math.Exp(Oil_Compressibility * (pb_MPa - p_MPa))
        End If
    End Function

    Function Unf_pvt_viscosity_oil_Standing_cP(ByVal rs_m3m3 As Double,
                                               ByVal Dead_oil_viscosity As Double,
                                               ByVal p_MPa As Double,
                                               ByVal pb_MPa As Double) As Double
        Dim A As Double, b As Double
        A = 5.6148 * rs_m3m3 * (0.1235 * 10 ^ (-5) * rs_m3m3 - 0.00074)
        b = 0.68 / 10 ^ (0.000484 * rs_m3m3) + 0.25 / 10 ^ (0.006176 * rs_m3m3) + 0.062 / 10 ^ (0.021 * rs_m3m3)

        Unf_pvt_viscosity_oil_Standing_cP = 10 ^ A * Dead_oil_viscosity ^ b

        If pb_MPa < p_MPa Then
            Unf_pvt_viscosity_oil_Standing_cP += 0.14504 * (p_MPa - pb_MPa) * (0.024 * Unf_pvt_viscosity_oil_Standing_cP ^ 1.6 + 0.038 * Unf_pvt_viscosity_oil_Standing_cP ^ 0.56)
        End If
    End Function

    Function Unf_pvt_density_oil_Standing_kgm3(ByVal rs_m3m3 As Double,
                                               ByVal gamma_gas As Double,
                                               ByVal gamma_oil As Double,
                                               ByVal p_MPa As Double,
                                               ByVal FVF_m3m3 As Double,
                                               ByVal BP_p_MPa As Double,
                                               ByVal Compressibility_1MPa As Double) As Double
        Unf_pvt_density_oil_Standing_kgm3 = (1000 * gamma_oil + 1.224 * gamma_gas * rs_m3m3) / FVF_m3m3
        If p_MPa > BP_p_MPa Then
            Unf_pvt_density_oil_Standing_kgm3 *= Math.Exp(Compressibility_1MPa * (p_MPa - BP_p_MPa))
        End If
    End Function

    Function Unf_pvt_GOR_Standing_m3m3(ByVal p_MPa As Double,
                                       ByVal gamma_gas As Double,
                                       ByVal Temperature_K As Double,
                                       ByVal gamma_oil As Double) As Double
        Dim yg As Double
        yg = 1.225 + 0.001648 * Temperature_K - 1.769 / gamma_oil
        Unf_pvt_GOR_Standing_m3m3 = gamma_gas * (1.92 * p_MPa / 10 ^ yg) ^ 1.204
    End Function

    Function Unf_pvt_pb_Valko_McCain_MPa(ByVal rsb_m3m3 As Double,
                                         ByVal gamma_gas As Double,
                                         ByVal Temperature_K As Double,
                                         ByVal gamma_oil As Double) As Double
        Const Min_rsb As Double = 1.8
        Const Max_rsb As Double = 800

        Dim rsb_old As Double
        Dim API As Double, z1 As Double, z2 As Double, z3 As Double, z4 As Double, z As Double, lnpb As Double

        rsb_old = rsb_m3m3
        If (rsb_m3m3 < Min_rsb) Then
            rsb_m3m3 = Min_rsb
        End If

        If (rsb_m3m3 > Max_rsb) Then
            rsb_m3m3 = Max_rsb
        End If

        API = 141.5 / gamma_oil - 131.5
        z1 = -4.814074834 + 0.7480913 * Math.Log(rsb_m3m3) + 0.1743556 * Math.Log(rsb_m3m3) ^ 2 - 0.0206 * Math.Log(rsb_m3m3) ^ 3
        z2 = 1.27 - 0.0449 * API + 4.36 * 10 ^ (-4) * API ^ 2 - 4.76 * 10 ^ (-6) * API ^ 3
        z3 = 4.51 - 10.84 * gamma_gas + 8.39 * gamma_gas ^ 2 - 2.34 * gamma_gas ^ 3
        z4 = -7.2254661 + 0.043155 * Temperature_K - 8.5548 * 10 ^ (-5) * Temperature_K ^ 2 + 6.00696 * 10 ^ (-8) * Temperature_K ^ 3
        z = z1 + z2 + z3 + z4
        lnpb = 2.498006 + 0.713 * z + 0.0075 * z ^ 2

        Unf_pvt_pb_Valko_McCain_MPa = 2.718282 ^ lnpb

        If (rsb_old < Min_rsb) Then
            Unf_pvt_pb_Valko_McCain_MPa = (Unf_pvt_pb_Valko_McCain_MPa - 0.1013) * rsb_old / Min_rsb + 0.1013
        End If

        If (rsb_old > Max_rsb) Then
            Unf_pvt_pb_Valko_McCain_MPa = (Unf_pvt_pb_Valko_McCain_MPa - 0.1013) * rsb_old / Max_rsb + 0.1013
        End If
    End Function

    Function Unf_pvt_GOR_Velarde_m3m3(ByVal p_MPa As Double,
                                      ByVal pb_MPa As Double,
                                      ByVal gamma_gas As Double,
                                      ByVal Temperature_K As Double,
                                      ByVal gamma_oil As Double,
                                      ByVal rsb_m3_m3 As Double) As Double
        Dim API As Double
        API = 141.5 / gamma_oil - 131.5
        Const MaxRs As Double = 800
        If (pb_MPa > Unf_pvt_pb_Valko_McCain_MPa(MaxRs, gamma_gas, Temperature_K, gamma_oil)) Then
            If p_MPa < pb_MPa Then
                Unf_pvt_GOR_Velarde_m3m3 = (rsb_m3_m3) * (p_MPa / pb_MPa)
            Else
                Unf_pvt_GOR_Velarde_m3m3 = rsb_m3_m3
            End If
            Exit Function
        End If
        Dim Pr As Double
        If (pb_MPa > 0) Then
            Pr = (p_MPa - 0.101) / (pb_MPa)
        Else
            Pr = 0
        End If
        If Pr <= 0 Then
            Unf_pvt_GOR_Velarde_m3m3 = 0
            Exit Function
        End If
        Dim a_0 As Double
        Dim a_1 As Double
        Dim a_2 As Double
        Dim a_3 As Double
        Dim a_4 As Double
        Dim A1 As Double
        Dim b_0 As Double
        Dim b_1 As Double
        Dim b_2 As Double
        Dim b_3 As Double
        Dim b_4 As Double
        Dim A2 As Double
        Dim c_0 As Double
        Dim c_1 As Double
        Dim c_2 As Double
        Dim c_3 As Double
        Dim c_4 As Double
        Dim A3 As Double, Rsr As Double
        If Pr >= 1 Then
            Unf_pvt_GOR_Velarde_m3m3 = rsb_m3_m3
        Else
            'If Pr < 1 Then
            a_0 = 1.8653 * 10 ^ (-4)
            a_1 = 1.672608
            a_2 = 0.92987
            a_3 = 0.247235
            a_4 = 1.056052

            A1 = a_0 * gamma_gas ^ a_1 * API ^ a_2 * (1.8 * Temperature_K - 460) ^ a_3 * pb_MPa ^ a_4

            b_0 = 0.1004
            b_1 = -1.00475
            b_2 = 0.337711
            b_3 = 0.132795
            b_4 = 0.302065

            A2 = b_0 * gamma_gas ^ b_1 * API ^ b_2 * (1.8 * Temperature_K - 460) ^ b_3 * pb_MPa ^ b_4

            c_0 = 0.9167
            c_1 = -1.48548
            c_2 = -0.164741
            c_3 = -0.09133
            c_4 = 0.047094

            A3 = c_0 * gamma_gas ^ c_1 * API ^ c_2 * (1.8 * Temperature_K - 460) ^ c_3 * pb_MPa ^ c_4

            Rsr = A1 * Pr ^ A2 + (1 - A1) * Pr ^ A3

            Unf_pvt_GOR_Velarde_m3m3 = Rsr * rsb_m3_m3
        End If
    End Function

    Function Unf_pvt_FVF_McCain_m3m3(ByVal rs_m3m3 As Double,
                                     ByVal gamma_gas As Double,
                                     ByVal STO_density_kg_m3 As Double,
                                     ByVal Reservoir_oil_density_kg_m3 As Double) As Double
        Unf_pvt_FVF_McCain_m3m3 = (STO_density_kg_m3 + 1.22117 * rs_m3m3 * gamma_gas) / Reservoir_oil_density_kg_m3
    End Function

    Function Unf_pvt_density_McCain_kgm3(ByVal p_MPa As Double,
                                         ByVal gamma_gas As Double,
                                         ByVal Temperature_K As Double,
                                         ByVal gamma_oil As Double,
                                         ByVal Rs_m3_m3 As Double,
                                         ByVal BP_p_MPa As Double,
                                         ByVal Compressibility As Double) As Double
        Dim API As Double, ropo As Double, pm As Double, pmmo As Double, epsilon As Double
        Dim i As Integer, counter As Integer
        Dim a0 As Double, A1 As Double, A2 As Double, A3 As Double, a4 As Double, a5 As Double
        Dim roa As Double
        API = 141.5 / gamma_oil - 131.5
        'limit input range to Rs = 800, Pb =1000
        If (Rs_m3_m3 > 800) Then
            Rs_m3_m3 = 800
            BP_p_MPa = Unf_pvt_pb_Valko_McCain_MPa(Rs_m3_m3, gamma_gas, Temperature_K, gamma_oil)
        End If
        ropo = 845.8 - 0.9 * Rs_m3_m3
        pm = ropo
        pmmo = 0
        epsilon = 0.000001
        i = 0
        counter = 0
        Const MaxIter As Integer = 100
        While (Math.Abs(pmmo - pm) > epsilon And counter < MaxIter)
            i += 1
            pmmo = pm
            a0 = -799.21
            A1 = 1361.8
            A2 = -3.70373
            A3 = 0.003
            a4 = 2.98914
            a5 = -0.00223

            roa = a0 + A1 * gamma_gas + A2 * gamma_gas * ropo + A3 * gamma_gas * ropo ^ 2 + a4 * ropo + a5 * ropo ^ 2
            ropo = (Rs_m3_m3 * gamma_gas + 818.81 * gamma_oil) / (0.81881 + Rs_m3_m3 * gamma_gas / roa)
            pm = ropo
            counter += 1
            ' Debug.Assert counter < 20
        End While


        Dim dpp As Double, pbs As Double, dPT As Double
        If p_MPa <= BP_p_MPa Then
            dpp = (0.167 + 16.181 * (10 ^ (-0.00265 * pm))) * (2.32328 * p_MPa) - 0.16 * (0.299 + 263 * (10 ^ (-0.00376 * pm))) * (0.14503774 * p_MPa) ^ 2
            pbs = pm + dpp
            dPT = (0.04837 + 337.094 * pbs ^ (-0.951)) * (1.8 * Temperature_K - 520) ^ 0.938 - (0.346 - 0.3732 * (10 ^ (-0.001 * pbs))) * (1.8 * Temperature_K - 520) ^ 0.475
            pm = pbs - dPT
            Unf_pvt_density_McCain_kgm3 = pm
        Else
            dpp = (0.167 + 16.181 * (10 ^ (-0.00265 * pm))) * (2.32328 * BP_p_MPa) - 0.16 * (0.299 + 263 * (10 ^ (-0.00376 * pm))) * (0.14503774 * BP_p_MPa) ^ 2
            pbs = pm + dpp
            dPT = (0.04837 + 337.094 * pbs ^ (-0.951)) * (1.8 * Temperature_K - 520) ^ 0.938 - (0.346 - 0.3732 * (10 ^ (-0.001 * pbs))) * (1.8 * Temperature_K - 520) ^ 0.475
            pm = pbs - dPT
            Unf_pvt_density_McCain_kgm3 = pm * Math.Exp(Compressibility * (p_MPa - BP_p_MPa))
        End If
    End Function

    Function Unf_pvt_viscosity_dead_oil_Beggs_Robinson_cP(ByVal Temperature_K As Double,
                                                          ByVal gamma_oil As Double) As Double
        Dim x As Double
        x = (1.8 * Temperature_K - 460) ^ (-1.163) * Math.Exp(13.108 - 6.591 / gamma_oil)
        Unf_pvt_viscosity_dead_oil_Beggs_Robinson_cP = 10 ^ (x) - 1
    End Function

    Function Unf_pvt_viscosity_saturated_oil_Beggs_Robinson_cP(ByVal GOR_pb_m3m3 As Double,
                                                               ByVal Dead_oil_viscosity As Double) As Double
        Dim A As Double
        Dim b As Double
        A = 10.715 * (5.615 * GOR_pb_m3m3 + 100) ^ (-0.515)
        b = 5.44 * (5.615 * GOR_pb_m3m3 + 150) ^ (-0.338)
        Unf_pvt_viscosity_saturated_oil_Beggs_Robinson_cP = A * (Dead_oil_viscosity) ^ b
    End Function

    Function Unf_pvt_viscosity_oil_Vasquez_Beggs_cP(ByVal Saturated_oil_viscosity As Double,
                                                    ByVal p_MPa As Double,
                                                    ByVal BP_p_MPa As Double) As Double
        Dim C1 As Double
        Dim C2 As Double
        Dim C3 As Double
        Dim C4 As Double
        Dim M As Double

        C1 = 957
        C2 = 1.187
        C3 = -11.513
        C4 = -0.01302
        M = C1 * p_MPa ^ C2 * Math.Exp(C3 + C4 * p_MPa)
        Unf_pvt_viscosity_oil_Vasquez_Beggs_cP = Saturated_oil_viscosity * (p_MPa / BP_p_MPa) ^ M
    End Function

    Function Unf_pvt_viscosity_oil_Beggs_Robinson_Vasques_Beggs_cP(ByVal rs_m3m3 As Double,
                                                                   ByVal GOR_pb_m3m3 As Double,
                                                                   ByVal p_MPa As Double,
                                                                   ByVal BP_p_MPa As Double,
                                                                   ByVal Dead_oil_viscosity As Double) As Double
        If (p_MPa < BP_p_MPa) Then 'saturated
            Unf_pvt_viscosity_oil_Beggs_Robinson_Vasques_Beggs_cP = Unf_pvt_viscosity_saturated_oil_Beggs_Robinson_cP(rs_m3m3, Dead_oil_viscosity)
        Else 'undersaturated
            Unf_pvt_viscosity_oil_Beggs_Robinson_Vasques_Beggs_cP = Unf_pvt_viscosity_oil_Vasquez_Beggs_cP(
     Unf_pvt_viscosity_saturated_oil_Beggs_Robinson_cP(GOR_pb_m3m3, Dead_oil_viscosity), p_MPa, BP_p_MPa)
        End If

    End Function

    Function Unf_pvt_viscosity_Grace_cP(ByVal p_MPa As Double,
                                        ByVal pb_MPa As Double,
                                        ByVal rho_kgm3 As Double,
                                        ByVal BP_rho_kgm3 As Double) As Double
        Dim density As Double
        Dim Bubblepoint_Density As Double
        Dim rotr As Double
        Dim MU As Double
        Dim robtr As Double
        Dim M As Double
        density = rho_kgm3 * 0.06243
        Bubblepoint_Density = BP_rho_kgm3 * 0.06243
        rotr = 0.0008 * density ^ 3 - 0.1017 * density ^ 2 + 4.3344 * density - 63.001
        MU = Math.Exp(0.0281 * rotr ^ 3 - 0.0447 * rotr ^ 2 + 1.2802 * rotr + 0.0359)
        If pb_MPa < p_MPa Then

            robtr = -68.1067 * Math.Log(Bubblepoint_Density) ^ 3 + 783.2173 * Math.Log(Bubblepoint_Density) ^ 2 - 2992.2353 * Math.Log(Bubblepoint_Density) + 3797.6
            M = Math.Exp(0.1124 * robtr ^ 3 - 0.0618 * robtr ^ 2 + 0.7356 * robtr + 2.3328)
            MU *= (rho_kgm3 / BP_rho_kgm3) ^ M

        End If
        Unf_pvt_viscosity_Grace_cP = MU
    End Function



    ' ==================================================================
    ' PVT gas
    ' ==================================================================
    Function Unf_pvt_viscosity_gas_cP(ByVal t_K As Double,
                                      ByVal p_MPa As Double,
                                      ByVal z As Double,
                                      ByVal GammaGas As Double) As Double
        ' rnt 20150303
        ' расчет вязкости газа after Lee    http://petrowiki.org/Gas_viscosity
        ' похоже, что отсюда
        ' Lee, A.L., Gonzalez, M.H., and Eakin, B.E. 1966. The Viscosity of Natural Gases. J Pet Technol 18 (8): 997–1000. SPE-1340-PA. http://dx.doi.org/10.2118/1340-PA
        '

        Dim r As Double, mwg As Double, gd As Double, A As Double, b As Double, c As Double

        r = 1.8 * t_K
        mwg = 28.966 * GammaGas
        gd = p_MPa * mwg / (z * t_K * 8.31)
        A = (9.379 + 0.01607 * mwg) * r ^ 1.5 / (209.2 + 19.26 * mwg + r)
        b = 3.448 + 986.4 / r + 0.01009 * mwg
        c = 2.447 - 0.2224 * b
        Unf_pvt_viscosity_gas_cP = 0.0001 * A * Math.Exp(b * gd ^ c)

    End Function


    Public Function Unf_pvt_dZdt(ByVal t_K As Double,
                                 ByVal p_MPa As Double,
                                 ByVal GammaGas As Double,
                        Optional ByVal z_cor As Z_CORRELATION = Z_CORRELATION.z_Kareem,
                        Optional ByRef z_val As Double = 1) As Double
        Dim z1 As Double
        Dim z2 As Double
        Dim dtz As Double
        dtz = 0.1   ' dangerous to reduce for dranchuk correlation

        z1 = Unf_pvt_Zgas_d(t_K, p_MPa, GammaGas, z_cor)
        z2 = Unf_pvt_Zgas_d(t_K + dtz, p_MPa, GammaGas, z_cor)
        Unf_pvt_dZdt = (z2 - z1) / dtz
        z_val = z1
    End Function


    Public Function Unf_pvt_dZdp(ByVal t_K As Double,
                                 ByVal p_MPa As Double,
                                 ByVal GammaGas As Double,
                        Optional ByVal z_cor As Z_CORRELATION = Z_CORRELATION.z_Kareem,
                        Optional ByRef z_val As Double = 1) As Double
        Dim z1 As Double
        Dim z2 As Double
        Dim dp As Double
        dp = 0.01 ' dangerous to reduce for dranchuk correlation

        z1 = Unf_pvt_Zgas_d(t_K, p_MPa, GammaGas, z_cor)
        z2 = Unf_pvt_Zgas_d(t_K, p_MPa + dp, GammaGas, z_cor)
        Unf_pvt_dZdp = (z2 - z1) / dp
        z_val = z1
    End Function



    Public Function Unf_pvt_Zgas_d(ByVal t_K As Double, ByVal p_MPa As Double, ByVal GammaGas As Double, Optional ByVal z_cor As Z_CORRELATION = Z_CORRELATION.z_Kareem) As Double
        ' calculus of z factor
        ' rnt 20150303 не желательно использовать значение корреляции отличное от 0
        ' http://petrowiki.org/Real_gases
        ' расчет по Дранчуку или по Саттону
        ' Dranchuk, P.M. and Abou-Kassem, H. 1975. Calculation of Z Factors For Natural Gases Using Equations of State. J Can Pet Technol 14 (3): 34. PETSOC-75-03-03. http://dx.doi.org/10.2118/75-03-03
        ' Sutton, R.P. 1985. Compressibility Factors for High-Molecular-Weight Reservoir Gases. Presented at the SPE Annual Technical Conference and Exhibition, Las Vegas, Nevada, USA, 22-26 September. SPE-14265-MS. http://dx.doi.org/10.2118/14265-MS

        Dim T_pc As Double
        Dim p_pc As Double
        Dim z As Double
        Try
            If z_cor = Z_CORRELATION.z_Dranchuk Then
                T_pc = PseudoTemperatureStanding(GammaGas)
                p_pc = PseudoPressureStanding(GammaGas)
                z = ZFactorDranchuk(t_K / T_pc, p_MPa / p_pc)
            ElseIf z_cor = Z_CORRELATION.z_Kareem Then
                T_pc = PseudoTemperatureStanding(GammaGas)
                p_pc = PseudoPressureStanding(GammaGas)
                z = ZFactor2015_Kareem(t_K / T_pc, p_MPa / p_pc)
            Else
                T_pc = PseudoTemperature(GammaGas)
                p_pc = PseudoPressure(GammaGas)
                z = ZFactor(t_K / T_pc, p_MPa / p_pc)
            End If
            Unf_pvt_Zgas_d = z '* aaaa
            Exit Function
        Catch ex As Exception
            Dim msg As String
            msg = "Unf_pvt_Zgas_d: error with " & "t_K = " & CStr(t_K) & ", P_MPa = " & CStr(p_MPa) & ", GammaGas = " & CStr(GammaGas) & ", z_cor = " & CStr(z_cor) & ": " & ex.Message

            Throw New ApplicationException(msg)

        End Try
    End Function


    Public Function Unf_pvt_Bg_m3m3(ByVal t_C As Double, ByVal p_atma As Double, ByVal GammaGas As Double, Optional correlation As Z_CORRELATION = Z_CORRELATION.z_Kareem) As Double
        'function calculates gas formation volume factor

        ' t_С   -   temprature, C
        ' p_atma   -   pressure, atma
        ' gamma_g   - specific gas density
        ' correlation
        '    0 - using Dranchuk and Abou-Kassem correlation
        '    else - using Sutton correlation for the pseudocritical properties of hydrocarbon mixtures
        Dim t_K As Double
        Dim p_MPa As Double
        Dim z As Double
        '  Debug.Assert GammaGas > 0.5
        t_K = t_C + 273
        p_MPa = p_atma / 10.13
        z = Unf_pvt_Zgas_d(t_K, p_MPa, GammaGas, correlation)
        Unf_pvt_Bg_m3m3 = Unf_pvt_Bg_z_m3m3(t_K, p_MPa, z)
    End Function

    Function Unf_pvt_Bg_z_m3m3(ByVal t_K As Double, ByVal p_MPa As Double, ByVal z As Double) As Double
        ' Расчет объемного коэффициента газа при известном коэффиенте сжимаемости газа
        ' rnt 20150303
        ' хорошо определить при какой температуре рассчитан объемный коэффициент газа

        Unf_pvt_Bg_z_m3m3 = 0.00034722 * t_K * z / p_MPa
    End Function

    '======================= сервисные функции для расчета свойств газа ==============================

    Private Function ZFactorEstimateDranchuk(ByVal Tpr As Double, ByVal Ppr As Double, ByVal z As Double) As Double
        'Continious function which return 0 if Z factor is correct for given pseudoreduced temperature and pressure
        'Used to find Z factor

        Const A1 As Double = 0.3265
        Const A2 As Double = -1.07
        Const A3 As Double = -0.5339
        Const a4 As Double = 0.01569
        Const a5 As Double = -0.05165
        Const a6 As Double = 0.5475
        Const a7 As Double = -0.7361
        Const a8 As Double = 0.1844
        Const a9 As Double = 0.1056
        Const a10 As Double = 0.6134
        Const a11 As Double = 0.721
        Dim rho_r As Double
        rho_r = 0.27 * Ppr / (z * Tpr)
        ZFactorEstimateDranchuk = -z + (A1 + A2 / Tpr + A3 / Tpr ^ 3 + a4 / Tpr ^ 4 + a5 / Tpr ^ 5) * rho_r +
       (a6 + a7 / Tpr + a8 / Tpr ^ 2) * rho_r ^ 2 - a9 * (a7 / Tpr + a8 / Tpr ^ 2) * rho_r ^ 5 +
       a10 * (1 + a11 * rho_r ^ 2) * rho_r ^ 2 / Tpr ^ 3 * Math.Exp(-a11 * rho_r ^ 2) + 1.0#
    End Function

    Private Function ZFactorDranchuk(ByVal Tpr As Double, ByVal Ppr As Double,
                            Optional ByRef msg As String = "") As Double
        ' rnt_bug 2015/03/03  расчет может быть не корректным при определенных значения приведенных давелния и температуры
        ' необходимо исправить - заменить метод деления отрезка пополам на метод ньютона модифицированный
        ' необходимо вставить сообщение о выходе расчета за границы применимости и предупреждение


        ' Debug.Assert Ppr >= const_Ppr_min
        Dim y_low As Double
        Dim y_hi As Double
        Dim Z_low As Double
        Dim Z_hi As Double
        Dim Z_mid As Double
        Z_low = 0.1
        Z_hi = 5
        Dim i As Integer
        i = 0
        'find foot of ZFactorEstimateDranchuk function by bisection of [Z_low...Z_hi] interval
        Do
            'we assume that for Z_low and Z_hi ZFactorEstimateDranchuk function of different signes
            Z_mid = 0.5 * (Z_hi + Z_low)
            y_low = ZFactorEstimateDranchuk(Tpr, Ppr, Z_low)
            y_hi = ZFactorEstimateDranchuk(Tpr, Ppr, Z_mid)
            If (y_low * y_hi < 0) Then
                Z_hi = Z_mid
            Else
                Z_low = Z_mid
            End If
            i += 1
        Loop Until (i > 200 Or Math.Abs(Z_low - Z_hi) < 0.000001)
        ' rnt check iteration convergence ******************************
        If i > 20 And Math.Abs(Z_low - Z_hi) > 0.001 Then
            ' solution not found
            msg = "*****ZFactorDranchuk: z фактор не найден из за расхождения итераций по Дранчуку. Tpr = " & Tpr & "  Ppr = " & Ppr
            ' error_probability = increment_error_probability(error_probability, 1000)
        End If
        '****************************************************************
        ZFactorDranchuk = Z_mid
        ' rnt 20150312 костыль для исправления ошибки при низких приведенных давлениях
        '              исправление дает небольшую погрешность - следует заменить далее на корректный метод решения уравнения дранчука
        If ZFactorDranchuk > 4.99 Then
            msg = " ZFactorDranchuk: корректировка z фактора из за расхождения итерация по Дранчуку. Tpr = " & Tpr & "  Ppr = " & Ppr
            ZFactorDranchuk = ZFactor(Tpr, Ppr)
        End If
        ' rnt 20150312 end
    End Function



    Private Function ZFactor2015_Kareem(ByVal Tpr As Double, ByVal Ppr As Double) As Double
        ' based on  https://link.springer.com/article/10.1007/s13202-015-0209-3
        '
        ' Kareem, L.A., Iwalewa, T.M. & Al-Marhoun, M.
        ' New explicit correlation for the compressibility factor of natural gas: linearized z-factor isotherms.
        ' J Petrol Explor Prod Technol 6, 481–492 (2016).
        ' https://doi.org/10.1007/s13202-015-0209-3

        Dim t As Double
        Dim AA As Double
        Dim BB As Double
        Dim CC As Double
        Dim DD As Double
        Dim EE As Double
        Dim FF As Double
        Dim gg As Double
        Dim A(19) As Double
        Dim y As Double
        Dim z As Double

        A(1) = 0.317842
        A(11) = -1.966847
        A(2) = 0.382216
        A(12) = 21.0581
        A(3) = -7.768354
        A(13) = -27.0246
        A(4) = 14.290531
        A(14) = 16.23
        A(5) = 0.000002
        A(15) = 207.783
        A(6) = -0.004693
        A(16) = -488.161
        A(7) = 0.096254
        A(17) = 176.29
        A(8) = 0.16672
        A(18) = 1.88453
        A(9) = 0.96691
        A(19) = 3.05921
        A(10) = 0.063069



        Dim t2 As Double
        Dim t3 As Double
        Dim y2 As Double

        t = 1 / Tpr
        t2 = t ^ 2
        t3 = t ^ 3
        AA = (A(1) * t * Math.Exp(A(2) * (1 - t) ^ 2) * Ppr) ^ 2
        BB = A(3) * t + A(4) * t2 + A(5) * t ^ 6 * Ppr ^ 6
        CC = A(9) + A(8) * t * Ppr + A(7) * t2 * Ppr ^ 2 + A(6) * t3 * Ppr ^ 3
        DD = A(10) * t * Math.Exp(A(11) * (1 - t) ^ 2)
        EE = A(12) * t + A(13) * t2 + A(14) * t3
        FF = A(15) * t + A(16) * t2 + A(17) * t3
        gg = A(18) + A(19) * t

        Dim DPpr As Double
        DPpr = DD * Ppr
        y = DPpr / ((1 + AA) / CC - AA * BB / (CC ^ 3))
        y2 = y ^ 2

        z = DPpr * (1 + y + y2 - y ^ 3) / (DPpr + EE * y2 - FF * y ^ gg) / ((1 - y) ^ 3)

        ZFactor2015_Kareem = z


    End Function

    Private Function ZFactor(ByVal Tpr As Double, ByVal Ppr As Double) As Double
        ' rnt_warning 20150303 известно, что функция дает большую погрешность при расчете
        ' не рекомендуется использовать

        Dim A As Double, b As Double, c As Double, d As Double

        A = 1.39 * (Tpr - 0.92) ^ 0.5 - 0.36 * Tpr - 0.101
        b = Ppr * (0.62 - 0.23 * Tpr) + Ppr ^ 2 * (0.006 / (Tpr - 0.86) - 0.037) + 0.32 * Ppr ^ 6 / Math.Exp(20.723 * (Tpr - 1))
        c = 0.132 - 0.32 * Math.Log(Tpr) / Math.Log(10)
        d = Math.Exp(0.715 - 1.128 * Tpr + 0.42 * Tpr ^ 2)

        ZFactor = A + (1 - A) * Math.Exp(-b) + c * Ppr ^ d


        '    ranges_good = ranges_good And CheckRanges(ZFactor, "ZFactor", const_Z_min, const_Z_max, _
        '                                                "Расчитанный z вне диапазона (используйте ZFactorDranchuk)", "ZFactor", True, error_probability)


    End Function

    Private Function PseudoTemperature(ByVal gamma_gas As Double) As Double
        PseudoTemperature = 95 + 171 * gamma_gas
    End Function

    Private Function PseudoPressure(ByVal gamma_gas As Double) As Double
        PseudoPressure = 4.9 - 0.4 * gamma_gas
    End Function

    Private Function PseudoTemperatureStanding(ByVal gamma_gas As Double) As Double
        PseudoTemperatureStanding = 93.3 + 180 * gamma_gas - 6.94 * gamma_gas ^ 2
    End Function

    Private Function PseudoPressureStanding(ByVal gamma_gas As Double) As Double
        PseudoPressureStanding = 4.6 + 0.1 * gamma_gas - 0.258 * gamma_gas ^ 2
    End Function

    ' ================================================================
    ' PVT water
    ' ================================================================

    ' Water viscosity
    Function Unf_pvt_viscosity_wat_cP(ByVal p_MPa As Double, ByVal Temperature_K As Double, ByVal Salinity_ppm As Double) As Double
        ' http://petrowiki.org/Produced_water_properties

        Dim wpTDS As Double, A As Double, b As Double, visc As Double, psi As Double


        wpTDS = Salinity_ppm / (10000)  ' weigth percent salinity

        A = 109.574 - 8.40564 * wpTDS + 0.313314 * wpTDS ^ 2 + 0.00872213 * wpTDS ^ 3
        b = -1.12166 + 0.0263951 * wpTDS - 0.000679461 * wpTDS ^ 2 - 5.47119 * 10 ^ (-5) * wpTDS ^ 3 + 1.55586 * 10 ^ (-6) * wpTDS ^ 4

        visc = A * (1.8 * Temperature_K - 460) ^ b
        psi = p_MPa * 145.04
        Unf_pvt_viscosity_wat_cP = visc * (0.9994 + 4.0295 * 10 ^ (-5) * psi + 3.1062 * 10 ^ (-9) * psi ^ 2)

    End Function

    ' Water density
    Function Unf_pvt_Bw_d(ByVal p_MPa As Double, ByVal Temperature_K As Double, ByVal Salinity_ppm As Double) As Double
        ' http://petrowiki.org/Produced_water_density


        Unf_pvt_Bw_d = Unf_pvt_BwSC_d(Salinity_ppm) / Unf_pvt_Bw_m3m3(p_MPa, Temperature_K)

    End Function

    ' Water FVF
    Function Unf_pvt_Bw_m3m3(ByVal p_MPa As Double, ByVal Temperature_K As Double) As Double
        ' http://petrowiki.org/Produced_water_formation_volume_factor


        Dim F As Double, psi As Double, dvwp As Double, dvwt As Double
        F = 1.8 * Temperature_K - 460
        psi = p_MPa * 145.04

        dvwp = -1.95301 * 10 ^ (-9) * psi * F - 1.72834 * 10 ^ (-13) * psi ^ 2 * F - 3.58922 * 10 ^ (-7) * psi - 2.25341 * 10 ^ (-10) * psi ^ 2
        dvwt = -1.0001 * 10 ^ (-2) + 1.33391 * 10 ^ (-4) * F + 5.50654 * 10 ^ (-7) * F ^ 2
        Unf_pvt_Bw_m3m3 = (1 + dvwp) * (1 + dvwt)

    End Function


    ' Water density at standard conditions
    Function Unf_pvt_BwSC_d(ByVal Salinity_ppm As Double) As Double

        Dim wpTDS As Double

        wpTDS = Salinity_ppm / (10000)
        Unf_pvt_BwSC_d = 0.0160185 * (62.368 + 0.438603 * wpTDS + 0.00160074 * wpTDS ^ 2)

    End Function


    Function Unf_pvt_Sal_BwSC_ppm(ByVal BwSC As Double) As Double
        ' функция для оценки солености воды по объемному коэффициенту (получена как обратная к Unf_pvt_BwSC_d)
        Unf_pvt_Sal_BwSC_ppm = ((624.711071129603 * BwSC / 0.0160185 - 20192.9595437054) ^ 0.5 - 137.000074965329) * 10000

    End Function

    ' GWR
    Function Unf_pvt_GWR_m3m3(ByVal p_MPa As Double, ByVal Temperature_K As Double, ByVal Salinity_ppm As Double) As Double
        ' rnt 20150303 надо найти источник корреляции - скорее всего крига Брилла по многофазному потоку
        ' 201503 не используется в расчетах

        Dim F As Double, psi As Double, wpTDS As Double, A As Double, b As Double, c As Double, Rswws As Double

        ' rnt 20150319 проверка диапазонов значений

        F = 1.8 * Temperature_K - 460
        psi = p_MPa * 145.04
        wpTDS = Salinity_ppm / (10000)

        A = 8.15839 - 0.0612265 * F + 1.91663 * 10 ^ (-4) * F ^ 2 - 2.1654 * 10 ^ (-7) * F ^ 3
        b = 1.01021 * 10 ^ (-2) - 7.44241 * 10 ^ (-5) * F + 3.05553 * 10 ^ (-7) * F ^ 2 - 2.94883 * 10 ^ (-10) * F ^ 3
        c = (-9.02505 + 0.130237 * F - 8.53425 * 10 ^ (-4) * F ^ 2 + 2.34122 * 10 ^ (-6) * F ^ 3 - 2.37049 * 10 ^ (-9) * F ^ 4) * 10 ^ (-7)
        Rswws = (A + b * psi + c * psi ^ 2) * 0.1781
        Unf_pvt_GWR_m3m3 = Rswws * 10 ^ (-0.0840655 * wpTDS * F ^ (-0.285854))

    End Function

    ' Water compressibility
    Function Unf_pvt_compressibility_wat_1atma(ByVal p_MPa As Double, ByVal Temperature_K As Double, ByVal Salinity_ppm As Double) As Double
        ' http://petrowiki.org/Produced_water_compressibility
        ' 201503 не используется в расчетах

        Dim F As Double, psi As Double

        F = 1.8 * Temperature_K - 460
        psi = p_MPa * 145.04

        Unf_pvt_compressibility_wat_1atma = 0.1 * 145.04 / (7.033 * psi + 0.5415 * Salinity_ppm - 537 * F + 403300)

    End Function


    Public Function Unf_pvt_gas_heat_capacity_ratio(gg As Double, t_K As Double) As Double
        'http://www.jmcampbell.com/tip-of-the-month/2013/05/variation-of-ideal-gas-heat-capacity-ratio-with-temperature-and-relative-density/
        ' eq 6
        ' temp range - 25C to 150 C
        ' gg range 0.55 to 2

        Dim k As Double
        Dim A As Double

        A = 0.000286
        k = (1.6 - 0.44 * gg + 0.097 * gg * gg) * (1 + 0.0385 * gg - A * t_K)
        Unf_pvt_gas_heat_capacity_ratio = k

    End Function

    ' функции для расчета границ образования гидратов
    ' Горидько К 2019

    'Y.F. Makogon, Hydrates of Natural Gas, PennWell, 1981, pp. 12–13.

    Public Function Unf_fa_hydrate_p_Makogon_atma(ByVal t_C As Double,
                                Optional ByVal Gas_gravity As Double = 0.7) As Double

        Dim b As Double
        Dim k As Double
        Dim A As Double
        Dim GL_P_Makogon_Hydrate_MPa As Double

        b = 2.681 - 3.811 * Gas_gravity + 1.679 * Gas_gravity ^ 2
        k = -0.006 + 0.11 * Gas_gravity + 0.011 * Gas_gravity
        A = b + 0.0497 * (t_C + k * t_C ^ 2) - 1
        GL_P_Makogon_Hydrate_MPa = Math.Exp(A)

        Unf_fa_hydrate_p_Makogon_atma = GL_P_Makogon_Hydrate_MPa * const_convert_MPa_atma

    End Function


    Public Function Unf_fa_hydrate_t_Moitee_C(ByVal p_atma As Double,
                             Optional ByVal Gas_gravity As Double = 0.7) As Double
        ' Motiee M. Estimate possibility of hydrate. Hydrological Processes. July 1991;70(7):98–99

        Dim C1 As Double
        Dim C2 As Double
        Dim C3 As Double
        Dim C4 As Double
        Dim C5 As Double
        Dim C6 As Double

        Dim p_psi As Double
        Dim GL_T_Moitee_Hydrate_F As Double

        p_psi = p_atma * const_convert_atma_psi

        C1 = -238.24469
        C2 = 78.99667
        C3 = -5.352544
        C4 = 349.473877
        C5 = -150.854675
        C6 = -27.604065
        GL_T_Moitee_Hydrate_F = C1 + C2 * Math.Log(p_psi) / Math.Log(10) + C3 * (Math.Log(p_psi) / Math.Log(10)) ^ 2 + C4 * Gas_gravity + C5 * Gas_gravity ^ 2 + C6 * Gas_gravity * Math.Log(p_psi) / Math.Log(10)

        Unf_fa_hydrate_t_Moitee_C = UC_Temperature_F_to_C(GL_T_Moitee_Hydrate_F)

    End Function


    Public Function Unf_fa_hydrate_t_Sun_C(ByVal p_atma As Double,
                             Optional ByVal Gas_gravity As Double = 0.7) As Double
        ' Sun C.-Y, Chen G.-J, Lin W, Guo T.-M. Hydrate formation conditions of sour
        ' natural gases. Journal of Chemical and Engineering Data. 2003;48:600–603.

        Dim C1 As Double
        Dim C2 As Double
        Dim C3 As Double
        Dim C4 As Double

        Dim p_MPa As Double
        Dim t1K As Double


        p_MPa = p_atma * const_convert_atma_MPa

        C1 = 4.343295
        C2 = 0.0010734
        C3 = -0.091984
        C4 = -1.071989

        t1K = 1000 / (C1 + C2 * p_MPa + C3 * Math.Log(p_MPa) + C4 * Gas_gravity)

        Unf_fa_hydrate_t_Sun_C = t1K - const_t_K_zero_C

    End Function


    'поправка Hammerschmidt на температуру гидратообразования при наличии ингибитора'
    Public Function Unf_fa_hydrate_dt_Hammerschmidt_C(ByVal C_W_perc As Double,
                                      Optional ByVal Ingibitor As Integer = 1) As Double

        'C_W_perc - weight rercent of ingibitor
        'type of ingibitor: 1 - methanol, 2 - ethylene glycol, 3 - triethylene glycol
        'k - constant for specific ingibitor
        'M - molecular weight

        Dim M As Double
        Dim k As Double

        If Ingibitor = 1 Then
            M = 32.04
            k = 2335
        ElseIf Ingibitor = 2 Then
            M = 62.07
            k = 2700
        ElseIf Ingibitor = 3 Then
            M = 150.17
            k = 5400
        Else
            Unf_fa_hydrate_dt_Hammerschmidt_C = 0
            Exit Function
        End If

        Unf_fa_hydrate_dt_Hammerschmidt_C = k * C_W_perc / (100 * M - C_W_perc * M)

    End Function


    'поправка carrol на температуру гидратообразования при наличии ингибитора'
    Public Function Unf_fa_hydrate_dt_Carroll_C(ByVal C_W_perc As Double,
                                      Optional ByVal Ingibitor As Integer = 1) As Double
        'Natural Gas Hydrates: A Guide for Engineers John Carroll



        'C_W_perc - weight rercent of ingibitor
        'type of ingibitor: 1 - methanol, 2 - ethylene glycol, 3 - triethylene glycol
        'k - constant for specific ingibitor
        'M - molecular weight

        Dim M As Double

        Dim A As Double
        Dim xl As Double



        If Ingibitor = 1 Then
            M = 32.04
            A = 0.21
        ElseIf Ingibitor = 2 Then
            M = 62.07
            A = -1.25
        ElseIf Ingibitor = 3 Then
            M = 150.17
            A = -15
        Else
            Unf_fa_hydrate_dt_Carroll_C = 0
            Exit Function
        End If
        Dim M_H20_gmol As Double
        M_H20_gmol = 18
        Dim wi As Double

        wi = C_W_perc / 100
        xl = wi / M / (wi / M + (1 - wi) / M_H20_gmol)

        Unf_fa_hydrate_dt_Carroll_C = -72 * (A * xl ^ 2 + Math.Log(1 - xl))

    End Function



End Module

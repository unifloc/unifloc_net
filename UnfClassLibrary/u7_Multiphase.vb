﻿'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
' Расчеты многофазного поток
' Расчет корреляции Ансари
Imports System.Math
Public Module u7_Multiphase
    Const c_p = 0.000009871668    ' переводной коэффициент
    Function sind(ang) As Double
        sind = Sin(CDbl(ang) / 180 * const_Pi)
    End Function
    Function cosd(ang) As Double
        cosd = Cos(CDbl(ang) / 180 * const_Pi)
    End Function
    ' временно ( не нашёл в alglib)
    Public Function MaxReal(ByVal M1 As Double, ByVal M2 As Double) As Double
        If M1 > M2 Then
            MaxReal = M1
        Else
            MaxReal = M2
        End If
    End Function

    Public Function MinReal(ByVal M1 As Double, ByVal M2 As Double) As Double
        If M1 < M2 Then
            MinReal = M1
        Else
            MinReal = M2
        End If
    End Function
    Public Function unf_GasGradient(ByVal arr_d_m As Double,
                                  ByVal arr_theta_deg As Double, ByVal eps_m As Double,
                                  ByVal Qg_rc_m3day As Double,
                                  ByVal Mug_rc_cP As Double,
                                  ByVal rho_grc_kgm3 As Double,
                                  ByVal p_atma As Double,
                                  Optional c_calibr_grav As Double = 1,
                                  Optional c_calibr_fric As Double = 1)
        ' gas gradient - for gaslift and annulus calculations

        ' for start - simplest estimation without friction and other
        Dim dPdLg_out_atmm As Double
        Dim dPdLf_out_atmm As Double
        Dim dPdLa_out_atmm As Double
        Dim Vsl_msec As Double
        Dim Vsg_msec As Double
        Dim Ap_m2 As Double
        Dim Hl_out_fr As Double
        Dim fpat_out_num As Double

        dPdLg_out_atmm = rho_grc_kgm3 * const_g * const_convert_Pa_atma * Sin(arr_theta_deg * const_Pi / 180)
        dPdLf_out_atmm = 0
        dPdLa_out_atmm = 0
        Vsl_msec = 0
        Ap_m2 = const_Pi * arr_d_m ^ 2 / 4
        Hl_out_fr = 0
        Vsl_msec = 0
        Vsg_msec = const_conver_sec_day * Qg_rc_m3day / Ap_m2
        fpat_out_num = 101 ' " gas" = gas

        unf_GasGradient = {dPdLg_out_atmm * c_calibr_grav + dPdLf_out_atmm * c_calibr_fric,
                            dPdLg_out_atmm * c_calibr_grav,
                            dPdLf_out_atmm * c_calibr_fric,
                            dPdLa_out_atmm,
                            Vsl_msec,
                            Vsg_msec,
                            Hl_out_fr,
                            fpat_out_num}
    End Function

    ' Федоров, Халиков (2016)
    Public Function unf_Saharov_Mokhov_Gradient(ByVal d As Double, ByVal THETA As Double, ByVal eps As Double, ByVal p As Double,
                                  ByVal q_osc As Double, ByVal q_wsc As Double, ByVal q_gsc As Double,
                                  ByVal b_o As Double, ByVal b_w As Double, ByVal b_g As Double, ByVal r_s As Double, ByVal mu_o As Double,
                                  ByVal mu_w As Double, ByVal mu_g As Double, ByVal sigma_o As Double, ByVal sigma_w As Double,
                                  ByVal rho_osc As Double, ByVal rho_wsc As Double, ByVal rho_gsc As Double,
                                  Optional Units As Integer = 1,
                                  Optional Payne_et_all_friction As Integer = 1, Optional correl3 As Double = 0,
                                  Optional c_calibr_grav As Double = 1,
                                  Optional c_calibr_fric As Double = 1)
        'function for calculation of pressure gradient in pipe according to Begs and Brill method
        'Return (psi/ft (atma/m))
        'Arguments
        'd - pipe internal diameter (ft (m))
        'theta - pipe inclination angel (degrees)
        'eps - pipe wall roughness (ft (m))
        'p - reference pressure (psi (atma))
        'q_oSC - oil rate at standard conditions (Stb/day (m3/day))
        'q_wSC - water rate at standard conditions (Stb/day (m3/day))
        'q_gSC - gas rate at standard conditions (scf/day (m3/day))
        'b_o - oil formation volume factor at reference pressure (bbl/stb (m3/sm3))
        'b_w - water formation volume factor at reference pressure (bbl/stb (m3/sm3))
        'b_g - gas formation volume factorat reference pressure (ft3/scf (m3/sm3))
        'rs - gas-oil solution ratio at reference pressure (Scf/stb (sm3/sm3))
        'mu_o - oil viscosity at reference pressure (cp)
        'mu_w - water viscosity at reference pressure (cp)
        'mu_g - gAs viscosity at reference pressure (cp)
        'sigma_o - oil-gAs surface tension coefficient (dynes/sm (Newton/m))
        'sigma_w - water-gAs surface tension coefficient (dynes/sm (Newton/m))
        'rho_oSC - oil density at standard conditions (lbm/ft3 (kg/m3))
        'rho_wSC - water density at standard conditions (lbm/ft3 (kg/m3))
        'rho_gSC - gas density at standard conditions(lbm/ft3 (kg/m3))
        'units - input/output units (0-field, 1 - metric)
        'Payne_et_all_holdup - flag indicationg weather to applied Payne et all correction and holdup (0 - not applied, 1 - applied)
        'Payne_et_all_friction - flag indicationg weather to apply Payne et all correction for friction (0 - not applied, 1 - applied)
        'dpdl_g - used to otput pressure gradient due to gravity (psi/ft (atma/m))
        'dpdl_f - used to output pressure gradient due to friction (psi/ft (atma/m))
        'v_sl - used to output liquid superficial velocity (ft/sec (m/sec))
        'v_sg - used to output gas superficial velocity (ft/sec (m/sec))
        'h_l - used to output liquid holdup
        Dim dPdLg_out_atmm As Double
        Dim dPdLf_out_atmm As Double
        Dim Vsl_out_msec As Double
        Dim Vsg_out_msec As Double
        Dim Hl_out_fr As Double
        Dim fpat_out_num
        Dim dPdLa_out_atmm As Double
        Dim V_sl As Double, V_sg As Double
        Dim e As Double
        Dim dPdL_out_atmm
        Dim dpdl_g As Double, dpdl_f As Double, dpdl_g1 As Double, dpdl_f1 As Double
        Dim h_l As Double
        'Conversion factors and constants (field / metric)
        'acceleration due to gravity
        Dim g(2) As Double
        g(0) = 32.174 : g(1) = 9.8
        'Oil, water velocity conversion
        Dim c_q(2) As Double
        c_q(0) = 5.6146 : c_q(1) = 1
        'Gas-oil solution ratio conversion factor
        Dim c_rs(2) As Double
        c_rs(0) = 0.17811 : c_rs(1) = 1
        'Reinolds number conversion factor
        Dim c_re(2) As Double
        c_re(0) = 1.488 : c_re(1) = 1
        'Pressure gradient conversion factor
        Dim c_p(2) As Double
        c_p(0) = 0.00021583 : c_p(1) = 0.000009871668
        'liquid velocity number conversion
        Dim c_sl(2) As Double
        c_sl(0) = 4.61561 : c_sl(1) = 1
        'Calculate auxilary values
        'Pipe cross-sectional area
        Dim a_p As Double
        a_p = const_Pi * d ^ 2 / 4
        'Calculate flow rates at reference pressure
        Dim q_o, q_w, q_l, q_g As Double
        q_o = c_q(Units) * q_osc * b_o
        q_w = c_q(Units) * q_wsc * b_w
        q_l = q_o + q_w
        q_g = b_g * (q_gsc - r_s * q_osc)
        'if gas rate is negative - assign gas rate to zero
        If q_g < 0 Then
            q_g = 0
        End If


        Dim f_w, lambda_l As Double
        Dim rho_o, rho_w, rho_l, rho_g, rho_n As Double

        Dim f_mohov As Double  'коэффициентобщих потерь
        Dim v_m As Double

        If q_l > 0 Then
            'calculate volume fraction of water in liquid at no-slip conditions
            f_w = q_w / q_l
            'volume fraction of liquid at no-slip conditions
            lambda_l = q_l / (q_l + q_g)
            'densities
            rho_o = (rho_osc + c_rs(Units) * r_s * rho_gsc) / b_o
            rho_w = rho_wsc / b_w
            rho_l = rho_o * (1 - f_w) + rho_w * f_w
            rho_g = rho_gsc / b_g
            'no-slip mixture density
            rho_n = rho_l * lambda_l + rho_g * (1 - lambda_l)
            'Liquid surface tension
            Dim sigma_l As Double
            sigma_l = sigma_o * (1 - f_w) + sigma_w * f_w
            'Liquid viscosity
            Dim mu_l As Double
            mu_l = mu_o * (1 - f_w) + mu_w * f_w
            'No slip mixture viscosity
            Dim mu_n As Double
            mu_n = mu_l * lambda_l + mu_g * (1 - lambda_l)
            'Sureficial velocities
            V_sl = 0.000011574 * q_l / a_p
            V_sg = 0.000011574 * q_g / a_p
            v_m = V_sl + V_sg
            'Reinolds number
            Dim n_re As Double
            n_re = c_re(Units) * 1000 * rho_n * v_m * d / mu_n
            'Froude number
            Dim n_fr As Double
            n_fr = v_m ^ 2 / (g(Units) * d)
            'Liquid velocity number
            Dim n_lv As Double
            n_lv = c_sl(Units) * V_sl * (rho_l / (g(Units) * sigma_l)) ^ 0.25
            'Pipe relative roughness
            e = eps / d
            '-----------------------------------------------------------------------
            'determine flow pattern
            Dim delta_rho As Double   ' увеличение плотности смеси за счёт относительного движения газа
            Dim We As Double ' Число Вебера
            Dim Ku As Double 'Безразмерный параметр, покритерию подобный критерию Кутателадзе
            delta_rho = rho_o - rho_g
            We = sigma_o / (delta_rho * d ^ 2 * g(Units))
            Ku = ((rho_l ^ 2) / (delta_rho ^ 2) * (n_fr ^ 2) / We) ^ (1 / 4)
            f_mohov = (0.13 * Ku + 1) / (1.13 * Ku + 1) * delta_rho / rho_l * 2 * (1 - lambda_l) / n_fr * Sin(const_Pi / 180 * THETA) + 0.11 * (68 / n_re + e / d) ^ 0.25
        Else
            f_mohov = 0
            v_m = 0
            rho_n = rho_wsc / b_w
        End If
        'calculate pressure gradient due to gravity
        dpdl_g = c_p(Units) * rho_n * g(Units) * Sin(const_Pi / 180 * THETA)
        'calculate pressure gradient due to friction
        dpdl_f = c_p(Units) * f_mohov * rho_n * v_m ^ 2 / (2 * d)
        'calculate pressure gradient
        dPdL_out_atmm = dpdl_g + dpdl_f
        dPdLg_out_atmm = dpdl_g
        dPdLf_out_atmm = dpdl_f
        dPdLa_out_atmm = 0
        Vsl_out_msec = V_sl
        Vsg_out_msec = V_sg
        Hl_out_fr = lambda_l
        fpat_out_num = 0
        unf_Saharov_Mokhov_Gradient = {dPdLg_out_atmm * c_calibr_grav + dPdLf_out_atmm * c_calibr_fric,
                            dPdLg_out_atmm * c_calibr_grav,
                            dPdLf_out_atmm * c_calibr_fric,
                            dPdLa_out_atmm,
                            Vsl_out_msec,
                            Vsg_out_msec,
                            Hl_out_fr,
                            fpat_out_num}
    End Function
    ' Расчет естественной сепарации по Маркезу
    Public Function unf_natural_separation(ByVal d_tub_m_ As Double, ByVal d_cas_m_ As Double,
                                  ByVal qliq_sm3day As Double, ByVal Qg_sc_m3day As Double,
                                  ByVal bo_m3m3 As Double, ByVal bg_m3m3 As Double, ByVal sigma_o As Double, ByVal sigma_w As Double,
                                  ByVal rho_osc As Double, ByVal rho_gsc As Double, ByVal WCT As Double, Optional Units As Integer = 1) As Double
        'function calculates pressure gradient for Zero Net Liquid flow in annulus
        'Return (psi/ft (atma/m))
        'Arguments
        'd_tub_m_ -  internal diameter arr_theta_deg( (m))
        'd_cas_m_ -  internal diameter arr_theta_deg( (m))
        'arr_theta_deg - pipe inclination angel (degrees)
        'p - reference pressure ( (atma))
        'q_oil_m3day - liquid rate at standard conditions ( (m3/day))
        'Qg_sc_m3day - gas rate at standard conditions ((m3/day))
        'Bo_m3m3 - oil formation volume factor at reference pressure ( (m3/sm3))
        'Bg_m3m3 - gas formation volume factor at reference pressure ( (m3/sm3))
        'Rs_m3m3 - gas-oil solution ratio at reference pressure ( (sm3/sm3))
        'sigma_o - oil-gAs surface tension coefficient ((Newton/m))
        'rho_osc - oil density at standard conditions ( (kg/m3))
        'rho_gsc - gas density at standard conditions((kg/m3))
        Dim a_p As Double
        Dim q_g As Double
        Dim q_l As Double
        Dim lambda_l As Double
        Dim rho_o, rho_w, rho_l, rho_g, rho_n As Double
        Dim V_sg As Double
        Dim V_sl As Double
        Dim sigma_l As Double
        Dim v_m As Double
        Dim n_fr As Double
        Dim flow_pattern As Integer
        Dim v_inf As Double
        Dim a As Double, b As Double, c As Double, d As Double, st As Double, backst As Double, M As Double


        If (qliq_sm3day = 0) Or (d_tub_m_ = d_cas_m_) Then
            unf_natural_separation = 1
            Exit Function
        End If

        'Calculate pressure gradient
        'Annulus cross-sectional area
        a_p = const_Pi * (d_cas_m_ ^ 2 - d_tub_m_ ^ 2) / 4
        q_g = bg_m3m3 * Qg_sc_m3day
        q_l = bo_m3m3 * qliq_sm3day * (1 - WCT / 100) + qliq_sm3day * WCT / 100
        'calculate oil density
        'volume fraction of liquid at no-slip conditions
        lambda_l = q_l / (q_l + q_g)
        'densities
        rho_o = (rho_osc) / bo_m3m3
        rho_w = 1000
        'TODO - replace water density
        rho_l = rho_o * (1 - WCT / 100) + rho_w * WCT / 100
        rho_g = rho_gsc / bg_m3m3
        'no-slip mixture density
        rho_n = rho_l * lambda_l + rho_g * (1 - lambda_l)
        'Gas sureficial velocity
        V_sg = 0.000011574 * q_g / a_p
        'Liquid sureficial velocity
        V_sl = 0.000011574 * q_l / a_p
        '----------------------
        'Liquid surface tension
        sigma_l = sigma_o * (1 - WCT / 100) + sigma_w * WCT / 100
        'Sureficial velocities
        v_m = V_sl + V_sg
        'Froude number
        n_fr = v_m ^ 2 / (const_g * (d_cas_m_ - d_tub_m_))
        '-----------------------------------------------------------------------
        'determine flow pattern
        If (n_fr >= 316 * lambda_l ^ 0.302 Or n_fr >= 0.5 * lambda_l ^ -6.738) Then
            flow_pattern = 2
        Else
            If (n_fr <= 0.000925 * lambda_l ^ -2.468) Then
                flow_pattern = 0
            Else
                If (n_fr <= 0.1 * lambda_l ^ -1.452) Then
                    flow_pattern = 3
                Else
                    flow_pattern = 1
                End If
            End If
        End If
        '-----------------------
        'Calculate terminal gas rise velosity
        If (flow_pattern = 0 Or flow_pattern = 1) Then
            v_inf = 1.53 * (const_g * sigma_l * (rho_l - rho_g) / rho_l ^ 2) ^ 0.25
        Else
            v_inf = 1.41 * (const_g * sigma_l * (rho_l - rho_g) / rho_l ^ 2) ^ 0.25
        End If
        ' calculate separation efficienty
        a = -0.0093
        b = 57.758
        c = 34.4
        d = 1.308
        st = 272
        backst = 1 / 272
        M = V_sl / v_inf
        If M > 13 Then
            unf_natural_separation = 0
            Exit Function
        End If
        unf_natural_separation = ((1 + (a * b + c * M ^ d) / (b + M ^ d)) ^ st + M ^ st) ^ backst - M
    End Function

    Public Function unf_BegsBrillGradient(ByVal arr_d_m As Double,
                                  ByVal arr_theta_deg As Double, ByVal eps_m As Double,
                                  ByVal Ql_rc_m3day As Double, ByVal Qg_rc_m3day As Double,
                                  ByVal Mul_rc_cP As Double, ByVal Mug_rc_cP As Double,
                                  ByVal sigma_l_Nm As Double,
                                  ByVal rho_lrc_kgm3 As Double,
                                  ByVal rho_grc_kgm3 As Double,
                                  Optional Payne_et_all_holdup As Integer = 0,
                                  Optional Payne_et_all_friction As Integer = 1,
                                  Optional c_calibr_grav As Double = 1,
                                  Optional c_calibr_fric As Double = 1)
        'function for calculation of pressure gradient in pipe according to Begs and Brill method
        'Return (psi/ft (atma/m))
        'Arguments
        'd - pipe internal diameter ( (m))
        'arr_theta_deg - pipe inclination angel (degrees)
        'eps_m - pipe wall roughness ( (m))
        'p - reference pressure ( (atma))
        'q_oSC - oil rate at standard conditions ( (m3/day))
        'q_wSC - water rate at standard conditions ( (m3/day))
        'q_gSC - gas rate at standard conditions ((m3/day))
        'Bo_m3m3 - oil formation volume factor at reference pressure ( (m3/sm3))
        'Bw_m3m3 - water formation volume factor at reference pressure ( (m3/sm3))
        'Bg_m3m3 - gas formation volume factorat reference pressure ( (m3/sm3))
        'rs - gas-oil solution ratio at reference pressure ( (sm3/sm3))
        'mu_oil_cP - oil viscosity at reference pressure (cp)
        'mu_wat_cP - water viscosity at reference pressure (cp)
        'mu_gas_cP - gAs viscosity at reference pressure (cp)
        'sigma_oil_gas_Nm - oil-gAs surface tension coefficient ((Newton/m))
        'sigma_wat_gas_Nm - water-gAs surface tension coefficient ( (Newton/m))
        'rho_oSC - oil density at standard conditions ( (kg/m3))
        'rho_wSC - water density at standard conditions ( (kg/m3))
        'rho_gSC - gas density at standard conditions((kg/m3))
        '
        'Payne_et_all_holdup - flag indicationg weather to applied Payne et all correction and holdup (0 - not applied, 1 - applied)
        'Payne_et_all_friction - flag indicationg weather to apply Payne et all correction for friction (0 - not applied, 1 - applied)  obsolete
        'dpdl_g - used to otput pressure gradient due to gravity ( (atma/m))
        'dpdl_f - used to output pressure gradient due to friction ( (atma/m))
        'v_sl - used to output liquid superficial velocity ( (m/sec))
        'v_sg - used to output gas superficial velocity ( (m/sec))
        'h_l - used to output liquid holdup
        'Calculate auxilary values
        Dim roughness_d As Double

        Dim dPdLg_out_atmm As Double
        Dim dPdLf_out_atmm As Double
        Dim Hl_out_fr As Double
        Dim fpat_out_num
        Dim dPdLa_out_atmm As Double
        Dim Ap_m2 As Double ' площадь трубы
        Dim lambda_l As Double
        Dim Vsl_msec, Vsg_msec, Vsm_msec As Double
        Dim Rho_n_kgm3 As Double   ' no slip density
        Dim rho_s As Double        ' mix density
        Dim Mu_n_cP As Double
        Dim n_re As Double 'Reinolds number
        Dim n_fr As Double 'Froude number
        Dim n_lv As Double 'Liquid velocity number
        Dim flow_pattern As Integer
        Dim l_2, l_3, AA As Double
        Dim f_n As Double ' normalized friction factor
        Dim f As Double ' friction factor
        Dim y, S As Double
        Dim fy As Double
        Const c_p = 0.000009871668   ' переводной коэффициент

        Ap_m2 = const_Pi * arr_d_m ^ 2 / 4
        If Ql_rc_m3day = 0 Then
            ' специально отработае случай нулевого дебита
            lambda_l = 1
            Hl_out_fr = 1
            f = 0
            Rho_n_kgm3 = rho_lrc_kgm3 * lambda_l + rho_grc_kgm3 * (1 - lambda_l) ' No-slip mixture density
            flow_pattern = 0
        Else
            lambda_l = Ql_rc_m3day / (Ql_rc_m3day + Qg_rc_m3day)
            roughness_d = eps_m / arr_d_m
            Vsl_msec = const_conver_sec_day * Ql_rc_m3day / Ap_m2
            Vsg_msec = const_conver_sec_day * Qg_rc_m3day / Ap_m2
            Vsm_msec = Vsl_msec + Vsg_msec
            Rho_n_kgm3 = rho_lrc_kgm3 * lambda_l + rho_grc_kgm3 * (1 - lambda_l) ' No-slip mixture density
            Mu_n_cP = Mul_rc_cP * lambda_l + Mug_rc_cP * (1 - lambda_l) ' No slip mixture viscosity
            n_re = 1000 * Rho_n_kgm3 * Vsm_msec * arr_d_m / Mu_n_cP
            n_fr = Vsm_msec ^ 2 / (const_g * arr_d_m)
            n_lv = Vsl_msec * (rho_lrc_kgm3 / (const_g * sigma_l_Nm)) ^ 0.25
            '-----------------------------------------------------------------------
            'determine flow pattern
            If (n_fr >= 316 * lambda_l ^ 0.302 Or n_fr >= 0.5 * lambda_l ^ -6.738) Then
                flow_pattern = 2
            Else
                If (n_fr <= 0.000925 * lambda_l ^ -2.468) Then
                    flow_pattern = 0
                Else
                    If (n_fr <= 0.1 * lambda_l ^ -1.452) Then
                        flow_pattern = 3
                    Else
                        flow_pattern = 1
                    End If
                End If
            End If
            '-----------------------------------------------------------------------
            'determine liquid holdup
            If (flow_pattern = 0 Or flow_pattern = 1 Or flow_pattern = 2) Then
                Hl_out_fr = h_l_arr_theta_deg(flow_pattern, lambda_l, n_fr, n_lv, arr_theta_deg, Payne_et_all_holdup)
            Else
                l_2 = 0.000925 * lambda_l ^ -2.468
                l_3 = 0.1 * lambda_l ^ -1.452
                AA = (l_3 - n_fr) / (l_3 - l_2)
                Hl_out_fr = AA * h_l_arr_theta_deg(0, lambda_l, n_fr, n_lv, arr_theta_deg, Payne_et_all_holdup) +
                          (1 - AA) * h_l_arr_theta_deg(1, lambda_l, n_fr, n_lv, arr_theta_deg, Payne_et_all_holdup)
            End If
            'Calculate normalized friction factor
            f_n = unf_friction_factor(n_re, roughness_d)
            'calculate friction factor correction for multiphase flow
            y = MaxReal(lambda_l / Hl_out_fr ^ 2, 0.000001)
            fy = MaxReal(Log(y), 0.000001)
            If (y > 1 And y < 1.2) Then
                S = Log(2.2 * y - 1.2)
            Else
                S = fy / (-0.0523 + 3.182 * fy - 0.8725 * fy ^ 2 + 0.01853 * fy ^ 4)
            End If
            'calculate friction factor
            f = f_n * Exp(S)
        End If

        rho_s = rho_lrc_kgm3 * Hl_out_fr + rho_grc_kgm3 * (1 - Hl_out_fr) 'calculate mixture density
        dPdLg_out_atmm = c_p * rho_s * const_g * sind(arr_theta_deg) 'calculate pressure gradient due to gravity
        dPdLf_out_atmm = c_p * f * Rho_n_kgm3 * Vsm_msec ^ 2 / (2 * arr_d_m) 'calculate pressure gradient due to friction
        dPdLa_out_atmm = 0  'calculate pressure gradient ' not acounted in BeggsBrill
        fpat_out_num = flow_pattern

        unf_BegsBrillGradient = {dPdLg_out_atmm * c_calibr_grav + dPdLf_out_atmm * c_calibr_fric,
                                    dPdLg_out_atmm * c_calibr_grav,
                                    dPdLf_out_atmm * c_calibr_fric,
                                    dPdLa_out_atmm,
                                    Vsl_msec,
                                    Vsg_msec,
                                    Hl_out_fr,
                                    fpat_out_num}

    End Function
    Public Function unf_AnsariGradient(ByVal arr_d_m As Double,
                                  ByVal arr_theta_deg As Double, ByVal eps_m As Double,
                                  ByVal Ql_rc_m3day As Double, ByVal Qg_rc_m3day As Double,
                                  ByVal Mul_rc_cP As Double, ByVal Mug_rc_cP As Double,
                                  ByVal sigma_l_Nm As Double,
                                  ByVal rho_lrc_kgm3 As Double,
                                  ByVal rho_grc_kgm3 As Double,
                                  ByVal p_atma As Double,
                                  Optional c_calibr_grav As Double = 1,
                                  Optional c_calibr_fric As Double = 1)

        Dim dPdLg_out_atmm As Double
        Dim dPdLf_out_atmm As Double
        Dim Hl_out_fr As Double
        Dim fpat_out_num As String
        Dim fpat_out_ As Double
        Dim dPdLa_out_atmm As Double
        Dim dPdL_out_atmm As Double
        Dim pgf_out_atmm As Double
        Dim pge_out_atmm As Double
        Dim pga_out_atmm As Double
        Dim pgt_out_atmm As Double

        ' znlf - calculates zero net liquid flow - gas flow through liquid column
        '=================

        Dim roughness_d As Double
        Dim Ap_m2 As Double ' площадь трубы
        Dim lambda_l As Double
        Dim Vsl_msec As Double, Vsg_msec As Double
        'Dim flow_pattern As Integer
        'Dim iErr
        Dim ang1 As Double
        'Dim timeStamp

        'timeStamp = Time()

        Try
            roughness_d = eps_m / arr_d_m
            Ap_m2 = const_Pi * arr_d_m ^ 2 / 4
            If Ql_rc_m3day = 0 Then
                lambda_l = 1
            Else
                lambda_l = Ql_rc_m3day / (Ql_rc_m3day + Qg_rc_m3day)
            End If
            Vsl_msec = const_conver_sec_day * Ql_rc_m3day / Ap_m2
            Vsg_msec = const_conver_sec_day * Qg_rc_m3day / Ap_m2
            ang1 = arr_theta_deg
            If arr_theta_deg < 0 Then
                ' Ansari not working for downward flow
                AddLogMsg("AnsariGradient: arr_theta_deg = " & arr_theta_deg & " negative. Ansari not for downward flow. Angle inverted")
                arr_theta_deg = -arr_theta_deg
            End If
            If arr_theta_deg < 75 Then
                ang1 = 75
                '   addLogMsg "AnsariGradient: arr_theta_deg = " & arr_theta_deg & " less than 75 degrees. 75 used for calc"
            End If
            If arr_theta_deg > 90 Then
                ang1 = 90
                AddLogMsg("AnsariGradient: arr_theta_deg = " & arr_theta_deg & " greater than 90 degrees. 90 deg used for calc")
            End If
            Call Ansari(ang1, arr_d_m, roughness_d, p_atma, Vsl_msec, Vsg_msec, lambda_l,
                            rho_grc_kgm3, rho_lrc_kgm3, Mug_rc_cP, Mul_rc_cP,
                            sigma_l_Nm,
                            Hl_out_fr, pgf_out_atmm, pge_out_atmm, pga_out_atmm, pgt_out_atmm, fpat_out_num)

            dPdL_out_atmm = (pge_out_atmm * Sin(arr_theta_deg * const_Pi / 180) / Sin(ang1 * const_Pi / 180) + pgf_out_atmm + pga_out_atmm)
            dPdLg_out_atmm = pge_out_atmm * Sin(arr_theta_deg * const_Pi / 180) / Sin(ang1 * const_Pi / 180)
            dPdLf_out_atmm = pgf_out_atmm
            dPdLa_out_atmm = pga_out_atmm

            Select Case fpat_out_num
                Case " liq" : fpat_out_ = 100 ' " liq" = liquid
                Case " gas" : fpat_out_ = 101 ' " gas" = gas
                Case "anul" : fpat_out_ = 105 ' "anul" = annular
                Case "dbub" : fpat_out_ = 104 ' "dbub" = dispersed bubble
                Case "slug" : fpat_out_ = 103 ' "slug" = slug
                Case "bubl" : fpat_out_ = 102 ' "bubl" = bubbly
                Case "  na" : fpat_out_ = 199
            End Select '(fpat)


            unf_AnsariGradient = {dPdLg_out_atmm * c_calibr_grav + dPdLf_out_atmm * c_calibr_fric,
                                    dPdLg_out_atmm * c_calibr_grav,
                                    dPdLf_out_atmm * c_calibr_fric,
                                    dPdLa_out_atmm,
                                    Vsl_msec,
                                    Vsg_msec,
                                    Hl_out_fr,
                                    fpat_out_}
            Exit Function
        Catch ex As Exception
            unf_AnsariGradient = "error"
        End Try

    End Function

    Public Function unf_UnifiedTUFFPGradient(ByVal arr_d_m As Double,
                                  ByVal arr_theta_deg As Double, ByVal eps_m As Double,
                                  ByVal Ql_rc_m3day As Double, ByVal Qg_rc_m3day As Double,
                                  ByVal Mul_rc_cP As Double, ByVal Mug_rc_cP As Double,
                                  ByVal sigma_l_Nm As Double,
                                  ByVal rho_lrc_kgm3 As Double,
                                  ByVal rho_grc_kgm3 As Double,
                                  ByVal p_atma As Double,
                                  Optional c_calibr_grav As Double = 1,
                                  Optional c_calibr_fric As Double = 1)

        Dim dPdLg_out_atmm As Double
        Dim dPdLf_out_atmm As Double
        Dim Hl_out_fr As Double
        Dim fpat_out_num As Integer
        Dim dPdLa_out_atmm As Double
        Dim dPdL_out_atmm As Double
        Dim roughness_d As Double
        Dim Ap_m2 As Double ' площадь трубы
        ' Dim lambda_l   As Double
        Dim Vsl_msec As Double, Vsg_msec As Double
        Dim pgf_out As Double
        Dim pge_out As Double
        Dim pga_out As Double
        Dim pgt_out As Double
        Dim iErr As Double
        Dim vf#, hlf#, SL#, FF#, hls#, cu#, fqn#, rsu#, icon#, cs#, cf#, VC# 

        roughness_d = eps_m / arr_d_m
        Ap_m2 = const_Pi * arr_d_m ^ 2 / 4
        ' If Ql_rc_m3day > 0 Then
        '     lambda_l = Ql_rc_m3day / (Ql_rc_m3day + Qg_rc_m3day)
        ' Else
        '    lambda_l = 1
        ' End If
        Vsl_msec = const_conver_sec_day * Ql_rc_m3day / Ap_m2
        Vsg_msec = const_conver_sec_day * Qg_rc_m3day / Ap_m2
        'On Error Resume Next
        Dim flpat As String
        Try
            ' Debug.Print "start" + " arr_d_m  =" + CStr(arr_d_m) + " roughness_d  =" + CStr(roughness_d) + " arr_theta_deg  =" + CStr(arr_theta_deg) _
            '                     + " Vsl_msec  =" + CStr(Vsl_msec) + " Vsg_msec  =" + CStr(Vsg_msec) + " rho_lrc_kgm3  =" + CStr(rho_lrc_kgm3) + " rho_grc_kgm3  =" + CStr(rho_grc_kgm3)
            ' Debug.Print " Mul_rc_cP  =" + CStr(Mul_rc_cP) _
            '                    + " Mug_rc_cP  =" + CStr(Mug_rc_cP) + "  sigma_l_Nm =" + CStr(sigma_l_Nm) + " p_atma = " + CStr(p_atma)
            Call zhangmodel(arr_d_m, roughness_d, arr_theta_deg, Vsl_msec, Vsg_msec, rho_lrc_kgm3, rho_grc_kgm3, Mul_rc_cP, Mug_rc_cP, sigma_l_Nm, p_atma,
                        Hl_out_fr, pgt_out, pgf_out, flpat,
                        vf, hlf, SL, FF, hls, cu, fqn, rsu, icon, cs, cf, VC, pge_out, pga_out)
            ' Debug.Print "done -------" + " pgt_out  =" + CStr(-pgt_out)
        Catch ex As Exception
            dPdL_out_atmm = -pgt_out
            dPdLg_out_atmm = -pge_out
            dPdLf_out_atmm = -pgf_out
            dPdLa_out_atmm = -pga_out

            Select Case flpat
                Case "liq" : fpat_out_num = 200  '                 " liq" = liquid
                Case "gas" : fpat_out_num = 201  '                 " gas" = gas
                Case "ann" : fpat_out_num = 207  '                 "anul" = annular
                Case "d-b" : fpat_out_num = 206 '                 "dbub" = dispersed bubble
                Case "slug" : fpat_out_num = 205 '                 "slug" = slug
                Case "bub" : fpat_out_num = 202  '                 "bubl" = bubbly
                Case "int" : fpat_out_num = 203
                Case "str" : fpat_out_num = 204
                Case "n-a" : fpat_out_num = 299
            End Select '(fpat)

            'Debug.Assert Abs(dPdL_out_atmm - (dPdLg_out_atmm + dPdLf_out_atmm)) < 0.1
            ' check of gradient calculation output. general output from correlation calc must be equal to its patrs
            ' otherwise - error in correlation output

            unf_UnifiedTUFFPGradient = {dPdLg_out_atmm * c_calibr_grav + dPdLf_out_atmm * c_calibr_fric,
                                        dPdLg_out_atmm * c_calibr_grav,
                                        dPdLf_out_atmm * c_calibr_fric,
                                        dPdLa_out_atmm,
                                        Vsl_msec,
                                        Vsg_msec,
                                        Hl_out_fr,
                                        fpat_out_num}
        End Try

    End Function

    ' Данный модуль содержит функцию расчета градиента давления по методике Грея
    '
    ' Федоров, Халиков (2016)
    Public Function unf_GrayModifiedGradient(ByVal d_m As Double,
                                      ByVal theta_deg As Double, ByVal eps_m As Double,
                                      ByVal Ql_rc_m3day As Double, ByVal Qg_rc_m3day As Double,
                                      ByVal Mul_rc_cP As Double, ByVal Mug_rc_cP As Double,
                                      ByVal sigma_l_Nm As Double,
                                      ByVal rho_lrc_kgm3 As Double,
                                      ByVal rho_grc_kgm3 As Double,
                                      Optional Payne_et_all_holdup As Integer = 0,
                                      Optional Payne_et_all_friction As Integer = 1,
                                      Optional ByVal correl3 As Double = 0,
                                      Optional c_calibr_grav As Double = 1,
                                      Optional c_calibr_fric As Double = 1)

        'function for calculation of pressure gradient in pipe according to Begs and Brill method
        'Return (psi/ft (atma/m))
        'Arguments
        'd - pipe internal diameter (ft (m))
        'theta - pipe inclination angel (degrees)
        'eps - pipe wall roughness (ft (m))
        'p - reference pressure (psi (atma))
        'q_oSC - oil rate at standard conditions (Stb/day (m3/day))
        'q_wSC - water rate at standard conditions (Stb/day (m3/day))
        'q_gSC - gas rate at standard conditions (scf/day (m3/day))
        'b_o - oil formation volume factor at reference pressure (bbl/stb (m3/sm3))
        'b_w - water formation volume factor at reference pressure (bbl/stb (m3/sm3))
        'b_g - gas formation volume factorat reference pressure (ft3/scf (m3/sm3))
        'rs - gas-oil solution ratio at reference pressure (Scf/stb (sm3/sm3))
        'mu_o - oil viscosity at reference pressure (cp)
        'mu_w - water viscosity at reference pressure (cp)
        'mu_g - gAs viscosity at reference pressure (cp)
        'sigma_o - oil-gAs surface tension coefficient (dynes/sm (Newton/m))
        'sigma_w - water-gAs surface tension coefficient (dynes/sm (Newton/m))
        'rho_oSC - oil density at standard conditions (lbm/ft3 (kg/m3))
        'rho_wSC - water density at standard conditions (lbm/ft3 (kg/m3))
        'rho_gSC - gas density at standard conditions(lbm/ft3 (kg/m3))
        'units - input/output units (0-field, 1 - metric)
        'Payne_et_all_holdup - flag indicationg weather to applied Payne et all correction and holdup (0 - not applied, 1 - applied)
        'Payne_et_all_friction - flag indicationg weather to apply Payne et all correction for friction (0 - not applied, 1 - applied)
        'dpdl_g - used to otput pressure gradient due to gravity (psi/ft (atma/m))
        'dpdl_f - used to output pressure gradient due to friction (psi/ft (atma/m))
        'v_sl - used to output liquid superficial velocity (ft/sec (m/sec))
        'v_sg - used to output gas superficial velocity (ft/sec (m/sec))
        'h_l - used to output liquid holdup


        Dim roughness_d As Double

        Dim dPdLg_out_atmm As Double
        Dim dPdLf_out_atmm As Double
        Dim Hl_out_fr As Double
        Dim fpat_out_num As Double
        Dim dPdLa_out_atmm As Double
        Dim Ap_m2 As Double ' площадь трубы
        Dim lambda_l As Double
        Dim Vsl_msec, Vsg_msec, Vsm_msec As Double
        Dim Rho_n_kgm3 As Double   ' no slip density
        Dim rho_s As Double        ' mix density
        Dim Mu_n_cP As Double
        Dim n_re As Double 'Reinolds number
        Dim n_fr As Double 'Froude number
        Dim n_lv As Double 'Liquid velocity number
        ' Dim flow_pattern As Integer
        Dim f_n As Double ' normalized friction factor
        Dim f As Double ' friction factor
        Dim y, S As Double
        Dim r As Double 'dimensionless'superficial liquid to gas ratio parameter
        Dim e As Double    'Pipe relative roughness
        Dim E1 As Double
        Dim crit As Double
        Dim Nv As Double ' dimensionless 'velocity number
        Dim Nd As Double 'nominal diameter
        Dim b As Double
        Dim h_l As Double 'liquid holdup by original Gray
        Dim dpdl_g As Double, dpdl_f As Double, dpdl_g1 As Double, dpdl_f1 As Double
        Dim dPdL_out_atmm As Double

        Const c_p = 0.000009871668   ' переводной коэффициент

        Ap_m2 = const_Pi * d_m ^ 2 / 4
        If Ql_rc_m3day + Qg_rc_m3day > 0 Then
            lambda_l = Ql_rc_m3day / (Ql_rc_m3day + Qg_rc_m3day)
        Else
            lambda_l = 1
        End If
        roughness_d = eps_m / d_m
        Vsl_msec = const_conver_sec_day * Ql_rc_m3day / Ap_m2
        Vsg_msec = const_conver_sec_day * Qg_rc_m3day / Ap_m2
        Vsm_msec = Vsl_msec + Vsg_msec
        Rho_n_kgm3 = rho_lrc_kgm3 * lambda_l + rho_grc_kgm3 * (1 - lambda_l) ' No-slip mixture density
        Mu_n_cP = Mul_rc_cP * lambda_l + Mug_rc_cP * (1 - lambda_l) ' No slip mixture viscosity
        n_re = 1000 * Rho_n_kgm3 * Vsm_msec * d_m / Mu_n_cP
        n_fr = Vsm_msec ^ 2 / (const_g * d_m)
        n_lv = Vsl_msec * (rho_lrc_kgm3 / (const_g * sigma_l_Nm)) ^ 0.25

        'liquid holdup is calculated by Gray Original formula
        If Vsm_msec = 0 Then
            h_l = 1
            correl3 = 1
            f = 0
        Else
            If Vsg_msec > 0 Then
                r = Vsl_msec / Vsg_msec
            Else
                r = 1000
            End If
            e = roughness_d
            E1 = 28.5 * sigma_l_Nm / (Rho_n_kgm3 * Vsm_msec ^ 2)
            If r >= 0.007 Then e = E1
            If r < 0.007 Then e = e + r * (E1 - e) / 0.007  ' corrected error by Kiyan Artem from /0.0007
            crit = 2.8
            Nv = Rho_n_kgm3 ^ 2 * Vsm_msec ^ 4 / (const_g * sigma_l_Nm * (rho_lrc_kgm3 - rho_grc_kgm3))
            Nd = const_g * (rho_lrc_kgm3 - rho_grc_kgm3) * d_m ^ 2 / sigma_l_Nm

            b = 0.0814 * (1 - 0.0554 * Log(1 + 730 * r / (r + 1)))
            h_l = (r + Exp(-2.314 * (Nv * (1 + 250 / Nd)) ^ b)) / (r + 1)
            'Calculate normalized friction factor
            f_n = unf_friction_factor(n_re, e, Payne_et_all_friction)
            'calculate friction factor correction for multiphase flow
            y = MaxReal(sigma_l_Nm / h_l ^ 2, 0.001)
            If (y > 1 And y < 1.2) Then
                S = Log(2.2 * y - 1.2)
            Else
                S = Log(y) / (-0.0523 + 3.182 * Log(y) - 0.8725 * (Log(y)) ^ 2 + 0.01853 * (Log(y)) ^ 4)
            End If
            f = f_n
        End If
        rho_s = rho_lrc_kgm3 * h_l + rho_grc_kgm3 * (1 - h_l) 'calculate mixture density
        dpdl_g = c_p * rho_s * const_g * sind(theta_deg)  'calculate pressure gradient due to gravity
        dpdl_f = c_p * f * Rho_n_kgm3 * Vsm_msec ^ 2 / (2 * d_m)  'calculate pressure gradient due to friction
        dpdl_g1 = c_p * Rho_n_kgm3 * const_g * sind(theta_deg)
        dpdl_f1 = c_p * f * Rho_n_kgm3 * Vsm_msec ^ 2 / (2 * d_m) 'calculate pressure gradient due to friction
        Select Case correl3
            Case 0
                dPdL_out_atmm = dpdl_g + dpdl_f
                dPdLg_out_atmm = dpdl_g
                dPdLf_out_atmm = dpdl_f
                dPdLa_out_atmm = 0
            Case 1
                dPdL_out_atmm = dpdl_g1 + dpdl_f1
                dPdLg_out_atmm = dpdl_g1
                dPdLf_out_atmm = dpdl_f1
                dPdLa_out_atmm = 0
        End Select
        Hl_out_fr = h_l
        fpat_out_num = 0
        unf_GrayModifiedGradient = {dPdLg_out_atmm * c_calibr_grav + dPdLf_out_atmm * c_calibr_fric,
                                dPdLg_out_atmm * c_calibr_grav,
                                dPdLf_out_atmm * c_calibr_fric,
                                    dPdLa_out_atmm,
                                    Vsl_msec,
                                    Vsg_msec,
                                    Hl_out_fr,
                                    fpat_out_num}
    End Function

    ' Федоров, Халиков (2016)
    '
    '
    '


    Public Function unf_HagedornandBrawnmodified(ByVal d_m As Double,
                                      ByVal theta_deg As Double, ByVal eps_m As Double,
                                      ByVal Ql_rc_m3day As Double, ByVal Qg_rc_m3day As Double,
                                      ByVal Mul_rc_cP As Double, ByVal Mug_rc_cP As Double,
                                      ByVal sigma_l_Nm As Double,
                                      ByVal rho_lrc_kgm3 As Double,
                                      ByVal rho_grc_kgm3 As Double,
                                      ByVal p_atma As Double,
                                      Optional Payne_et_all_holdup As Integer = 0,
                                      Optional Payne_et_all_friction As Integer = 1,
                                      Optional ByVal correl3 As Double = 0,
                                      Optional c_calibr_grav As Double = 1,
                                      Optional c_calibr_fric As Double = 1)

        'function for calculation of pressure gradient in pipe according to Begs and Brill method
        'Return (psi/ft (atma/m))
        'Arguments
        'd - pipe internal diameter (ft (m))
        'theta - pipe inclination angel (degrees)
        'eps - pipe wall roughness (ft (m))
        'p - reference pressure (psi (atma))
        'q_oSC - oil rate at standard conditions (Stb/day (m3/day))
        'q_wSC - water rate at standard conditions (Stb/day (m3/day))
        'q_gSC - gas rate at standard conditions (scf/day (m3/day))
        'b_o - oil formation volume factor at reference pressure (bbl/stb (m3/sm3))
        'b_w - water formation volume factor at reference pressure (bbl/stb (m3/sm3))
        'b_g - gas formation volume factorat reference pressure (ft3/scf (m3/sm3))
        'rs - gas-oil solution ratio at reference pressure (Scf/stb (sm3/sm3))
        'mu_o - oil viscosity at reference pressure (cp)
        'mu_w - water viscosity at reference pressure (cp)
        'mu_g - gAs viscosity at reference pressure (cp)
        'sigma_o - oil-gAs surface tension coefficient (dynes/sm (Newton/m))
        'sigma_w - water-gAs surface tension coefficient (dynes/sm (Newton/m))
        'rho_oSC - oil density at standard conditions (lbm/ft3 (kg/m3))
        'rho_wSC - water density at standard conditions (lbm/ft3 (kg/m3))
        'rho_gSC - gas density at standard conditions(lbm/ft3 (kg/m3))
        'units - input/output units (0-field, 1 - metric)
        'Payne_et_all_holdup - flag indicationg weather to applied Payne et all correction and holdup (0 - not applied, 1 - applied)
        'Payne_et_all_friction - flag indicationg weather to apply Payne et all correction for friction (0 - not applied, 1 - applied)
        'dpdl_g - used to otput pressure gradient due to gravity (psi/ft (atma/m))
        'dpdl_f - used to output pressure gradient due to friction (psi/ft (atma/m))
        'v_sl - used to output liquid superficial velocity (ft/sec (m/sec))
        'v_sg - used to output gas superficial velocity (ft/sec (m/sec))
        'h_l - used to output liquid holdup
        Dim roughness_d As Double
        Dim dPdLg_out_atmm As Double
        Dim dPdLf_out_atmm As Double
        Dim Hl_out_fr As Double
        Dim fpat_out_num
        Dim dPdLa_out_atmm As Double
        Dim Ap_m2 As Double ' площадь трубы
        Dim lambda_l As Double
        Dim Vsl_msec, Vsg_msec, Vsm_msec As Double
        Dim Rho_n_kgm3 As Double   ' no slip density
        Dim rho_s As Double        ' mix density
        Dim Mu_n_cP As Double
        Dim n_re As Double 'Reinolds number
        'Dim n_fr As Double 'Froude number
        'Dim n_lv1 As Double 'Liquid velocity number
        'Dim flow_pattern As Integer
        Dim f_n As Double ' normalized friction factor
        Dim f As Double ' friction factor
        ' Dim y, S As Double
        ' Dim r As Double 'dimensionless'superficial liquid to gas ratio parameter
        Dim e As Double    'Pipe relative roughness
        'Dim E1 As Double
        'Dim crit As Double
        'Dim Nv As Double ' dimensionless 'velocity number
        ' Dim Nd As Double 'nominal diameter
        '   Dim b As Double
        Dim h_l As Double 'liquid holdup by original Gray
        Dim dpdl_g As Double, dpdl_f As Double, dpdl_g1 As Double, dpdl_f1 As Double
        Dim dPdL_out_atmm As Double
        'Liquid velocity number
        Dim n_lv As Double
        Dim n_gv As Double
        Dim n_d As Double
        Dim n_l As Double

        Dim a As Double, HB_complex As Double
        Dim N_lc As Double
        'Dim AA As Double, 
        Dim Hl_phi As Double
        Dim b As Double
        Dim l As Double
        'Dim b_1 As Double
        Dim phi As Double
        Dim rho_tp As Double
        Dim mu_tp As Double

        Dim B0 As Double
        Dim b1 As Double
        Dim B2 As Double
        Dim B3 As Double
        Dim B4 As Double
        Dim C0 As Double
        Dim C1 As Double
        Dim C2 As Double
        Dim C3 As Double
        Dim C4 As Double


        Ap_m2 = const_Pi * d_m ^ 2 / 4
        If Ql_rc_m3day + Qg_rc_m3day > 0 Then
            lambda_l = Ql_rc_m3day / (Ql_rc_m3day + Qg_rc_m3day)
        Else
            lambda_l = 1
        End If
        roughness_d = eps_m / d_m
        Vsl_msec = const_conver_sec_day * Ql_rc_m3day / Ap_m2
        Vsg_msec = const_conver_sec_day * Qg_rc_m3day / Ap_m2
        Vsm_msec = Vsl_msec + Vsg_msec
        Rho_n_kgm3 = rho_lrc_kgm3 * lambda_l + rho_grc_kgm3 * (1 - lambda_l) ' No-slip mixture density
        Mu_n_cP = Mul_rc_cP * lambda_l + Mug_rc_cP * (1 - lambda_l) ' No slip mixture viscosity
        n_re = 1000 * Rho_n_kgm3 * Vsm_msec * d_m / Mu_n_cP
        'n_fr = Vsm_msec ^ 2 / (const_g * d_m)
        'n_lv1 = Vsl_msec * (rho_lrc_kgm3 / (const_g * sigma_l_Nm)) ^ 0.25

        If Vsm_msec > 0 Then

            B0 = -0.1030658
            b1 = 0.617774
            B2 = -0.632946
            B3 = 0.29598
            B4 = -0.0401

            C0 = 0.9116257
            C1 = -4.821756
            C2 = 1232.25
            C3 = -22253.58
            C4 = 116174.3

            'Determine Duns and ROs dimensionless groups

            'Determine Liquid velosity number
            n_lv = 1.938 * Vsl_msec / 0.3048 * (rho_lrc_kgm3 * 0.06243 / (sigma_l_Nm * 1000)) ^ 0.25
            ''Determine Gas velosity number
            n_gv = 1.938 * Vsg_msec / 0.3048 * (rho_lrc_kgm3 * 0.06243 / (sigma_l_Nm * 1000)) ^ 0.25
            ''Determine Diametr number
            n_d = 120.872 * d_m / 0.3048 * (rho_lrc_kgm3 * 0.06243 / (sigma_l_Nm * 1000)) ^ 0.5
            ''Determine Liquid viscosity number
            n_l = 0.15726 * Mul_rc_cP * (1 / (rho_lrc_kgm3 * 0.06243 * (sigma_l_Nm * 1000) ^ 3)) ^ 0.25

            ' проверим режим потока, чтобы определить надо использовать поправку Гриффитса для пузырькового режима
            a = 1.071 - (0.2218 * (Vsm_msec * const_convert_m_ft) ^ 2) / (d_m * const_convert_m_ft)
            If a < 0.13 Then a = 0.13
            b = Vsg_msec / Vsm_msec
            If (b - a) >= 0 Then  ' считаем по Хайгедорну Брауну
                n_lv = Vsl_msec * (rho_lrc_kgm3 / (const_g * sigma_l_Nm)) ^ 0.25 'Determine Liquid velosity number  dimensionless
                n_gv = Vsg_msec * (rho_lrc_kgm3 / (const_g * sigma_l_Nm)) ^ 0.25 'Determine Gas velosity number
                n_d = d_m * (rho_lrc_kgm3 * const_g / sigma_l_Nm) ^ 0.5 'Determine Diametr number
                n_l = Mul_rc_cP * const_convert_cP_Pasec * (const_g / (rho_lrc_kgm3 * sigma_l_Nm ^ 3)) ^ 0.25 'Determine Liquid viscosity number
                N_lc = -0.0259 * n_l ^ 4 + 0.1011 * n_l ^ 3 - 0.1272 * n_l ^ 2 + 0.0619 * n_l + 0.0018  ' корреляция подобрана по графуку рнт
                HB_complex = (n_lv / n_gv ^ 0.575) * (N_lc / n_d) * (p_atma) ^ 0.1
                Hl_phi = B0 + b1 * (Log10(HB_complex) + 6) + B2 * (Log10(HB_complex) + 6) ^ 2 + B3 * (Log10(HB_complex) + 6) ^ 3 + B4 * (Log10(HB_complex) + 6) ^ 4
                b = (n_gv * n_l ^ 0.38) / n_d ^ 2.14
                phi = C0 + C1 * b + C2 * b ^ 2 + C3 * b ^ 3 + C4 * b ^ 4

                h_l = Hl_phi * phi 'determine liquid holdup
                If h_l < lambda_l Then h_l = lambda_l


            Else    ' считаем по гриффитсу
                Dim vs_ftsec As Double, vs_msec As Double   ' bubble rise velocity
                vs_ftsec = 0.8
                vs_msec = vs_ftsec * const_convert_ft_m
                '  Dim Hg As Double
                h_l = 1 - 0.5 * (1 + Vsm_msec / vs_msec - ((1 + Vsm_msec / vs_msec) ^ 2 - 4 * (Vsg_msec / vs_msec)) ^ 0.5)
            End If


            rho_tp = rho_lrc_kgm3 * h_l + rho_grc_kgm3 * (1 - h_l)
            mu_tp = Mul_rc_cP ^ h_l + Mug_rc_cP ^ (1 - h_l)

            f_n = unf_friction_factor(n_re, e, Payne_et_all_friction)
            f = f_n

            rho_s = rho_lrc_kgm3 * h_l + rho_grc_kgm3 * (1 - h_l) 'calculate mixture density
            dpdl_g = c_p * rho_tp * const_g * sind(theta_deg) 'calculate pressure gradient due to gravity
            dpdl_f = c_p * f * Rho_n_kgm3 ^ 2 * Vsm_msec ^ 2 / (2 * d_m * rho_tp) 'calculate pressure gradient due to friction
        Else
            f = 0
            correl3 = 1
            rho_tp = Rho_n_kgm3
        End If

        dpdl_g1 = c_p * Rho_n_kgm3 * const_g * sind(theta_deg)
        dpdl_f1 = c_p * f * Rho_n_kgm3 ^ 2 * Vsm_msec ^ 2 / (2 * d_m * rho_tp) 'calculate pressure gradient due to friction


        Select Case correl3
            Case 0
                'calculate pressure gradient
                dPdL_out_atmm = dpdl_g + dpdl_f

                dPdLg_out_atmm = dpdl_g
                dPdLf_out_atmm = dpdl_f
                dPdLa_out_atmm = 0
            Case 1
                'calculate pressure gradient
                dPdL_out_atmm = dpdl_g1 + dpdl_f1

                dPdLg_out_atmm = dpdl_g1
                dPdLf_out_atmm = dpdl_f1
                dPdLa_out_atmm = 0
        End Select

        Hl_out_fr = h_l
        fpat_out_num = 0
        unf_HagedornandBrawnmodified = {dPdLg_out_atmm * c_calibr_grav + dPdLf_out_atmm * c_calibr_fric,
                                    dPdLg_out_atmm * c_calibr_grav,
                                    dPdLf_out_atmm * c_calibr_fric,
                                    dPdLa_out_atmm,
                                    Vsl_msec,
                                    Vsg_msec,
                                    Hl_out_fr,
                                    fpat_out_num}

    End Function

    Function unf_friction_factor(ByVal n_re As Double,
                             ByVal roughness_d As Double,
                    Optional ByVal friction_corr_type As Integer = 3,
                    Optional ByVal smoth_transition As Boolean = False) As Double
        'Calculates friction factor given pipe relative roughness and Reinolds number
        'Parameters
        'n_re - Reinolds number
        'roughness_d - pipe relative roughness
        'friction_corr_type - flag indicating correlation type selection
        ' 0 - Colebrook equation solution
        ' 1 - Drew correlation for smooth pipes
        '

        Dim f_n, f_n_new, f_int As Double
        Dim i As Integer
        Dim ed As Double
        Dim Svar As Double
        Dim f_1 As Double
        Dim Re_save As Double
        Const lower_Re_lim = 2000.0#
        Const upper_Re_lim = 4000.0#

        ed = roughness_d

        If n_re = 0 Then
            f_n = 0
        ElseIf n_re < lower_Re_lim Then 'laminar flow
            f_n = 64 / n_re
        Else 'turbulent flow
            Re_save = -1
            If smoth_transition And (n_re > lower_Re_lim And n_re < upper_Re_lim) Then
                ' be ready to interpolate for smooth transition
                Re_save = n_re
                n_re = upper_Re_lim
            End If
            Select Case friction_corr_type
                Case 0
                    'calculate friction factor for rough pipes according to Moody method - Payne et all modification for Beggs&Brill correlation
                    ' Zigrang and Sylvester  1982  https://en.wikipedia.org/wiki/Darcy_friction_factor_formulae
                    f_n = (2 * Log10(2 / 3.7 * ed - 5.02 / n_re * Log10(2 / 3.7 * ed + 13 / n_re))) ^ -2

                    i = 0
                    Do 'iterate until error in friction factor is sufficiently small
                        'https://en.wikipedia.org/wiki/Darcy_friction_factor_formulae
                        ' expanded form  of the Colebrook equation
                        f_n_new = (1.7384 - 2 * Log10(2 * ed + 18.574 / (n_re * f_n ^ 0.5))) ^ -2
                        i = i + 1
                        f_int = f_n
                        f_n = f_n_new
                        'stop when error is sufficiently small or max number of iterations exceedied
                    Loop Until (Abs(f_n_new - f_int) <= 0.001 Or i > 19)
                Case 1
                    'Calculate friction factor for smooth pipes using Drew correlation - original Begs&Brill with no modification
                    f_n = 0.0056 + 0.5 * n_re ^ -0.32

                Case 2
                    ' Zigrang and Sylvester  1982  https://en.wikipedia.org/wiki/Darcy_friction_factor_formulae
                    f_n = (2 * Log10(1 / 3.7 * ed - 5.02 / n_re * Log10(1 / 3.7 * ed + 13 / n_re))) ^ -2
                Case 3
                    ' Brkic shows one approximation of the Colebrook equation based on the Lambert W-function
                    '  Brkic, Dejan (2011). "An Explicit Approximation of Colebrook's equation for fluid flow friction factor" (PDF). Petroleum Science and Technology. 29 (15): 1596–1602. doi:10.1080/10916461003620453
                    ' http://hal.archives-ouvertes.fr/hal-01586167/file/article.pdf
                    ' https://en.wikipedia.org/wiki/Darcy_friction_factor_formulae
                    ' http://www.math.bas.bg/infres/MathBalk/MB-26/MB-26-285-292.pdf
                    Svar = Log(n_re / (1.816 * Log(1.1 * n_re / (Log(1 + 1.1 * n_re)))))
                    f_1 = -2 * Log10(ed / 3.71 + 2 * Svar / n_re)
                    f_n = 1 / (f_1 ^ 2)
                Case 4
                    ' from unified TUFFP model
                    ' Haaland equation   Haaland, SE (1983). "Simple and Explicit Formulas for the Friction Factor in Turbulent Flow". Journal of Fluids Engineering. 105 (1): 89–90. doi:10.1115/1.3240948
                    ' with smooth transition zone
                    Dim fr2 As Double, fr3 As Double
                    fr2 = 16.0# / 2000.0#
                    fr3 = 1.0# / (3.6 * Log10(6.9 / 3000.0# + (ed / 3.7) ^ 1.11)) ^ 2
                    If n_re = 0 Then
                        f_n = 0
                    ElseIf (n_re < 2000.0#) Then
                        f_n = 16.0# / n_re
                    ElseIf (n_re > 3000.0#) Then
                        f_n = 1.0# / (3.6 * Log10(6.9 / n_re + (ed / 3.7) ^ 1.11)) ^ 2
                    ElseIf (n_re >= 2000.0# And n_re <= 3000.0#) Then
                        f_n = fr2 + (fr3 - fr2) * (n_re - 2000.0#) / 1000.0#
                    End If
                    f_n = 4 * f_n
                Case 5
                    ' from unified TUFFP model
                    ' Haaland equation   Haaland, SE (1983). "Simple and Explicit Formulas for the Friction Factor in Turbulent Flow". Journal of Fluids Engineering. 105 (1): 89–90. doi:10.1115/1.3240948

                    f_n = 4.0# / (3.6 * Log10(6.9 / n_re + (ed / 3.7) ^ 1.11)) ^ 2



            End Select

            Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double

            If smoth_transition And Re_save > 0 Then
                x1 = lower_Re_lim
                y1 = 64.0# / lower_Re_lim
                x2 = n_re
                y2 = f_n
                f_n = ((y2 - y1) * Re_save + (y1 * x2 - y2 * x1)) / (x2 - x1)
            End If

        End If

        unf_friction_factor = f_n

    End Function

    Private Function h_l_arr_theta_deg(flow_pattern As Integer, ByVal lambda_l As Double, ByVal n_fr As Double,
                    ByVal n_lv As Double, ByVal arr_theta_deg As Double, ByVal Payne_et_all As Integer) As Double
        'function calculating liquid holdup
        'flow_pattern - flow pattern (0 -Segregated, 1 - Intermittent, 2 - Distributed)
        'lambda_l - volume fraction of liquid at no-slip conditions
        'n_fr - Froude number
        'n_lv - liquid velocity number
        'arr_theta_deg - pipe inclination angle, (Degrees)
        'payne_et_all - flag indicationg weather to applied Payne et all correction for holdup (0 - not applied, 1 - applied)
        'Constants to determine liquid holdup
        Dim a(2) As Double
        a(0) = 0.98
        a(1) = 0.845
        a(2) = 1.065
        Dim b(2) As Double
        b(0) = 0.4846
        b(1) = 0.5351
        b(2) = 0.5824
        Dim c(2) As Double
        c(0) = 0.0868
        c(1) = 0.0173
        c(2) = 0.0609
        'constants to determine liquid holdup correction
        Dim e(2) As Double
        e(0) = 0.011
        e(1) = 2.96
        e(2) = 1
        Dim f(2) As Double
        f(0) = -3.768
        f(1) = 0.305
        f(2) = 0
        Dim g(2) As Double
        g(0) = 3.539
        g(1) = -0.4473
        g(2) = 0
        Dim h(2) As Double
        h(0) = -1.614
        h(1) = 0.0978
        h(2) = 0

        Dim h_l_0 As Double
        Dim CC As Double
        Dim psi As Double
        Dim arr_theta_deg_d As Double

        h_l_0 = a(flow_pattern) * lambda_l ^ b(flow_pattern) / n_fr ^ c(flow_pattern) 'calculate liquid holdup at no slip conditions
        CC = MaxReal(0, (1 - lambda_l) * Log(e(flow_pattern) * lambda_l ^ f(flow_pattern) * n_lv ^ g(flow_pattern) * n_fr ^ h(flow_pattern))) 'calculate correction for inclination angle

        arr_theta_deg_d = const_Pi / 180 * arr_theta_deg 'convert angle to radians
        psi = 1 + CC * (Sin(1.8 * arr_theta_deg_d) - 0.333 * (Sin(1.8 * arr_theta_deg_d)) ^ 3)  ' corrected sign by issue #37
        'calculate liquid holdup with payne et al. correction factor
        If Payne_et_all > 0 Then
            If arr_theta_deg > 0 Then 'uphill flow
                h_l_arr_theta_deg = MaxReal(MinReal(1, 0.924 * h_l_0 * psi), lambda_l)
            Else  'downhill flow
                h_l_arr_theta_deg = MaxReal(MinReal(1, 0.685 * h_l_0 * psi), lambda_l)
            End If
        Else
            h_l_arr_theta_deg = MaxReal(MinReal(1, h_l_0 * psi), lambda_l)
        End If

    End Function

    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '     comprehensive mechanistic model for pressure gradient, liquid
    '     holdup and flow pattern predictions
    '     written by,    asfandiar m. ansari
    '     revised by,    asfandiar m. ansari
    '     revised by,    tuffp                  last revision: november 89
    '              * *  tulsa university fluid flow projects  * *
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '     this subroutine calculates two phase liquid holdup, flow pattern
    '     and pressure gradient using the mechanistic approach developed from
    '     the separate models for flow pattern prediction and flow behavior
    '     prediction of the individual flow patterns. the english system of
    '     units is used for the input data but converted to si units for the
    '     subsequent calculations.
    '                               reference
    '                               ---------
    '     1.  ansari, a. m., " mechanistic model for two-phase upward flow."
    '         m.s thesis, the university of tulsa (1988).
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '                   input/output logical file variables
    '                   -----------------------------------
    '     ioerr = output file for error messages when input values passed
    '             to the subroutine are out of range or error occurs in
    '             the calculation.
    '                          subsubroutines called
    '                          ------------------
    '     upfpdet = this subroutine predicts flow pattern only for upward
    '               flow using taitel, barnea, & dukler model.
    '     single = this subroutine calculates pressure gradient for single
    '              phase flow of liquid or gas.
    '     bubble = this subroutine calculates pressure gradient both for
    '              dispersed bubble and bubbly flows.
    '     slug   = this subroutine calculates pressure gradient for slug
    '              flow.
    '     anmist = this subroutine calculates pressure gradient for
    '              annular-mist flow.
    '                       variable description
    '                       --------------------
    '     *ang   = angle of flow from horizontal. (deg.)
    '      angr  = angle of flow from horizontal. (rad)
    '     *deng  = gas density. (lbm/ft^3)
    '     *denl  = liquid density. (lbm/ft^3)
    '     *di    = inside pipe diameter. (m)
    '      e     = liquid holdup fraction.
    '     *ed    = relative pipe roughness.
    '     *ens   = no-slip liquid holdup fraction.
    '      fpat  = flow pattern, (chr)
    '                 " liq" = liquid
    '                 " gas" = gas
    '                 "bubl" = bubbly
    '                 "slug" = slug
    '                 "dbub" = dispersed bubble
    '                 "anul" = annular
    '     *p     = pressure. (psia)
    '      pga   = acceleration pressure gradient. (psi/ft)
    '      pge   = elevation pressure gradient. (psi/ft)
    '      pgf   = friction pressure gradient. (psi/ft)
    '      pgt   = total pressure gradient. (psi/ft)
    '     *surl  = gas-liquid surface tension. (dynes/cm)
    '     *visg  = gas viscosity. (cp)
    '     *visl  = liquid viscosity. (cp)
    '     *vm    = mixture velocity. (ft/sec)
    '      vsg   = superficial gas velocity. (ft/sec)
    '      vsl   = superficial liquid velocity. (ft/sec)
    '      (*indicates input variables)
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    Private Sub Ansari(ang#, di_m#, ed#, p_atma#, vsl_m3sec#, vsg_m3sec#, ens#, deng_kgm3#, denl_kgm3#,
                       visg#, visl#, surl#, e#, pgf#, pge#, pga#, pgt#, fpat As String,
                Optional ByVal znlf As Boolean = False)

        '     --------------------------------
        '     initialize the output variables.
        '     --------------------------------
        e = 0#
        pgf = 0#
        pge = 0#
        pga = 0#
        pgt = 0#
        fpat = "    "
        '     --------------------------------------
        '     convert input variables into si units.
        '     --------------------------------------
        Dim angr As Double, p_Pa As Double
        angr = ang * 3.1416 / 180.0#
        p_Pa = p_atma * 101325.0# '/ 14.7
        visg = visg * 0.001
        visl = visl * 0.001
        '     ------------------------------------------
        '     check for single phase gas or liquid flow.
        '     ------------------------------------------
        If (ens > 0.99999) Then '        single phase liquid flow.
            fpat = " liq"
            Call single1(angr, di_m, ed, vsl_m3sec, denl_kgm3, visl, p_Pa, pgf, pge, pga, pgt)
            e = 1.0#
        ElseIf (ens < 0.00001) And Not znlf Then        '        single phase gas flow.
            fpat = " gas"
            Call single1(angr, di_m, ed, vsg_m3sec, deng_kgm3, visg, p_Pa, pgf, pge, pga, pgt)
            e = 0#
        Else
            '        -----------------------------------------------------------
            '        determine flow pattern using taitel, barnea & dukler model.
            '        -----------------------------------------------------------
            Call fpup(vsl_m3sec, vsg_m3sec, di_m, ed, denl_kgm3, deng_kgm3, visl, visg, ang, surl, fpat)
            If (fpat = "anul") Then '           annular-mist flow exists.
                Call anmist(angr, di_m, ed, denl_kgm3, deng_kgm3, visl, visg, vsl_m3sec, vsg_m3sec,
                                   surl, fpat, e, pgf, pge, pga, pgt)
                If (fpat = "slug") Then '              annular flow not confirmed. slug flow persists.
                    Call slug(angr, di_m, ed, denl_kgm3, deng_kgm3, visl, visg, vsl_m3sec, vsg_m3sec,
                                        surl, e, pgf, pge, pga, pgt)
                End If
            ElseIf (fpat = "slug") Then '           slug flow exists.
                Call slug(angr, di_m, ed, denl_kgm3, deng_kgm3, visl, visg, vsl_m3sec, vsg_m3sec,
                                 surl, e, pgf, pge, pga, pgt)
            ElseIf (fpat = "bubl" Or fpat = "dbub") Then '           bubble flow exists.
                Call bubble(angr, di_m, ed, denl_kgm3, deng_kgm3, visl, visg, vsl_m3sec, vsg_m3sec,
                                       surl, fpat, e, pgf, pge, pga, pgt)
            Else
                fpat = "  na"
                AddLogMsg("ansari: error in flow pattern detection")
                Exit Sub
            End If
        End If
        '     -----------------------------------------------------------
        '     convert pressure gradients and diameter into english units.
        '     -----------------------------------------------------------
        pge = pge / 101325.0#
        pgf = pgf / 101325.0#
        pga = pga / 101325.0#
        pgt = pgt / 101325.0#
        visg = visg * 1000.0#
        visl = visl * 1000.0#

    End Sub
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '     this subroutine detects the flow pattern for inclined
    '     and vertical upward flow (+15 to +90 degrees)
    '     written by,  caetano, shoham and triggia
    '     revised by,  lorri jefferson                      march 1989
    '     revised by,  guohua zheng         last revision:  april 1989
    '               * *  tulsa university fluid flow projects  * *
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '     this subroutine detects the flow pattern for inclined and vertical
    '     upward flow (+15 to +90 degrees).  the si system of units is used.
    '                                references
    '                                ----------
    '     1.  e.f. caetano, o. shoham and a.a. triggia, "gas liquid
    '            two-phase flow pattern prediction computer library",
    '            journal of pipelines, 5 (1986) 207-220.
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '                           variable description
    '                           --------------------
    '     alfa   = angle of flow from horizontal. (radians)
    '    *ang    = angle of flow from horizontal. (deg)
    '    *deng   = gas density. (kg/m^3)
    '    *denl   = liquid density. (kg/m^3)
    '    *di     = inside pipe diameter. (m)
    '    *ed     = relative pipe roughness
    '     fpat   = flow pattern, (chr)
    '                 " liq" = liquid
    '                 " gas" = gas
    '                 "bubl" = bubbly
    '                 "slug" = slug
    '                 "dbub" = dispersed bubble
    '                 "anul" = annular
    '    *visg    = gas viscosity. (cp)
    '    *visl    = liquid viscosity. (cp)
    '    *vsg     = superficial gas velocity. (m/sec)
    '    *vsl     = superficial liquid velocity. (m/sec)
    '    *surl    = gas-liquid surface tension. (dynes/cm)
    '     (*indicates input variables)
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^



    Private Sub fpup(vsl#, vsg#, Di#, ed#, denl#, deng#, visl#, visg#, ang#, surl#, fpat As String)
        ' sub parameter : do not dim ! Dim fpat as string * 4
        Dim alfa#, vsgo#, vsg1#, vsl1#, vsg2#, vsl2#, vsg3#, vsl3#
        Dim vslb As Double
        Dim vsgb As Double

        alfa = 0.0174533 * ang
        Call mpoint(Di, ed, denl, deng, visl, visg, ang, surl, vsgo, vsg1, vsl1, vsg2, vsl2, vsg3, vsl3) '     calculate points on transition boundaries
        '     ----------------------
        '     check for annular flow
        '     ----------------------
        If Not (vsg < vsg3) Then
            fpat = "anul"
            Exit Sub
        End If
        '     -----------------------------------------
        '     check for bubble/slug or dispersed-bubble
        '     -----------------------------------------
        If Not (vsg > vsg2) Then
            Call dbtran(0#, vslb, vsg, Di, ed, denl, deng, visl, visg, ang, surl)
            If (vsl < vslb) Then
                If (vsgo > 0#) Then
                    vsgb = (vsl + 1.15 * (const_g * (denl - deng) * surl / denl ^ 2) ^ 0.25 * Sin(alfa)) / 3.0#
                    If (vsg > vsgb) Then
                        fpat = "slug"
                    Else
                        fpat = "bubl"
                    End If
                Else
                    fpat = "slug"
                End If
            Else
                fpat = "dbub"
            End If
        Else ' (vsg > vsg2)
            vslb = vsg / 0.76 - vsg
            If (vsl >= vslb) Then
                fpat = "dbub"
            Else
                fpat = "slug"
            End If
        End If
    End Sub

    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '     transition boundaries for upward vertical flow
    '     written by,  caetano, shoham, and triggia
    '     revised by,  lorri jefferson                      march 1989
    '     revised by,  guohua zheng         last revision:  april 1989
    '               * *  tulsa university fluid flow projects  * *
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '     this subroutine calculates points on the transition boundaries for
    '     upward vertical flow.  the si system of units is used.
    '                                references
    '                                ----------
    '     1.  e.f. caetano, o. shoham and a.a. triggia, "gas liquid
    '            two-phase flow pattern prediction computer library",
    '            journal of pipelines, 5 (1986) 207-220.
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '                          variable description
    '                          --------------------
    '     alfa   = angle of flow from horizontal. (radians)
    '    *ang    = angle of flow from horizontal. (deg)
    '    *deng   = gas density. (kg/m^3)
    '    *denl   = liquid density. (kg/m^3)
    '    *di     = inside pipe diameter. (m)
    '    *ed     = relative pipe roughness
    '    *visg   = gas viscosity. (cp)
    '    *visl   = liquid viscosity. (cp)
    '     vsgs   = superficial gas velocity on transition boundaries.
    '              (m/sec)
    '     vsls   = superficial liquid velocity on transition boundaries.
    '              (m/sec)
    '    *surl   = gas-liquid surface tension. (dynes/cm)
    '     (*indicates input variables)
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    Private Sub mpoint(Di#, ed#, denl#, deng#, visl#, visg#, ang#, surl#, vsgo#, vsg1#, vsl1#, vsg2#, vsl2#, vsg3#, vsl3#)
        '   iErr = 0

        Dim alfa#, DMin#
        Dim vsl As Double

        alfa = 0.0174533 * ang
        '     -----------------------------------------------------------
        '     calculate vsgo
        '     minimum pipe diameter and inclination angle for bubble flow
        '     existed at low liquid flow rates
        '     -----------------------------------------------------------
        DMin = 19.0# * Sqrt((denl - deng) * surl / (denl ^ 2 * const_g))
        If (ang > 70.0# And Di > DMin * 0.95) Then
            vsl = 0.001
            vsgo = (vsl + 1.15 * (const_g * (denl - deng) * surl / denl ^ 2) ^ 0.25 * Sin(alfa)) / 3.0#
        Else
            vsgo = -1.0#
        End If
        vsg3 = 3.1 * (surl * const_g * Sin(alfa) * (denl - deng)) ^ 0.25 / Sqrt(deng) '     calculate vsg3
        '     --------------
        '     calculate vsg1
        '     --------------
        vsg1 = -1.0#
        vsl1 = -1.0#
        If (vsgo > 0#) Then Call dbtran(0.25, vsl1, vsg1, Di, ed, denl, deng, visl, visg, ang, surl)
        '     --------------
        '     calculate vsg2
        '     --------------
        vsg2 = 0.2
        Call dbtran(0.76, vsl2, vsg2, Di, ed, denl, deng, visl, visg, ang, surl)
        If (vsg2 >= vsg3) Then
            vsg2 = vsg3
            Call dbtran(0#, vsl2, vsg2, Di, ed, denl, deng, visl, visg, ang, surl)
            vsl3 = vsl2
            If (vsg1 < vsg2) Then GoTo L999
            vsg1 = vsg2
            GoTo L999
        End If
        '     -----------------------------------
        '     calculate vsl3 on boundary line "c"
        '     -----------------------------------
        vsl3 = (vsg3 / 0.76 - vsg3)
L999:

    End Sub
    Private Sub anmist(angr#, Di#, ed#, denl#, deng#, visl#, visg#, vsl#, vsg#, surl#, fpat As String, e#, pgf#, pge#, pga#, pgt#)

        'Dim nf As Double
        'Dim NC As Double
        '     --------------------------------------
        '     calculate fe using wallis correlation.
        '     --------------------------------------
        Dim X As Double, fe As Double
        Dim c As Double
        Dim alfc As Double, vsc#, denc#, visc#, recs#, ffcs#, rels#, ffls#, relf#, fflf#, a#
        Dim pgfcs#, pgfls#, xm2#, xmo2#, ym#
        Dim deldmx#, deldmn#, deldac#, deld#, ec#, phic2#, phif2#

        X = (deng / denl) ^ 0.5 * 10000.0# * vsg * visg / surl
        fe = 1.0# - Exp(-0.125 * (X - 1.5))
        If (fe <= 0#) Then fe = 0#
        If (fe >= 1.0#) Then fe = 1.0#
        '     -----------------------------------------------------------
        '     use appropriate correlation factor for interfacial friction
        '     according to the entrainment fraction.
        '     -----------------------------------------------------------
        If (fe > 0.9) Then
            c = 300.0# '        use wallis correlation factor.
        Else
            c = 24.0# * (denl / deng) ^ (1.0# / 3.0#) '        use whalley correlation factor.
        End If
        '     ---------------------------------------------------
        '     calculate superficial pressure gradients for entire
        '     liquid and gas-liquid core.
        '     ---------------------------------------------------
        alfc = 1.0# / (1.0# + fe * vsl / vsg)
        vsc = vsg + fe * vsl
        denc = deng * alfc + denl * (1.0# - alfc)
        visc = visg * alfc + visl * (1.0# - alfc)
        recs = denc * vsc * Di / visc
        ffcs = unf_friction_factor(recs, ed, 2) / 4  '  Fanning friction factor required
        rels = denl * vsl * Di / visl
        ffls = unf_friction_factor(rels, ed, 2) / 4  '  Fanning friction factor required
        If (fe < 0.9999) Then
            relf = denl * vsl * (1 - fe) * Di / visl
            fflf = unf_friction_factor(relf, ed, 2) / 4  '  Fanning friction factor required
            a = (1.0# - fe) ^ 2 * (fflf / ffls)
        Else
            a = 1.0#
        End If

        pgfcs = 4.0# * ffcs * denc * vsc * vsc / (2.0# * Di)
        pgfls = 4.0# * ffls * denl * vsl * vsl / (2.0# * Di)
        '     --------------------------------------------------------
        '     calculate modified lockhart and martinelli parameters as
        '     defined by alves including entrainment fraction.
        '     --------------------------------------------------------
        xm2 = pgfls / pgfcs
        xmo2 = xm2 * a
        ym = 9.81 * Sin(angr) * (denl - denc) / pgfcs
        '     ------------------------------------------------------------
        '     calculate film thickness if entrainment is less than 99.99%.
        '     ------------------------------------------------------------
        If (fe < 0.9999) Then
            deldmx = 0.499
            deldmn = 0.000001
            deldac = 0.000001
            deld = itsafe(xmo2, ym, c, 0#, 0#, 0#, 4, deldmn, deldmx, deldac)
            '        ----------------------------------------------
            '        check whether annular flow could exist or not.
            '        ----------------------------------------------
            ec = 1.0# - alfc
            Call chkan(xmo2, ym, deld, ec, e, fpat)
            If (fpat = "slug") Then
                '            -----------------------------------------------------
                '            annular flow not confirmed by barnea"s criteria. slug
                '            flow continues to exist.
                '            -----------------------------------------------------
                Exit Sub
            End If
            '        ------------------------------------------------------
            '        calculate dimensionless groups defined by alves.
            '        ------------------------------------------------------
            phic2 = (1.0# + c * deld) / (1.0# - 2.0# * deld) ^ 5
            phif2 = (phic2 - ym) / xm2
        Else
            '        ------------------------------------------------------
            '        assume 100 % entrainment and therefore no liquid film.
            '        ------------------------------------------------------
            fe = 1.0#
            deld = 0#
            phic2 = 1.0#
            phif2 = 0#
            e = 1.0# - alfc
            If (e > 0.12) Then
                fpat = "slug"
                Exit Sub
            End If
        End If
        '     -----------------------------------------
        '     calculate pressure gradients in the core.
        '     -----------------------------------------
        Dim pgec#, pgfc#, pgtc#, pgef#, pgff#, pgtf#
        pgec = 9.81 * denc * Sin(angr)
        pgfc = pgfcs * phic2
        pgtc = pgec + pgfc
        '     -----------------------------------------
        '     calculate pressure gradients in the film.
        '     -----------------------------------------
        pgef = 9.81 * denl * Sin(angr)
        pgff = pgfls * phif2
        pgtf = pgef + pgff
        '     --------------------------------------------------------------
        '     assume core pressure gradients to be the gradients for annular
        '     flow pattern. the total pressure gradient can be that of film
        '     or core.
        '     --------------------------------------------------------------
        pge = pgec
        pgf = pgfc
        pgt = pgtc
        pga = 0#

    End Sub

    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '     dispersed bubble transition.
    '     written by,  caetano, triggia and shoham
    '     revised by,  lorri jefferson                      march 1989
    '     revised by,  guohua zheng         last revision:  april 1989
    '               * *  tulsa university fluid flow projects  * *
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '     this subroutine determines dispersed bubble transition boundaries.
    '     the si system of units is used.
    '                                references
    '                                ----------
    '     1.  e.f. caetano, o. shoham and a.a. triggia, "gas liquid
    '            two-phase flow pattern prediction computer library",
    '            journal of pipelines, 5 (1986) 207-220.
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '                          variable description
    '                          --------------------
    '    *ang    = angle of flow from horizontal. (deg)
    '    *deng   = gas density. (kg/m^3)
    '    *denl   = liquid density. (kg/m^3)
    '    *di     = inside pipe diameter. (m)
    '    *ed     = relative pipe roughness
    '    *hgg    = guessed gas void fraction
    '    *visg   = gas viscosity. (cp)
    '    *visl   = liquid viscosity. (cp)
    '    *vsg    = superficial gas velocity. (m/sec)
    '    *vsl    = superficial liquid velocity. (m/sec)
    '     (*indicates input variables, vsg and vsl Close #ioerr:exit subed to calling
    '      subroutine)
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    Private Sub dbtran(hgg#, vsl#, vsg#, Di#, ed#, denl#, deng#, visl#, visg#, ang#, surl#)
        Dim vmc As Double, ratio As Double
        Dim c As Double, vme As Double, iter As Double
        Dim Hg As Double
        Dim rhom, vism, FFM As Double
        Dim Re As Double
        Dim VM As Double

        '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        '     trial and error calculation of dispersed bubble transition
        '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        c = 2.0# * ((0.4 * surl) / ((denl - deng) * const_g)) ^ 0.5 * (denl / surl) ^ 0.6 * (2.0# / Di) ^ 0.4
        vme = vsg + 1.5 '     estimate a mixture velocity
        iter = 0
        For iter = 0 To 50
            If (hgg = 0#) Then
                Hg = vsg / vme
            Else
                Hg = hgg
                vsg = Hg * vme
                vsl = vme - vsg
            End If
            rhom = denl * (1.0# - Hg) + deng * Hg
            vism = visl * (1.0# - Hg) + visg * Hg
            Re = Di * rhom * vme / vism
            FFM = unf_friction_factor(Re, ed) '     get frictional factor
            vmc = ((0.725 + 4.15 * Sqrt(Hg)) / c / (FFM / 4.0#) ^ 0.4) ^ 0.8333 '     calculate new mixture velocity
            ratio = vmc / vme
            If (ratio >= 0.99 And ratio <= 1.01) Then '     check for convergence
                Exit For
            End If
            vme = (vmc + vme) / 2.0#
        Next iter
        If ratio < 0.8 Then
            AddLogMsg("dbtran: calculation proceeds without convergence on vm after 50 iterations. ratio = " & String.Format("0.000", ratio))
        End If
        VM = (vmc + vme) / 2
        vsl = VM * (1.0# - hgg)
        If (hgg > 0#) Then vsg = VM * hgg

    End Sub



    '================================================ Ansari =======================================================================

    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '     mechanistic model for pressure gradient in single phase (liquid
    '     or gas) flow.
    '     written by,    asfandiar m. ansari
    '     revised by,    asfandiar m. ansari     last revision: march 1989
    '              * *  tulsa university fluid flow projects  * *
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '     this subroutine calculates single phase pressure gradient using
    '     simple mechanistic approach. an explicit equation developed by
    '     zigrang and sylvester is used for friction factor.the si system
    '     of units is used.
    '                       variable description
    '                       --------------------
    '     *angr  = angle of flow from horizontal. (rad)
    '     *den   = density of liquid or gas. (kg/cum)
    '     *di    = inside pipe diameter. (m)
    '     *ed    = relative pipe roughness.
    '      ekk   = kinetic energy term used to determine if critical flow
    '              exists.
    '      ff    = friction factor.
    '      icrit = critical flow indicator (0-noncritical, 1-critical)
    '      ierr  = error code. (0=ok, 1=input variables out of range,
    '              2=extrapolation of correlation occurring)
    '     *ioerr = output file for error messages when input values
    '              passed to the subroutine are out of range.
    '     *p     = pressure. (pa)
    '      pga   = acceleration pressure gradient. (pa/m)
    '      pge   = elevation pressure gradient. (pa/m)
    '      pgf   = friction pressure gradient. (pa/m)
    '      pgt   = total pressure gradient. (pa/m)
    '      re    = reynolds number for liquid or gas.
    '     *vis   = viscosity. of liquid or gas (kg/m-s)
    '     *v     = velocity. of liquid or gas (m/s)
    '      (*indicates input variables)
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


    Private Sub single1(angr#, Di#, ed#, v#, den#, vis#, p#, pgf#, pge#, pga#, pgt#)
        Dim Re As Double, FF As Double
        Dim ekk As Double

        pge = den * Sin(angr) * 9.81 '     calculate elevation pressure gradient.
        If v > 0 Then
            Re = Di * den * v / vis
            FF = unf_friction_factor(Re, ed, 2) '     calculate frictional pressure gradient.
        Else
            FF = 0
        End If
        pgf = 0.5 * den * FF * v * v / Di
        ekk = den * v * v / p
        If (ekk > 0.95) Then ekk = 0.95
        pgt = (pge + pgf) / (1.0# - ekk)
        pga = pgt * ekk '     calculate accelerational pressure gradient.
        pgt = (pge + pgf)
    End Sub

    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^*
    '     mechanistic model for pressure gradient and liquid holdup in
    '     bubble flow.
    '     written by,    asfandiar m. ansari
    '     revised by,    asfandiar m. ansari     last revision: march 1989
    '              * *  tulsa university fluid flow projects  * *
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^*
    '     this subroutine calculates liquid holdup and pressure gradient
    '     both for dispersed bubble and bubbly flows using a mechanistic
    '     approach. for dispersed bubble flow the subroutine assumes no
    '     slippage, whereas for bubbly flow slippage is considered between
    '     the two phases. an explicit equation developed by zigrang and
    '     sylvester is used for friction factor. the si system of units is
    '     used.
    '                               references
    '                               ----------
    '     1.  ansari, a. m. and sylvester, n. d., " a mechanistic model for
    '         upward bubble flow in pipes ", aiche j., 8, 34, 1392-1394,
    '         (aug 1988).
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^*
    '                       variable description
    '                       --------------------
    '     *angr  = angle of flow from horizontal. (rad)
    '      den   = slip / no-slip density (kg/cum)
    '     *deng  = gas density. (kg/cum)
    '     *denl  = liquid density. (kg/cum)
    '      df    = derivative of the function used in newton-raphson method.
    '     *di    = inside pipe diameter. (m)
    '      e     = liquid holdup fraction.
    '      eacc  = accuracy required in iteration for e.
    '     *ed    = relative pipe roughness.
    '      emax  = upper limit for e during iteration.
    '      emin  = lower limit for e during iteration.
    '      ens   = no-slip liquid holdup fraction.
    '      emin  = lower limit for e during iteration.
    '      f     = function defined for newton-raphson method.
    '      ff    = friction factor
    '      fpat  = flow pattern, (chr)
    '                 "dbub" = dispersed bubble
    '                 "bubl" = bubbly
    '      ierr  = error code. (0=ok, 1=input variables out of range,
    '              2=extrapolation of correlation occurring)
    '     *ioerr = output file for error messages when input values
    '              passed to the subroutine are out of range.
    '      pga   = acceleration pressure gradient. (pa/m)
    '      pge   = elevation pressure gradient. (pa/m)
    '      pgf   = friction pressure gradient. (pa/m)
    '      pgt   = total pressure gradient. (pa/m)
    '      re    = reynolds number.
    '     *visg  = gas viscosity. (kg/m-s)
    '     *visl  = liquid viscosity. (kg/m-s)
    '      visns = no-slip viscosity. (kg/m-s)
    '      vs    = slip velocity (m/s)
    '     *vsg   = superficial gas velocity. (m/s)
    '     *vsl   = superficial liquid velocity. (m/s)
    '      (*indicates input variables)
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^




    Private Sub bubble(angr#, Di#, ed#, denl#, deng#, visl#, visg#, vsl#, vsg#, surl#, fpat As String, e#, pgf#, pge#, pga#, pgt#)

        '     --------------------------------------
        '     calculate slip and no-slip parameters.
        '     --------------------------------------
        Dim ens#, visns#, vs#
        ens = vsl / (vsl + vsg)
        visns = visl * ens + visg * (1.0# - ens)
        vs = 1.53 * (surl * 9.81 * (denl - deng) / denl ^ 2) ^ 0.25
        If (fpat = "dbub") Then
            e = ens '        dispersed  bubble flow exists, calculate no-slip holdup.
        Else
            '        --------------------------------------------------
            '        bubbly flow exists, calculate actual liquid holdup
            '        using function itsafe for iteration.
            '        --------------------------------------------------
            Dim emin#, emax#, eacc#
            emin = ens
            emax = 0.999
            eacc = 0.001
            e = itsafe(vsl, vsg, vs, 0#, 0#, 0#, 1, emin, emax, eacc)
        End If
        '     --------------------------------------
        '     calculate elevation pressure gradient.
        '     --------------------------------------
        Dim den#
        den = denl * e + deng * (1.0# - e)
        pge = den * Sin(angr) * 9.81
        '     ---------------------------------------
        '     calculate frictional pressure gradient.
        '     ---------------------------------------
        Dim Re#, FF#
        Re = den * (vsl + vsg) * Di / visns
        FF = unf_friction_factor(Re, ed, 2)
        pgf = 0.5 * den * FF * (vsl + vsg) ^ 2 / Di
        '     ---------------------------------------------------------
        '     calculate total pressure gradient neglecting acceleration
        '     component.
        '     ---------------------------------------------------------
        pga = 0#
        pgt = pge + pgf + pga

    End Sub

    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^*
    '     mechanistic model for pressure gradient and liquid holdup in
    '     slug flow.
    '     written by,    asfandiar m. ansari
    '     revised by,    asfandiar m. ansari       last revision: nov 1988
    '              * *  tulsa university fluid flow projects  * *
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^*
    '     this subroutine calculates liquid holdup and pressure gradient
    '     for slug flow based on the flow mechanics. the concept of develop-
    '     ing slug flow adopted by e.f. caetano is incorporated in the model.
    '     the si system of units is used.
    '                               references
    '                               ----------
    '     1.  sylvester, n. d., " a mechanistic model for two-phase
    '         vertical slug flow in pipes ", asme j.energy resources tech.,
    '         vol. 109,(1987),206-213.
    '     2.  caetano, e. f., " upward vertical two-phase flow through an
    '         annulus ",phd dissertation, the university of tulsa (1985)
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '                       variable description
    '                       --------------------
    '     alfacc = accuracy required in iteration for alftb.
    '     alfls  = void fraction in liquid slug.
    '     alfns  = no-slip void fraction in liquid slug.
    '     alfmax = upper limit for alftb during iteration.
    '     alfmin = lower limit for alftb during iteration.
    '     alftb  = av. void fraction in taylor bubble for a developed flow or
    '              the local void fraction at the bubble tail for a develop-
    '              ing slug flow.
    '     alftba = av.void fraction in taylor bubble for a developing flow.
    '     alftbn = void fraction in taylor bubble at nusselt thickness.
    '     alfsu  = void fraction in slug unit for developed slug flow.
    '    *angr   = angle of flow from horizontal. (rad)
    '     beta   = ratio of ltb and lsu.
    '     delacc = accuracy required in iteration for deln. (m)
    '     delmax = upper limit for deln during iteration. (m)
    '     delmin = lower limit for deln during iteration. (m)
    '     deln   = nusselt film thickness. (m)
    '    *deng   = gas density. (kg/cum)
    '    *denl   = liquid density. (kg/cum)
    '     denns  = no-slip density. (kg/cum)
    '     dens   = slip density. kg.cum)
    '    *di     = inside pipe diameter. (m)
    '     e      = liquid holdup fraction.
    '     esu    = liquid holdup fraction for a slug unit.
    '    *ed     = relative pipe roughness.
    '     ff     = friction factor
    '     ffls   = friction factor for liquid slug.
    '     ind    = indicator for the flow,
    '                0 = developed flow.
    '                + = developing flow.
    '     lc     = length of taylor bubble cap in developing slug flow.(m)
    '     lls    = length of liquid slug. (m)
    '     lsu    = length of slug unit for developed slug flow. (m)            slug
    '     lsua   = length of slug unit for developing slug flow. (m)
    '     ltb    = length of taylor bubble in developed slug flow. (m)
    '     ltba   = length of taylor bubble in developing slug flow. (m)
    '     pga    = acceleration pressure gradient. (pa/m)
    '     pgels  = elevation pressure gradient for liquid slug. (pa/m)
    '     pgfls  = friction pressure gradient for liquid slug. (pa/m)
    '     pgt    = total pressure gradient. (pa/m)
    '     rels   = reynolds number for liquid slug.
    '     vgls   = velocity of gas in liquid slug. (m/s)
    '     vgtb   = velocity of gas in taylor bubble. (m/s)
    '    *visg   = gas viscosity. (kg/m-s)
    '    *visl   = liquid viscosity. (kg/m-s)
    '     visns  = no-slip viscosity. (kg/m-s)
    '     vlls   = velocity of liquid in  liquid slug. (m/s)
    '     vltb   = velocity of liquid in taylor bubble. (m/s)
    '     vmls   = velocity of mixture in liquid slug. (m/s)
    '     vs     = slip velocity. (m/s)
    '    *vsg    = superficial gas velocity. (m/s)
    '     vsgls  = superficial gas velocity in liquid slug. (m/s)
    '    *vsl    = superficial liquid velocity. (m/s)
    '     vslls  = superficial liquid velocity in liquid slug. (m/s)
    '     c,f and df are dummy variables.
    '     (*indicates input variables)
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    Private Sub slug(angr#, Di#, ed#, denl#, deng#, visl#, visg#, vsl#, vsg#, surl#, e#, pgf#, pge#, pga#, pgt#)


        Dim lls As Double
        Dim lsu As Double
        Dim ltb As Double
        Dim ltb1 As Double
        Dim ltb2 As Double
        Dim LC As Double
        'Dim msg As String
        Dim ind As Integer, c#, d#, f#, h#, alftba#, alfsu#, esu#, dens#, pgels#
        Dim vtb#, vs#, alfls#, alfmin#, alfmax#, alfacc#, alftb#, vltb#, vlls#, vgls#, vgtb#, beta1#, beta2#, Beta#, g_#, delmin#, delmax#, delacc#, deln#, alftbn#, vgtbn#, vltbn#
        '     -----------------------------------------------
        '     assume lls to be 30 times the diameter of pipe.
        '     -----------------------------------------------
        lls = 30.0# * Di
        '     -----------------------------------------------------
        '     calculate void fraction in taylor bubble using itsafe
        '     and assuming developed slug flow.
        '     -----------------------------------------------------
        vtb = 1.2 * (vsl + vsg) + 0.35 * Sqrt(9.81 * Di * (denl - deng) / denl)
        vs = 1.53 * (9.81 * surl * (denl - deng) / denl ^ 2) ^ 0.25
        alfls = vsg / (0.425 + 2.65 * (vsl + vsg))
        alfmin = 0.7
        alfmax = 0.9999
        alfacc = 0.000001

        alftb = itsafe(vsl, vsg, vtb, vs, alfls, Di, 2, alfmin, alfmax, alfacc)
        '     --------------------------------
        '     calculate additional parameters.
        '     --------------------------------
        vltb = 9.916 * Sqrt(9.81 * Di * (1.0# - Sqrt(alftb)))
        vlls = vtb - (vtb + vltb) * (1.0# - alftb) / (1.0# - alfls)
        vgls = 1.2 * (vsl + vsg) + 1.53 * (9.81 * surl * (denl - deng) / denl ^ 2) ^ 0.25 * (1.0# - alfls) ^ 0.5
        If (alfls > 0.25) Then vgls = vlls
        vgtb = vtb * (1.0# - alfls / alftb) + vgls * alfls / alftb
        beta1 = (vlls * (1.0# - alfls) - vsl) / (vltb * (1.0# - alftb) + vlls * (1.0# - alfls))
        beta2 = (vsg - alfls * vgls) / (alftb * vgtb - alfls * vgls)
        If (Abs(beta1 - beta2) > 0.1) Then
            AddLogMsg("   slug: error in beta conv.")
        End If
        Beta = (beta1 + beta2) / 2.0#
        If (Beta <= 0# Or Beta >= 1.0#) Then
            AddLogMsg("   slug: unreal value for beta")
        End If
        lsu = lls / (1.0# - Beta)
        ltb = lsu - lls
        '     ---------------------------------------------------------
        '     calculate nusselt film thickness iteratively using itsafe
        '     ---------------------------------------------------------
        g_ = visl / (9.81 * (denl - deng))
        delmin = 0.00001
        delmax = 0.499 * Di
        delacc = 0.000001
        deln = itsafe(Di, vtb, vgls, vsl, vsg, g_, 3, delmin, delmax, delacc)
        '     --------------------------------------------------------
        '     calculate lc using the values of the parameters at deln.
        '     --------------------------------------------------------
        alftbn = (1.0# - 2.0# * deln / Di) ^ 2
        vgtbn = (vtb * alftbn - (vtb - vgls) * alfls) / alftbn
        vltbn = (vgtbn * alftbn - (vsl + vsg)) / (1.0# - alftbn)
        LC = (vltbn + vtb) ^ 2 / (2.0# * 9.81)
        '     ---------------------------------
        '     check for the nature of the flow.
        '     ---------------------------------
        If (LC > (0.75 * ltb)) Then
            '        --------------------------------------------------------
            '        developing slug flow exists. calculate new values for
            '        slug flow parameters starting with the length of
            '        taylor  bubble by solving a quadratic equation.
            '        --------------------------------------------------------
            ind = 1
            c = (vsg - vgls * alfls) / vtb
            d = 1.0# - vsg / vtb
            e = vtb - vlls
            f = (-2.0# * d * c * lls - 2.0# * (e * (1.0# - alfls)) ^ 2 / 9.81) / d ^ 2
            g_ = (c * lls / d) ^ 2
            h = f * f - 4.0# * g_
            If (h <= 0#) Then
                AddLogMsg("   slug: error in solving for ltb")
            End If
            ltb1 = (-f + Sqrt(h)) / 2.0#
            ltb2 = (-f - Sqrt(h)) / 2.0#
            If (ltb1 <= 0# And ltb2 <= 0#) Then
                AddLogMsg("   slug: error in ltb root")
            End If
            If (ltb1 > ltb2) Then ltb = ltb1
            If (ltb2 > ltb1) Then ltb = ltb2
            alftba = 1.0# - 2.0# * (vtb - vlls) * (1.0# - alfls) / Sqrt(2.0# * 9.81 * ltb)
            lsu = ltb + lls
            Beta = ltb / lsu
        Else
            '     ---------------------------------------------------------------
            '     developed slug flow exists. no new values of the parameters are
            '     required.
            '     ---------------------------------------------------------------
            ind = 0
        End If
        '     ----------------------------------------
        '     calculate liquid holdup for a slug unit.
        '     ----------------------------------------
        alfsu = alftb * Beta + alfls * (1.0# - Beta)
        If (ind = 1) Then alfsu = alftba * Beta + alfls * (1.0# - Beta)
        esu = 1.0# - alfsu
        '     -----------------------------------------------------
        '     calculate elevation pressure gradient for liquid slug
        '     using its slip density.
        '     -----------------------------------------------------
        dens = denl * (1.0# - alfls) + deng * alfls
        pgels = 9.81 * Sin(angr) * dens
        '     -----------------------------------------------------------
        '     calculate elevation pressure gradient for taylor  bubble
        '     using its average void fraction.
        '     -----------------------------------------------------------
        Dim pgetb#, vslls#, vsgls#, vmls#, alfns#, visns#, rels#, ffls#, pgfls#
        If (ind = 1) Then
            pgetb = 9.81 * Sin(angr) * (deng * alftba + denl * (1 - alftba))
        Else
            pgetb = 9.81 * Sin(angr) * deng
        End If
        '     --------------------------------------------------
        '     calculate frictional pressure gradient for liquid
        '     slug.
        '     --------------------------------------------------
        vslls = vlls * (1.0# - alfls)
        vsgls = vgls * alfls
        vmls = vslls + vsgls
        alfns = vsgls / vmls
        visns = visl * (1.0# - alfns) + visg * alfns
        rels = dens * vmls * Di / visns
        ffls = unf_friction_factor(rels, ed, 2)
        pgfls = dens * vmls ^ 2 * ffls / (2.0# * Di)
        pga = 0#    '     acceleration pressure gradient over a slug unit is zero.
        '     ---------------------------------------------------------
        '     assume constant pressure gradients and holdup for all the
        '     slug units within one pipe increment.
        '     ---------------------------------------------------------
        e = esu
        pge = pgels * (1.0# - Beta) + pgetb * Beta
        pgf = pgfls * (1.0# - Beta)
        pgt = pge + pgf + pga
        Exit Sub

    End Sub

    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^*
    '     this subroutine checks the taitel,barnea & dukler prediction of an-
    '     nular flow by using criteria developed by barnea. it checks annular
    '     flow bridging caused by liquid holdup of greater than 0.24. it also
    '     calculates maximum stable film thickness for annular flow and com-
    '     pares it with the existing film thickness for the stability of the
    '     annular flow. the subroutine calls itasfe to calculate maximum
    '     stable film thickness iteratively. it uses dimensionless parameters
    '     as input.
    '                              references
    '                              ----------
    '     1. barnea, d., " transition from annular flow and from dispersed
    '        bubble flow - unified models for the whole range of pipe in-
    '        clinations ", int. j. of multiphase flow, vol.12, (1986),
    '        733-744.
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '                   input/output logical file variables
    '                   -----------------------------------
    '     ioerr = output file for error messages when input values passed
    '             to the subroutine are out of range or error occurs in
    '             the calculation.
    '                           subsubroutines called
    '                           ------------------
    '     itsafe = this subroutine iterate safely within the specified
    '              limits of the variable.
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '                           variable description
    '                           --------------------
    '    *deld   = ratio of film thickness to pipe diameter.
    '     e      = liquid holdup fraction at a pipe cross-section.
    '     ec     = liquid holdup for core with respect to pipe cross-
    '              section.
    '     ef    = liquid holdup for film with respect to pipe cross-
    '              section.
    '    *ensc   = no-slip holdup for core with respect to core cross-
    '              section.
    '      fpat  = flow pattern, (chr)
    '                 "anul" = annular
    '                 "slug" = slug
    '     ierr   = error code, (0=ok, 1=input variable out of range.)
    '    *ioerr  = error message file.
    '    *xmo2   = dimensionless group defined in  anmist.
    '    *ym     = dimensionless group defined in  anmist.
    '     (* indicates input variables)
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    Private Sub chkan(xmo2#, ym#, deld#, ensc#, e#, fpat As Double)
        ' sub parameter : do not dim ! Dim fpat as string * 4
        'Dim itsafe As Double
        '     ------------------------------------------------------
        '     calculate total liquid holdup at a pipe cross-section.
        '     ------------------------------------------------------
        Dim ef#, ec#, deldmx#, deldmn#, deldac#, delds#
        ef = 4.0# * deld * (1.0# - deld)
        ec = ensc * (1.0# - 2.0# * deld) ^ 2
        e = ef + ec
        If (e > 0.12) Then
            fpat = "slug" '        annular flow is bridged resulting in slug flow.
        Else
            '        ----------------------------------------------------
            '        calculate maximum stable film thickness using itsafe
            '        for iteration.
            '        ----------------------------------------------------
            deldmx = 0.499
            deldmn = 0.00001
            deldac = 0.00001
            delds = itsafe(xmo2, ym, 0#, 0#, 0#, 0#, 5, deldmn, deldmx, deldac)
            If (delds < deld) Then
                fpat = "slug" '            film is unstable causing slug flow.
            Else
                fpat = "anul" '            annular flow is confirmed by barnea criteria.
            End If
        End If
    End Sub


    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^*
    '     this function carries out iteration within a fixed limits of a
    '     variable.
    '     written by,    asfandiar m. ansari
    '     revised by,    asfandiar m. ansari      last revision: march 1989
    '              * *  tulsa university fluid flow projects  * *
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^*
    '     this function uses newton-raphson technique for iteration. it
    '     keeps the solution within the specified limits by using bisection
    '     method when newton-raphson solution crosses the limits. subroutine
    '     func is called to define function and its derivative needed by
    '     newton-raphson technique. there is no restriction on the system of
    '     units as long as func can incorporate it.
    '                            references
    '                            ----------
    '    1. press, w. h., et al.," numerical recipes",cambridge university
    '       press, new york (1986).
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '                   input/output logical file variables
    '                   -----------------------------------
    '     ioerr = output file for error messages when input values passed
    '             to the subroutine are out of range or error occurs in
    '             the calculation.
    '                           subsubroutines called
    '                           ------------------
    '     func = this defines function and its derivative to be used for
    '            iteration.
    '                          variable description
    '                            --------------------
    '     df    = value of the derivative of f during iteration.
    '     dx    = difference between the two successive guesses.
    '     dxold = previous dx.
    '     f     = value of the function during iteration.
    '     fh    = highest value of the function.
    '     fl    = lowest value of the function.
    '     i     = indicator for variable to be iterated,
    '             + = holdup in  bubble flow.
    '             2 = void fraction in taylor  bubble.
    '             3 = nusselt film thickness around taylor  bubble.
    '             4 = film thickness in annular flow.
    '             5 = stable film thickness for annular film.
    '     ierr  = error code, (0=ok, 1=input variable out of range.)
    '    *ioerr = error message file.
    '     j     = do loop variable.
    '     maxit = iteration counter.
    '     swap  = dummy variable used to swap or interchange fh and fl.
    '     xh    = highest value for the variable.
    '     xl    = lowest value for the variable.
    '    *xacc  = accuracy acceptable for the solution.
    '    *x1    = upper limit for the solution.
    '    *x2    = lower limit for the solution.
    '    *a,b,c,d,e and g are input dummy variables that define the
    '     function.
    '     (* indicates input variables)
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^



    Private Function itsafe(a#, b#, c#, d#, e#, g_#, i As Integer, x1#, x2#, xacc#) As Double
        '     ---------------------------------------------------------
        '     calculate values of the function and its derivative at x1
        '     and x2.
        '     ---------------------------------------------------------
        Dim fl#, df#, fh#, xl#, xh#, swap#
        Dim dxold#, DX#
        Dim f#
        Dim j As Integer, temp As Double

        Call Func(a, b, c, d, e, g_, i, x1, fl, df)
        Call Func(a, b, c, d, e, g_, i, x2, fh, df)
        '     ---------------------------------------------------------
        '     interchange values of the function if it varies inversely
        '     with the variable.
        '     ---------------------------------------------------------
        If (fl < 0#) Then
            xl = x1
            xh = x2
        Else
            xh = x1
            xl = x2
            swap = fl
            fl = fh
            fh = swap
        End If
        '     -------------------------------------------------
        '     take the average of x1 and x2 as the first guess.
        '     -------------------------------------------------
        itsafe = 0.5 * (x1 + x2)
        '     ---------------------------------------------------
        '     define the difference in the two successive values.
        '     ---------------------------------------------------
        dxold = Abs(x2 - x1)
        DX = dxold
        '     ----------------------------------------
        '     call func again to use the guessed value.
        '     ----------------------------------------
        Call Func(a, b, c, d, e, g_, i, itsafe, f, df)
        '     -----------------------------------------------------
        '     carry out iteration by using newton-raphson method
        '     together with bisection approach to keep the variable
        '     within its limits.
        '     -----------------------------------------------------
        For j = 1 To MAXIT
            If (((itsafe - xh) * df - f) * ((itsafe - xl) * df - f) >= 0# Or Abs(2.0# * f) > Abs(dxold * df)) Then
                dxold = DX
                DX = 0.5 * (xh - xl)
                itsafe = xl + DX
                If (xl = itsafe) Then Exit Function
            Else
                dxold = DX
                DX = f / df
                temp = itsafe
                itsafe = itsafe - DX
                If (temp = itsafe) Then Exit Function
            End If
            If (Abs(DX) < xacc Or Abs(f) < xacc) Then Exit Function
            Call Func(a, b, c, d, e, g_, i, itsafe, f, df)
            If (f < 0#) Then
                xl = itsafe
                fl = f
            Else
                xh = itsafe
                fh = f
            End If
            If (f = 0#) Then Exit Function
        Next j
        AddLogMsg(" itsafe: no convergence even after 100 iterations")

    End Function

    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '     this subroutine defines a function and its derivative to be used
    '     for iteration.
    '     written by,    asfandiar m. ansari
    '     revised by,    asfandiar m. ansari     last revision: march 1989
    '              * *  tulsa university fluid flow projects  * *
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^*
    '     this subroutine is called by itsafe to get standard function and
    '     its derivative for newton-raphson iterative technique. the number
    '     arguments for this subroutine are based on the number of variables
    '     involved in the most complex function for which the subroutine is
    '     called. for simpler functions most of the arguments are taken as
    '     zero. the function to be used by itsafe is recognized by indicator
    '     i.
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '                           variable description
    '                           --------------------
    '     df   = derivative of f.
    '     f    = function to be defined for iteration.
    '     i    = indicator to select f.
    '     x    = variable to be iterated.
    '     a,b,c,d,e and f are input variables that define f.
    '     t"s and dt"s are dummy variables and their derivatives.
    '     ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    Private Sub Func(a#, b#, c#, d#, e#, g_#, i As Integer, X#, f#, df#)
        '     ----------------------------------------------------
        '     initialize dummy variables that are repeatedly used.
        '     ----------------------------------------------------
        Dim t1#, dt1#, t2#, dt2#, t3#, dt3#, t4#, dt4#
        t1 = 0#
        dt1 = 0#
        t2 = 0#
        dt2 = 0#
        t3 = 0#
        dt3 = 0#
        t4 = 0#
        dt4 = 0#
        Dim an#, t5#, t6#, t7#, t8#, t9#, t10#, t11#, t12#, t13#, t14#, dt5#, dt6#, dt7#, dt8#, dt9#, dt10#, dt11#, dt12#, dt13#, dt14#, z#, dz#
        If (i = 1) Then
            '        ------------------------------------------------
            '        define f and df to iterate for holdup in bubble.
            '        ------------------------------------------------
            an = 0.5
            f = c * X ^ an + 1.2 * (a + b) - b / (1.0# - X)
            df = an * c * X ^ (an - 1.0#) - b / (1.0# - X) ^ 2
        ElseIf (i = 2) Then
            '        ------------------------------------------------------
            '        define f and df to iterate for void fraction in taylor
            '         bubble in slug.
            '        ------------------------------------------------------
            t2 = Sqrt(9.81 * g_ * (1.0# - Sqrt(X)))
            t3 = 1.0# - X
            t4 = 1.0# - e
            t5 = 9.961 * t2
            t6 = c - (c + t5) * t3 / t4
            t7 = 1.2 * (a + b) + d * t4 ^ 0.5
            If (e > 0.25) Then t7 = t6
            t8 = c * (1.0# - e / X) + t7 * e / X
            t9 = t6 * t4 - a
            t10 = t5 * t3 + t6 * t4
            t11 = b - e * t7
            t12 = X * t8 - e * t7
            t13 = t9 / t10
            t14 = t11 / t12
            f = t13 - t14
            dt2 = -0.25 * 9.81 * g_ / t2 / Sqrt(X)
            dt5 = 9.961 * dt2
            dt6 = -(dt5 * t3 / t4 - (c + t5) / t4)
            dt7 = 0#
            If (e > 0.25) Then dt7 = dt6
            dt8 = c * (e / X ^ 2) + dt7 * e / X - t7 * e / X ^ 2
            dt9 = dt6 * t4
            dt10 = dt5 * t3 - t5 + dt6 * t4
            dt11 = -e * dt7
            dt12 = t8 + X * dt8 - e * dt7
            dt13 = dt9 / t10 - t9 * dt10 / t10 ^ 2
            dt14 = dt11 / t12 - t11 * dt12 / t12 ^ 2
            df = dt13 - dt14
        ElseIf (i = 3) Then
            '        --------------------------------------------------------------
            '        define f and df to iterate for nusselt film thickness in slug.
            '        --------------------------------------------------------------
            t1 = e / (0.425 + 2.65 * (d + e))
            t2 = (1.0# - 2.0# * X / a) ^ 2
            t3 = (b * t2 - (b - c) * t1) / t2
            t4 = (t3 * t2 - (d + e)) / (1.0# - t2)
            f = X ^ 3 - 0.75 * a * g_ * t4 * (1.0# - t2)
            dt2 = -4.0# * (1.0# - 2.0# * X / a) / a
            dt3 = (b - c) * t1 * dt2 / t2 ^ 2
            dt4 = (dt3 * t2 + t3 * dt2) / (1.0# - t2) + (t3 * t2 - (d + e)) * dt2 / (1.0# - t2) ^ 2
            df = 3.0# * X ^ 2 - 0.75 * a * g_ * (dt4 * (1.0# - t2) - t4 * dt2)
        ElseIf (i = 4) Then
            '        --------------------------------------------------------
            '        define f and df to iterate for film thickness in anmist.
            '        --------------------------------------------------------
            t1 = 4.0# * X * (1.0# - X)
            z = 1.0# + c * X
            dt1 = 4.0# * (1.0# - 2.0# * X)
            dz = c
            f = b - z / t1 / (1.0# - t1) ^ 2.5 + a / t1 ^ 3
            df = -dz / t1 / (1.0# - t1) ^ 2.5 - 2.5 * z * dt1 / t1 / (1.0# - t1) ^ 3.5 + z * dt1 / t1 ^ 2 / (1.0# - t1) ^ 2.5 - 3.0# * a / t1 ^ 4 * dt1
        ElseIf (i = 5) Then
            '         -----------------------------------------------------------
            '         define f and df to iterate  stable film thickness in chkan.
            '         -----------------------------------------------------------
            t1 = 1.0# - (1.0# - 2.0# * X) ^ 2
            dt1 = 4.0# * (1.0# - 2.0# * X)
            '         ---------------------------------------------------
            '         to avoid division of f by 0, adjust x if necessary.
            '         ---------------------------------------------------
            t2 = 1.0# / t1
            If (t2 = 1.5) Then X = X + 0.001
            t1 = 1.0# - (1.0# - 2.0# * X) ^ 2
            dt1 = 4.0# * (1.0# - 2.0# * X)
            f = b - (2.0# - 1.5 * t1) * a / t1 ^ 3 / (1.0# - 1.5 * t1)
            df = 1.5 * dt1 * a / t1 ^ 3 / (1.0# - 1.5 * t1) + 3.0# * a * dt1 * (1.0# - 2.0# * t1) _
             * (2.0# - 1.5 * t1) / t1 ^ 4 / (1.0# - 1.5 * t1) ^ 2
        End If
        Exit Sub
        Exit Sub
    End Sub

    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^*
    '     unified model for gas-liquid pipe flow via slug dynamics
    '     written by,          hong-quan (holden) zhang          july 20, 2001
    '     revised by,          hong-quan (holden) zhang          july 30, 2002
    '     revised by,          hong-quan (holden) zhang          oct. 16, 2002
    '     revised by,          hong-quan (holden) zhang          apr. 23, 2003
    '          * *          tulsa university fluid flow projects (tuffp)     * *
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^*
    '     this subroutine is for predictions of flow pattern, pressure gradient,
    '     liquid holdup and slug characteristics in gas-liquid pipe flow at
    '     different inclination angles from -90 to 90 deg.
    '     the main subroutine handles the input and output. calculations are made
    '     in gal (gas and liquid) and its subroutines.
    '                                   references
    '                                   ----------
    '     1. h.-q. zhang, "unified model for gas-liquid pipe flow - model development,"
    '       etce 2002, houston.
    '     2. h.-q. zhang, "unified model for gas-liquid pipe flow - model validation,"
    '       etce 2002, houston.
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^*
    '                              subroutines called
    '                              ------------------
    '
    '     gal          =     overall calculations
    '     dislug     =     calculates the superficial liquid velocity on the
    '                    boundary between dispersed bubble flow and slug flow
    '                    with a given superficial gas velocity
    '     stslug     =     calculates the superficial liquid velocity on the
    '                    boundary between slug flow and stratified (or annular)
    '                    flow with a given superficial gas velocity
    '                    (for horizontal and downward flow)
    '     anslug     =     calculates the superficial gas velocity on the
    '                    boundary between slug flow and annular (or stratified)
    '                    flow with a given superficial liquid velocity
    '                    (for upward flow)
    '     buslug     =     calculates the superficial gas velocity on the
    '                    boundary between slug flow and bubbly flow
    '                    with a given superficial liquid velocity
    '                    (for near vertical upward flow, >60 deg, and large d)
    '     single     =     calculates pressure gradient for single phase flow
    '                    of liquid or gas
    '     buflow     =     calculates pressure gradient and liquid holdup for bubbly
    '                    flow (with bubble rise velocity vo)
    '     dbflow     =     calculates pressure gradient and liquid holdup for dispersed
    '                    bubble flow (without bubble rise velocity)
    '     itflow     =     calculates pressure gradient, liquid holdup and slug
    '                    characteristics for intermittent flow
    '     saflow     =     calculates pressure gradient and liquid holdup for stratified
    '                    or annular flow
    '                              variable description
    '                              --------------------
    '     ac          =     cross sectional area of gas core (ft^2 m^2)
    '     af          =     cross sectional area of film (ft^2 m^2)
    '     *ang          =     angle of flow (pipe) from horizontal (deg.)
    '     axp          =     cross sectional area of pipe (m2)
    '     cf          =     film length in a slug unit (m)
    '     cs          =     slug length (m)
    '     cu          =     slug unit length (m)
    '     *d          =     inside pipe diameter (in. or m)
    '     *deng     =     gas density (lbm/ft3 or kg/m3)
    '     *dengo     =     gas density at atmospheric conditions (lbm/ft3 or kg/m3)
    '     *denl     =     liquid density (lbm/ft3 or kg/m3)
    '     *ea          =     absolute pipe wall roughness (in. or m)
    '     ed          =     relative pipe wall roughness (in. or m)
    '     ens          =     no-slip liquid holdup
    '     fe          =     entrainment fraction in gas core
    '     fec          =     maximum entrainment fraction in gas core
    '     fc          =     friction factor between gas core and pipe wall
    '     ff          =     friction factor between film and pipe wall
    '     fi          =     interfacial friction factor between gas and film
    '     fpt(fpo)=     flow pattern (chr.),               ifpt
    '                    "n-a"     =     not available           0
    '                    "int"     =     intermittent           1
    '                    "str"     =     stratified                2
    '                    "ann"     =     annular                3
    '                    "d-b"     =     dispersed bubble      4
    '                    "bub"     =     bubbly                     5
    '                    "liq"     =     liquid                     6
    '                    "gas"     =     gas                     7
    '     fqn           =     slug freq_Hz (1/s)
    '     fro          =     froude number
    '     *g          =     gravitational acceleration (m/s2)
    '     hl          =     liquid holdup
    '     hlc          =     liquid holdup in gas core
    '     hlf          =     liquid holdup in film
    '     hls          =     liquid holdup in slug body
    '     hlsc     =     maximum liquid holdup in slug body
    '     icon     =     counter of iteration times
    '     *iunit     =     unit indicator, 0 for si and 1 for british
    '     *p          =     pressure (psia or pa)
    '     pga          =     acceleration pressure gradient (psi/ft or pa/m)
    '     pge          =     elevation pressure gradient (psi/ft or pa/m)
    '     pgf          =     friction pressure gradient (psi/ft or pa/m)
    '     pgt          =     total pressure gradient (psi/ft or pa/m)
    '     const_Pi          =     ratio of the circumference of a circle to its diameter
    '     re          =     reynolds number
    '     rsu          =     ratio of slug length to slug unit length
    '     sc          =     perimeter contacted by gas core (ft or m)
    '     sf          =     perimeter wetted by film (ft or m)
    '     si          =     perimeter of interface (ft or m)
    '     sl          =     perimeter of of pipe wetted by liquid (ft or m)
    '     *surl     =     liquid surface tension (lbf/ft or n/m)
    '     *surw     =     water surface tension (lbf/ft or n/m)
    '     thf          =     film thickness (ft or m)
    '     vc          =     gas core velocity (ft/s or m/s)
    '     vf          =     film velocity (ft/s or m/s)
    '     *visg     =     gas viscosity (cp)
    '     *visl     =     liquid viscosity (cp)
    '     vm          =     mixture or slug velocity (ft/s or m/s)
    '     *vsg          =     superficial gas velocity (ft/s or m/s)
    '     *vsl          =     superficial liquid velocity (ft/s or m/s)
    '     vt          =     slug traslational (tail and front) velocity (ft/s or m/s)
    '     we          =     weber number
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^*
    Private Sub zhangmodel(d#, ed#, ang#, vsl#, vsg#, denl#, deng#, visl#, visg#, surl#, p#,
     Hl#, pgt#, pgf#, fpt As String, vf#, hlf#, SL#, FF#, hls#, cu#, fqn#, rsu#, icon#, cs#, cf#, VC#, pgg#, pga#)
        p = p * 101325.0#
        visg = visg * 0.001
        visl = visl * 0.001
        ' mixture velocity
        Dim VM#, E1#, fec#, hlsc#, surw#, dengo#, axp#, ens#
        VM = vsl + vsg
        ' tolerances for iterations
        E1 = 0.0001
        ' limitation for liquid entraiment fraction in gas core
        fec = 0.75
        ' limitation for liquid holdup in slug body
        hlsc = 0.36
        ' constants
        'const_Pi = 3.1415926
        'g = 9.81
        ' surface tension of water against air
        surw = 0.0731
        ' density of air at atmospheric pressure
        dengo = 1.2
        ' cross sectional area of the pipe
        axp = const_Pi * d * d / 4.0#
        '     ---------------------------
        '     check for single phase flow
        '     ---------------------------
        If vsl + vsg > 0 Then
            ens = vsl / (vsl + vsg)
        Else
            ens = 1
        End If

        If (ens >= 0.99999) Then
            fpt = "liq"
            Hl = 1.0#
            SL = d * const_Pi
            Call singlee(d, ed, ang, p, denl, vsl, visl, FF, pgt, pgf, pgg, pga)
            GoTo L60
        ElseIf (ens <= 0.00001) Then
            fpt = "gas"
            Hl = 0#
            SL = 0#
            Call singlee(d, ed, ang, p, deng, vsg, visg, FF, pgt, pgf, pgg, pga)
            GoTo L60
        End If
        '     -------------------------------
        '     check int - d-b transition boundary
        '     -------------------------------
        Dim vdb, vst As Double
        Call dislug(d, ed, ang, vsg, denl, deng, visl, surl, vdb)
        If (vsl > vdb) Then
            fpt = "d-b"
            SL = d * const_Pi
            Call dbflow(d, ed, ang, vsl, vsg, denl, deng, visl, Hl, FF, pgt, pgf, pgg, pga)
            GoTo L60
        End If
        If (ang > 0#) Then GoTo L50
        '     -------------------------------
        '     check i-sa transition boundary for downward flow (mostly i-s)
        '     -------------------------------
        Dim SF#, thf#
        Call stslug(d, ed, ang, vsg, denl, deng, dengo, visl, visg, surl, vst)
        If (vsl < vst) Then
            Call saflow(d, ed, ang, vsl, vsg, denl, deng, dengo, visl, visg, surl,
                        Hl, FF, pgt, pgf, fpt, p, hlf, vf, SF, thf, icon, hls, VC, pgg, pga)
            SL = SF
        Else
            fpt = "int"
            SL = d * const_Pi
            Call itflow(d, ed, ang, vsl, vsg, denl, deng, dengo, visl, visg, surl, Hl, FF, pgt, pgf, fpt, cu, hlf, vf, fqn, rsu, hls, icon,
            VC, cs, cf, pgg, pga)
            If (fpt = "d-b") Then
                SL = d * const_Pi
                Call dbflow(d, ed, ang, vsl, vsg, denl, deng, visl, Hl, FF, pgt, pgf, pgg, pga)
            End If
            If (fpt = "str" Or fpt = "ann") Then
                Call saflow(d, ed, ang, vsl, vsg, denl, deng, dengo, visl, visg, surl,
                Hl, FF, pgt, pgf, fpt, p, hlf, vf, SF, thf, icon, hls, VC, pgg, pga)
                SL = SF
            End If
        End If
        GoTo L60
        '     -------------------------------
        ' check i-sa transition boundary for upward flow (mostly i-a)
        '     -------------------------------
L50:
        'continue
        Dim van As Double
        Call anslug(d, ed, ang, vsl, denl, deng, dengo, visl, visg, surl, van)
        If (vsg > van) Then
            fpt = "ann"
            SL = d * const_Pi
            Call saflow(d, ed, ang, vsl, vsg, denl, deng, dengo, visl, visg,
               surl, Hl, FF, pgt, pgf, fpt, p, hlf, vf, SF, thf, icon, hls, VC, pgg, pga)
            SL = SF
            GoTo L60
        End If
        '     -------------------------------
        ' check i-bu transition boundary
        '     -------------------------------
        Dim ckd As Double
        ckd = (denl * denl * const_g * d / ((denl - deng) * surl)) ^ 0.25
        If (ckd <= 4.37) Then
            fpt = "int"
            SL = d * const_Pi
            Call itflow(d, ed, ang, vsl, vsg, denl, deng, dengo, visl, visg,
              surl, Hl, FF, pgt, pgf, fpt, cu, hlf, vf, fqn, rsu, hls, icon, VC, cs, cf, pgg, pga)
            If (fpt = "d-b") Then
                SL = d * const_Pi
                Call dbflow(d, ed, ang, vsl, vsg, denl, deng, visl, Hl, FF, pgt, pgf, pgg, pga)
            End If
            If (fpt = "str" Or fpt = "ann") Then
                Call saflow(d, ed, ang, vsl, vsg, denl, deng, dengo, visl, visg,
                      surl, Hl, FF, pgt, pgf, fpt, p, hlf, vf, SF, thf, icon, hls, VC, pgg, pga)
                SL = SF
            End If
            GoTo L60
        End If
        Dim vbu As Double
        Call buslug(d, vsl, vbu)
        If (vsg < vbu And ang > 60.0#) Then
            fpt = "bub"
            SL = d * const_Pi
            Call buflow(d, ed, ang, vsl, vsg, denl, deng, visl, surl, Hl, FF, pgt, pgf, pgg, pga)
        Else
            fpt = "int"
            SL = d * const_Pi
            Call itflow(d, ed, ang, vsl, vsg, denl, deng, dengo, visl, visg,
                 surl, Hl, FF, pgt, pgf, fpt, cu, hlf, vf, fqn, rsu, hls, icon, VC, cs, cf, pgg, pga)
            If (fpt = "d-b") Then
                SL = d * const_Pi
                Call dbflow(d, ed, ang, vsl, vsg, denl, deng, visl, Hl, FF, pgt, pgf, pgg, pga)
            End If
            If (fpt = "str" Or fpt = "ann") Then
                Call saflow(d, ed, ang, vsl, vsg, denl, deng, dengo, visl, visg,
                        surl, Hl, FF, pgt, pgf, fpt, p, hlf, vf, SF, thf, icon, hls, VC, pgg, pga)
                SL = SF
            End If
        End If
L60:
        'continue
        '     --------------------------------------
        '     change variables back to british units
        '     --------------------------------------
        pgf = pgf / 101325.0#
        pgg = pgg / 101325.0#
        pga = pga / 101325.0#
        pgt = pgt / 101325.0#
        p = p / 101325.0#
        visg = visg * 1000.0#
        visl = visl * 1000.0#
L900:
    End Sub

    Private Sub stslug(ByVal d As Double,
                       ByVal ed As Double,
                       ByVal ang As Double,
                       ByVal vsg As Double,
                       ByVal denl As Double,
                       ByVal deng As Double,
                       ByVal dengo As Double,
                       ByVal visl As Double,
                       ByVal visg As Double,
                       ByVal surl As Double,
                       ByVal vst As Double)
        ' tolerances for iterations
        Dim E1, fec, hlsc, surw, axp, cs, CC, an1, FI, VM As Double
        Dim g, vdb As Double
        E1 = 0.0001
        ' limitation for liquid entraiment fraction in gas core
        fec = 0.75
        ' limitation for liquid holdup in slug body
        hlsc = 0.36
        ' constants
        'const_Pi = 3.1415926
        g = 9.81
        ' surface tension of water against air
        surw = 0.0731
        ' density of air at atmospheric pressure
        dengo = 1.2
        ' cross sectional area of the pipe
        axp = const_Pi * d * d / 4.0#
        cs = (32.0# * cosd(ang) ^ 2 + 16.0# * sind(ang) ^ 2) * d
        CC = 1.25 - 0.5 * Abs(sind(ang))
        an1 = const_Pi * 0.5
        FI = 0.0142
        Call dislug(d, ed, ang, vsg, denl, deng, visl, surl, vdb)
        ' guess a vst
        vst = 0.5
        VM = vst + vsg
        Dim hls, feo, hlfo, hlso, icon As Double
        hls = 1.0# / (1.0# + (VM / 8.66) ^ 1.39)
        If (hls < hlsc) Then hls = hlsc
        feo = 0#
        hlfo = vst / VM
        hlso = hls
        icon = 0
L5:
        If (vst > vdb) Then GoTo L90
        ' entrainment fraction according to oliemans et al."s (1986) correlation
        Dim resg, web, fro, resl, ccc, fe, VT, hlf, hlc, af, ac As Double
        resg = Abs(deng * vsg * d / visg)
        web = Abs(deng * vsg * vsg * d / surl)
        fro = Abs(Sqrt(g * d) / vsg)
        resl = Abs(denl * vst * d / visl)
        ccc = 0.003 * web ^ 1.8 * fro ^ 0.92 * resl ^ 0.7 * (denl / deng) ^ 0.38 * (visl / visg) ^ 0.97 / resg ^ 1.24
        fe = ccc / (1.0# + ccc)
        If (fe > fec) Then fe = fec
        fe = (fe + 9.0# * feo) / 10.0#
        feo = fe
        ' translational velocity according to nicklin (1962), bendiksen (1984)
        ' and zhang et al. (2000)
        VT = 1.3 * VM + (0.54 * cosd(ang) + 0.35 * sind(ang)) * Sqrt(g * d * (denl - deng) / denl)
        hlf = ((hls * (VT - VM) + vst) * (vsg + vst * fe) - VT * vst * fe) / (VT * vsg)
        If (hlf <= 0#) Then hlf = Abs(hlf)
        If (hlf >= 1.0#) Then hlf = 1.0# / hlf
        hlf = (hlf + 9.0# * hlfo) / 10.0#
        hlfo = hlf
        hlc = (1.0# - hlf) * vst * fe / (VM - vst * (1.0# - fe))
        If (hlc < 0#) Then hlc = 0#
        af = hlf * axp
        ac = (1.0# - hlf) * axp
        ' wetted angle
        ' calculate wetted angle using newton"s method
        Dim an2, tha, an, th0, Th As Double
        If (af < axp) Then
L10:
            an2 = an1 - 0.5 * (8.0# * af / d / d + Sin(an1) - an1) / (Cos(an1) - 1.0#)
            If (an2 > 2.0# * const_Pi Or an2 < 0#) Then an2 = 1.5 * const_Pi
            tha = Abs((an2 - an1) / an1)
            If (tha > E1) Then
                an1 = an2
                GoTo L10
            Else
                an = an2
            End If
        Else
            an = 2.0# * const_Pi
            af = axp
        End If
        ' wetted wall fraction according to grolman et al., aiche (1996)
        If (Abs(ang) < 85.0#) Then
            th0 = an / (2.0# * const_Pi)
            Th = th0 * (surw / surl) ^ 0.15 + deng * (denl * vst * vst * d / surl) ^ 0.25 _
        * (vsg * vsg / ((1.0# - hlf) ^ 2 * g * d)) ^ 0.8 / ((denl - deng) * cosd(ang))
            If (Th > 1.0#) Then Th = 1.0#
        Else
            Th = 1.0#
        End If
        ' wetted perimeters
        Dim SF, sc, Ab, si, df, thf, DC, vf, VC, denc, ref, rec, FF, fc As Double
        SF = const_Pi * d * Th
        sc = const_Pi * d * (1.0# - Th)
        Ab = d * d * (const_Pi * Th - Sin(2.0# * Th * const_Pi) / 2.0#) / 4.0#
        si = (SF * (Ab - af) + d * Sin(const_Pi * Th) * af) / Ab
        ' the hydraulic diameters
        df = 4.0# * af / SF
        thf = 2.0# * af / (SF + si)
        DC = 4.0# * ac / (sc + si)
        vf = vst * (1.0# - fe) / hlf
        VC = (VM - vst * (1.0# - fe)) / (1.0# - hlf)
        ' reynolds numbers
        denc = (denl * hlc + deng * (1.0# - hlf - hlc)) / (1.0# - hlf)
        ref = Abs(denl * vf * df / visl)
        rec = Abs(deng * VC * DC / visg)
        ' friction factors
        FF = 0.046 / (ref) ^ 0.2
        fc = 0.046 / (rec) ^ 0.2
        '      ff=1.0/(3.6*log10(6.9/ref+(ed/3.7)^1.11))^2
        '      fc=1.0/(3.6*log10(6.9/rec+(ed/3.7)^1.11))^2
        ' interfacial friction factor
        ' stratified flow interfacial friction factor
        ' interfacial friction factor according to andritsos et al. (1987)
        ' modified by zhang (2001)
        Dim vsgt, abcd, vfn, abu, dpex, remx, FM, dpsl, ad As Double

        vsgt = 5.0# * Sqrt(dengo / deng)
        FI = fc * (1.0# + 15.0# * (2.0# * thf / d) ^ 0.5 * (vsg / vsgt - 1.0#))
        If (FI < fc) Then FI = fc
        abcd = (sc * fc * deng * VC * Abs(VC) / (2.0# * ac) + si * FI * deng * (VC - vf) * Abs(VC - vf) * (1.0# / af + 1.0# / ac) / 2.0# _
         - (denl - denc) * g * sind(ang)) * af * 2.0# / (SF * FF * denl)
        If (abcd < 0#) Then
            vfn = vf * 0.9
            GoTo L20
        Else
        End If
        vfn = Sqrt(abcd)
L20:
        abu = Abs((vfn - vf) / vf)
        If (abu > E1) Then
            vf = (vfn + 9.0# * vf) / 10.0#
            vst = vf * hlf / (1.0# - fe)
            VM = vst + vsg
            dpex = (denl * (VM - vf) * (VT - vf) * hlf + denc * (VM - VC) * (VT - VC) * (1.0# - hlf)) * d / cs / 4.0#
            '          denm=denl*hls+deng*(1.0-hls)
            remx = Abs(d * VM * denl / visl)
            FM = 1.0# / (3.6 * Log10(6.9 / remx + (ed / 3.7) ^ 1.11)) ^ 2
            dpsl = FM * denl * VM * VM / 2.0#
            ad = (dpsl + dpex) / (3.16 * CC * Sqrt(surl * (denl - deng) * g))
            hls = 1.0# / (1.0# + ad)
            If (hls < hlsc) Then hls = hlsc
            hls = (hls + 9.0# * hlso) / 10.0#
            hlso = hls
            icon = icon + 1
            GoTo L5
        Else
        End If
        vst = vfn * hlf / (1.0# - fe)
L90:
        If (vst > vdb) Then vst = vdb
        Exit Sub
    End Sub

    Private Sub anslug(ByVal d As Double,
                       ByVal ed As Double,
                       ByVal ang As Double,
                       ByVal vsl As Double,
                       ByVal denl As Double,
                       ByVal deng As Double,
                       ByVal dengo As Double,
                       ByVal visl As Double,
                       ByVal visg As Double,
                       ByVal surl As Double,
                       ByVal van As Double)

        Dim E1, fec, hlsc, g, surw, axp, cs, CC, an1, FI, VM, hls, hlso, feo, hlfo, web, fro As Double ' vdb
        Dim resl, resg, ccc, fe, VT, hlf, hlc, af, ac, an2, tha, an, th0, Th As Double
        Dim SF, sc, Ab, si, df, thf, DC, vf, VC, denc, ref, rec, fr2, fr3, FF, fc, vsgt As Double
        Dim abcd, vcn, abu, dpex, remx, FM, dpsl, ad As Double

        ' tolerances for iterations
        E1 = 0.0001
        ' limitation for liquid entraiment fraction in gas core
        fec = 0.75
        ' limitation for liquid holdup in slug body
        hlsc = 0.36
        ' constants
        'const_Pi = 3.1415926
        g = 9.81
        ' surface tension of water against air
        surw = 0.0731
        ' density of air at atmospheric pressure
        dengo = 1.2
        ' cross sectional area of the pipe
        axp = const_Pi * d * d / 4.0#
        cs = (32.0# * cosd(ang) ^ 2 + 16.0# * sind(ang) ^ 2) * d
        CC = 1.25 - 0.5 * Abs(sind(ang))
        an1 = const_Pi * 0.5
        FI = 0.0142
        ' guess a van
        van = 10.0#
        VM = vsl + van
        hls = 1.0# / (1.0# + (VM / 8.66) ^ 1.39)
        If (hls < hlsc) Then hls = hlsc
        hlso = hls
        feo = 0#
        hlfo = vsl / VM
        ' entrainment fraction according to oliemans et al"s (1986) correlation
L105:
        web = deng * van * van * d / surl
        fro = Sqrt(g * d) / van
        resl = denl * vsl * d / visl
        resg = deng * van * d / visg
        ccc = 0.003 * web ^ 1.8 * fro ^ 0.92 * resl ^ 0.7 * (denl / deng) ^ 0.38 * (visl / visg) ^ 0.97 / resg ^ 1.24
        fe = ccc / (1.0# + ccc)
        If (fe > fec) Then fe = fec
        fe = (fe + 9.0# * feo) / 10.0#
        feo = fe
        ' translational velocity according to nicklin (1962), bendiksen (1984)
        ' and zhang et al. (2000)
        VT = 1.3 * VM + (0.54 * cosd(ang) + 0.35 * sind(ang)) * Sqrt(g * d * (denl - deng) / denl)
        hlf = ((hls * (VT - VM) + vsl) * (van + vsl * fe) - VT * vsl * fe) / (VT * van)
        If (hlf <= 0#) Then hlf = Abs(hlf)
        If (hlf >= 1.0#) Then hlf = 1.0# / hlf
        hlf = (hlf + 9.0# * hlfo) / 10.0#
        hlfo = hlf
        hlc = (1.0# - hlf) * vsl * fe / (VM - vsl * (1.0# - fe))
        If (hlc < 0#) Then hlc = 0#
        af = hlf * axp
        ac = (1.0# - hlf) * axp
        ' wetted angle
L110:
        an2 = an1 - 0.5 * (8.0# * af / d / d + Sin(an1) - an1) / (Cos(an1) - 1.0#)
        If (an2 > 2.0# * const_Pi) Then an2 = 1.75 * const_Pi
        If (an2 < 0#) Then an2 = 0.25 * const_Pi
        tha = Abs((an2 - an1) / an1)
        If (tha > E1) Then
            an1 = an2
            GoTo L110
        Else
            an = an2
        End If
        ' wetted wall fraction according to grolman and fortuin (1996)
        If (Abs(ang) < 85.0#) Then
            th0 = an / (2.0# * const_Pi)
            Th = th0 * (surw / surl) ^ 0.15 + deng * (denl * vsl * vsl * d / surl) ^ 0.25 _
             * (van * van / ((1.0# - hlf) ^ 2 * g * d)) ^ 0.8 / ((denl - deng) * cosd(ang))
            If (Th > 1.0#) Then Th = 1.0#
        Else
            Th = 1.0#
        End If
        ' wetted perimeters
        SF = const_Pi * d * Th
        sc = const_Pi * d * (1.0# - Th)
        Ab = d * d * (const_Pi * Th - Sin(2.0# * Th * const_Pi) / 2.0#) / 4.0#
        si = (SF * (Ab - af) + d * Sin(const_Pi * Th) * af) / Ab
        ' the hydraulic diameters
        df = 4.0# * af / SF
        thf = 2.0# * af / (SF + si)
        DC = 4.0# * ac / (sc + si)
        vf = vsl * (1.0# - fe) / hlf
        VC = (VM - vsl * (1.0# - fe)) / (1.0# - hlf)
        ' reynolds numbers
        denc = (denl * hlc + deng * (1.0# - hlf - hlc)) / (1.0# - hlf)
        ref = Abs(denl * vf * df / visl)
        rec = Abs(deng * VC * DC / visg)
        ' frictional factors
        fr2 = 16.0# / 2000.0#
        fr3 = 1.0# / (3.6 * Log10(6.9 / 3000.0# + (ed / 3.7) ^ 1.11)) ^ 2
        '      fr3=0.046/(3000.0)^0.2
        If (ref < 2000.0#) Then FF = 16.0# / ref
        If (ref > 3000.0#) Then FF = 1.0# / (3.6 * Log10(6.9 / ref + (ed / 3.7) ^ 1.11)) ^ 2
        '      if(ref > 3000.0) ff=0.046/(ref)^0.2
        If (ref >= 2000.0# And ref <= 3000.0#) Then FF = fr2 + (fr3 - fr2) * (ref - 2000.0#) / 1000.0#
        '      ff=1.0/(3.6*log10(6.9/ref+(ed/3.7)^1.11))^2
        fc = 1.0# / (3.6 * Log10(6.9 / rec + (ed / 3.7) ^ 1.11)) ^ 2
        '      ff=0.046/(ref)^0.2
        '      fc=0.046/(rec)^0.2
        ' interfacial friction factor according to andritsos et al. (1987)
        ' modified by zhang (2001)
        vsgt = 5.0# * Sqrt(dengo / deng)
        FI = fc * (1.0# + 15.0# * (2.0# * thf / d) ^ 0.5 * (van / vsgt - 1.0#))
        If (FI < fc) Then FI = fc
        abcd = (SF * FF * denl * vf * Abs(vf) / (2.0# * af) - sc * fc * deng * VC * VC / (2.0# * ac) _
         + (denl - denc) * g * sind(ang)) * 2.0# / (si * FI * deng * (1.0# / af + 1.0# / ac))
        If (abcd < 0#) Then
            vcn = VC * 0.9
            GoTo L120
        Else
        End If
        vcn = Sqrt(abcd) + vf
L120:
        abu = Abs((vcn - VC) / VC)
        If (abu > E1) Then
            VC = (vcn + 9.0# * VC) / 10.0#
            van = VC * (1.0# - hlf) - vsl * fe
            VM = vsl + van
            dpex = (denl * (VM - vf) * (VT - vf) * hlf + denc * (VM - VC) * (VT - VC) * (1.0# - hlf)) * d / cs / 4.0#
            '          denm=denl*hls+deng*(1.0-hls)
            remx = Abs(d * VM * denl / visl)
            FM = 1.0# / (3.6 * Log10(6.9 / remx + (ed / 3.7) ^ 1.11)) ^ 2
            dpsl = FM * denl * VM * VM / 2.0#
            ad = (dpsl + dpex) / (3.16 * CC * Sqrt(surl * (denl - deng) * g))
            hls = 1.0# / (1.0# + ad)
            If (hls < hlsc) Then hls = hlsc
            hls = (hls + 9.0# * hlso) / 10.0#
            hlso = hls
            GoTo L105
        Else
        End If
        van = vcn * (1.0# - hlf) - vsl * fe
        Exit Sub
    End Sub
    ' boundary between slug and dispersed bubble flows
    Private Sub dislug(ByVal d As Double,
                       ByVal ed As Double,
                       ByVal ang As Double,
                       ByVal vsg As Double,
                       ByVal denl As Double,
                       ByVal deng As Double,
                       ByVal visl As Double,
                       ByVal surl As Double,
                       ByVal vdb As Double)

        Dim E1, fec, hlsc, g, surw, axp, CC, VM, DC As Double ' an1, FI, hls, hlso, feo, hlfo, web, fro, cs
        'Dim resl, resg, ccc, fe, VT, hlf, hlc, af, ac, an2, tha, an, th0, Th
        'Dim SF, sc, Ab, si, df, thf, vf, VC, denc, ref, rec, fr2, fr3, FF, fc, vsgt
        Dim FM, dengo, vdb1, hlb, denm, rem1, vmn, abm, vdb2 As Double ' abcd, vcn, abu, dpex, remx, dpsl, ad,
        ' tolerances for iterations
        E1 = 0.0001
        ' limitation for liquid entraiment fraction in gas core
        fec = 0.75
        ' limitation for liquid holdup in slug body
        hlsc = 0.36
        ' constants
        'const_Pi = 3.1415926
        g = 9.81
        ' surface tension of water against air
        surw = 0.0731
        ' density of air at atmospheric pressure
        dengo = 1.2
        ' cross sectional area of the pipe
        axp = const_Pi * d * d / 4.0#
        CC = 1.25 - 0.5 * Abs(sind(ang))
        ' guess a vdb
        vdb1 = 2.0#
        Dim i As Integer
        i = 0
L10:
        i = i + 1
        VM = vdb1 + vsg
        hlb = vdb1 / VM
        denm = (1.0# - hlb) * deng + hlb * denl
        rem1 = Abs(denm * d * VM / visl)
        FM = 1.0# / (3.6 * Log10(6.9 / rem1 + (ed / 3.7) ^ 1.11)) ^ 2
        'Open report.txt For Append As #1
        '      addLogMsg "ppp2",hlb,vdb1,vm,denm,rem1,fm,d,visl
        '      fm=0.046/abs(rem1)^0.2
        '      dc=2.0* sqr (0.4*surl/((denl-deng)*g))
        '      en=fm*denm*vm^2*hlb
        vmn = Sqrt(((1.0# / hlb - 1.0#) * 6.32 * CC * Sqrt((denl - deng) * g * surl)) / (FM * denm))
        abm = Abs((vmn - VM) / VM)
        If (abm > E1) And i < 100 Then
            VM = (vmn + 19.0# * VM) / 20.0#
            vdb1 = VM - vsg
            GoTo L10
        Else
        End If
        vdb1 = vmn - vsg
        vdb2 = 2.0#
        Dim ve, DV, En, vn, VD As Double
        ve = vdb2
L30:
        hlb = ve / (vsg + ve)
        VM = vsg + ve
        denm = (1.0# - hlb) * deng + hlb * denl
        rem1 = Abs(denm * d * VM / visl)
        FM = 1.0# / (3.6 * Log10(6.9 / rem1 + (ed / 3.7) ^ 1.11)) ^ 2
        '      fm=0.046/abs(rem1)^0.2
        DC = 2.0# * Sqrt(0.4 * surl / ((denl - deng) * g))
        DV = (denl / surl) ^ 0.6 * (2.0# * FM / d) ^ 0.4
        En = 4.15 * Sqrt(1.0# - hlb) + 0.725
        vn = (En / (DC * DV)) ^ 0.83 - vsg
        VD = Abs((vn - ve) / ve)
        If (VD > E1) Then
            ve = (vn + 9.0# * ve) / 10.0#
            GoTo L30
        Else
        End If
L40:
        vdb2 = vn
        If (vdb2 > vdb1) Then
            vdb = vdb2
        Else
            vdb = vdb1
        End If
    End Sub
    Private Sub buslug(ByVal d As Double,
                       ByVal vsl As Double,
                       ByVal vbu As Double)
        ' constants
        'const_Pi = 3.1415926
        Dim g, axp, vo, hgc As Double
        g = 9.81
        ' cross sectional area of the pipe
        axp = const_Pi * d * d / 4.0#
        vo = 0.35 * Sqrt(g * d)
        hgc = 0.25
        vbu = vsl * hgc / (1.0# - hgc) + vo * hgc
    End Sub

    '     single phase flow calculation
    Private Sub singlee(d#, ed#, ang#, p#, den#, v#, vis#, FF#, pgt#, pgf#, pge#, pga#)

        Dim E1#, fec#, hlsc#, surw#, dengo#, axp#, Re#, ekk# 'fr2#, fr3#,

        E1 = 0.0001 ' tolerances for iterations
        fec = 0.75 ' limitation for liquid entraiment fraction in gas core
        hlsc = 0.36 'limitation for liquid holdup in slug body
        surw = 0.0731 ' surface tension of water against air
        dengo = const_rho_air ' 1.2 ' density of air at atmospheric pressure
        axp = const_Pi * d * d / 4.0# ' cross sectional area of the pipe
        pge = -den * sind(ang) * const_g '     calculate elevation pressure gradient.
        Re = Abs(d * den * v / vis)
        FF = unf_friction_factor(Re, ed, 4) / 4
        pgf = -2.0# * FF * den * v * v / d '     calculate frictional pressure gradient.
        ekk = den * v * v / p
        If (ekk > 0.95) Then ekk = 0.95
        pgt = (pge + pgf) / (1.0# - ekk)
        pga = pgt * ekk '     calculate accelerational pressure gradient.
        pgt = (pge + pgf + pga)
        If (den > 400.0#) Then
            pgt = (pge + pgf)
            pga = 0#
        End If
    End Sub

    Private Sub dbflow(d#, ed#, ang#, vsl#, vsg#, denl#, deng#, visl#, Hl#, FM#, pgt#, pgf#, pge#, pga#)

        Dim E1#, fec#, hlsc#, surw#, dengo#, axp#, VM#, denm#, dens#, rem1# 'ekk#, icrit,

        ' tolerances for iterations
        E1 = 0.0001
        ' limitation for liquid entraiment fraction in gas core
        fec = 0.75
        ' limitation for liquid holdup in slug body
        hlsc = 0.36
        ' surface tension of water against air
        surw = 0.0731
        ' density of air at atmospheric pressure
        dengo = 1.2
        ' cross sectional area of the pipe
        axp = const_Pi * d * d / 4.0#
        VM = vsg + vsl
        '     calculate liquid holdup
        Hl = vsl / (vsg + vsl)
        '     calculate pressure gradients
        denm = (1.0# - Hl) * deng + Hl * denl
        dens = denm + Hl * (denl - denm) / 3.0#
        rem1 = Abs(dens * d * VM / visl)
        FM = unf_friction_factor(rem1, ed, 4) / 4
        pgf = -2.0# * FM * dens * VM ^ 2 / d
        pge = -const_g * denm * sind(ang)
        pga = 0#
        pgt = (pgf + pge + pga)
    End Sub

    Private Sub buflow(d#, ed#, ang#, vsl#, vsg#, denl#, deng#, visl#, surl#, Hl#, FM#, pgt#, pgf#, pge#, pga#)

        Dim E1#, fec#, hlsc#, surw#, dengo#, axp#, VM#, vo#, denm#, dens#, rem1# ' Re#, fr2#, fr3#, ekk#, icrit#,
        ' tolerances for iterations
        E1 = 0.0001
        ' limitation for liquid entraiment fraction in gas core
        fec = 0.75
        ' limitation for liquid holdup in slug body
        hlsc = 0.36
        ' surface tension of water against air
        surw = 0.0731
        ' density of air at atmospheric pressure
        dengo = 1.2
        ' cross sectional area of the pipe
        axp = const_Pi * d * d / 4.0#
        VM = vsg + vsl
        vo = 1.53 * (const_g * (denl - deng) * surl / denl / denl) ^ 0.25 * sind(ang)
        '     calculate liquid holdup
        If (Abs(ang) < 10.0#) Then
            Hl = vsl / (vsg + vsl)
        Else
            Hl = (Sqrt((VM - vo) ^ 2 + 4.0# * vsl * vo) - vsg - vsl + vo) / (2.0# * vo)
        End If
        '     calculate pressure gradients
        denm = (1.0# - Hl) * deng + Hl * denl
        dens = denm + Hl * (denl - denm) / 3.0#
        rem1 = Abs(dens * d * VM / visl)
        FM = unf_friction_factor(rem1, ed, 4) / 4
        pgf = -2.0# * FM * dens * VM ^ 2 / d
        pge = -const_g * denm * sind(ang)
        pga = 0#
        pgt = (pgf + pge + pga)
    End Sub

    ' intermittent flow calculation
    Private Sub itflow(d#, ed#, ang#, vsl#, vsg#, denl#, deng#, dengo#, visl#, visg#,
     surl#, Hl#, FF#, pgt#, pgf#, fpt As String, cu#, hlf#, vf#, fqn#, rsu#, hls#, icon#,
     VC#, cs#, cf#, pgg#, pga#)

        Dim VM#, hlsc#, surw#, axp#, hlso#, VT#, CC#, an1#, E1# 'eq#
        icon = 0
        VM = vsl + vsg
        E1 = 0.0001 ' tolerances for iterations
        hlsc = 0.36 ' limitation for liquid holdup in slug body
        surw = 0.0731 ' surface tension of water against air
        dengo = 1.2 ' density of air at atmospheric pressure
        axp = const_Pi * d * d / 4.0# ' cross sectional area of the pipe
        hls = 1.0# / (1.0# + (VM / 8.66) ^ 1.39)
        If (hls < hlsc) Then hls = hlsc
        hlso = hls
        ' translational velocity according to nicklin (1962), bendiksen (1984)
        ' and zhang et al. (2000)
        VT = 1.3 * VM + (0.54 * cosd(ang) + 0.35 * sind(ang)) * Sqrt(const_g * d * (denl - deng) / denl)
        ' slug length
        cs = (32.0# * cosd(ang) ^ 2 + 16.0# * sind(ang) ^ 2) * d
        CC = 1.25 - 0.5 * Abs(sind(ang))
        ' guess cu and cf
        cu = cs * VM / vsl
        cf = cu - cs
        hlf = vsl / VM
        ' frictional factors based on superfacial velocities and pipe diameter
        ' assuming the flow state is turbulent
        an1 = const_Pi * 0.7
L5:
        cu = cf + cs
        Dim vfn, hlfn, af, ac, dpex, rem1, FM, dpsl, ad As Double
        vfn = (cu * vsl - cs * hls * VM) * VT / (cf * VT * hls + cu * (vsl - hls * VM))
        vf = (vfn + 9.0# * vf) / 10.0#
        '      vf=vfn
        hlfn = (cf * VT * hls + cu * (vsl - hls * VM)) / (cf * VT)
        If (hlfn <= 0#) Then
            hlfn = vsl / VM 'Abs(hlfn)
        End If
        If (hlfn >= 1.0#) Then
            hlfn = vsl / VM '1# - 1# / hlfn
        End If
        hlf = (hlfn + 4.0# * hlf) / 5.0#
        '      hlf=hlfn
        VC = (VM - hlf * vf) / (1.0# - hlf)
        af = hlf * axp
        ac = (1.0# - hlf) * axp
        ' slug liquid holdup
        dpex = (denl * (VM - vf) * (VT - vf) * hlf + deng * (VM - VC) * (VT - VC) * (1.0# - hlf)) * d / cs / 4.0#
        '       denm=denl*hls+deng*(1.0-hls)
        rem1 = Abs(denl * VM * d / visl)
        FM = 1.0# / (3.6 * Log10(6.9 / rem1 + (ed / 3.7) ^ 1.11)) ^ 2
        dpsl = FM * denl * VM * VM / 2.0#
        ad = (dpsl + dpex) / (3.16 * CC * Sqrt(surl * (denl - deng) * const_g))
        hls = 1.0# / (1.0# + ad)
        If (hls < hlsc) Then hls = hlsc
        hls = (hls + 9.0# * hlso) / 10.0#
        hlso = hls
        ' wetted angle assuming flat film surface
        ' calculated using newton"s method
        Dim an2, tha, an, vsgf, vslf, th0, Th As Double
        If (af < axp) Then
L10:
            an2 = an1 - 0.5 * (8.0# * af / d / d + Sin(an1) - an1) / (Cos(an1) - 1.0#)
            If (an2 > 2.0# * const_Pi) Then an2 = 1.75 * const_Pi
            If (an2 < 0#) Then an2 = 0.25 * const_Pi
            tha = Abs((an2 - an1) / an1)
            If (tha > E1) Then
                an1 = an2
                GoTo L10
            Else
                an = an2
            End If
        Else
            an = 2.0# * const_Pi
            af = axp
        End If
        ' wetted wall fraction according to grolman et al., aiche (1996)
        vsgf = VC * (1.0# - hlf)
        vslf = vf * hlf
        If (Abs(ang) < 85.0#) Then
            th0 = an / (2.0# * const_Pi)
            Th = th0 * (surw / surl) ^ 0.15 + deng * (denl * vslf * vslf * d / surl) ^ 0.25 _
                 * (vsgf * vsgf / ((1.0# - hlf) ^ 2 * const_g * d)) ^ 0.8 / ((denl - deng) * cosd(ang))
            If (Th > 1.0#) Then Th = 1.0#
        Else
            Th = 1.0#
        End If
        ' wetted perimeters
        Dim SF, sc, Ab, si, df, thf, DC, ref, rec, fr2, fr3, fc, vsgt As Double
        Dim abcd, abu, denm As Double ', vmn, abm, vdb2, remx, vcn, vdb1, hlb, denc

        SF = const_Pi * d * Th
        sc = const_Pi * d * (1.0# - Th)
        Ab = d * d * (const_Pi * Th - Sin(2.0# * Th * const_Pi) / 2.0#) / 4.0#
        si = (SF * (Ab - af) + d * Sin(const_Pi * Th) * af) / Ab
        ' the hydraulic diameters
        df = 4.0# * af / SF
        thf = 2.0# * af / (SF + si)
        DC = 4.0# * ac / (sc + si)
        ' reynolds numbers
        ref = Abs(denl * vf * df / visl)
        rec = Abs(deng * VC * DC / visg)
        ' frictional factors
        fr2 = 16.0# / 2000.0#
        fr3 = 1.0# / (3.6 * Log10(6.9 / 3000.0# + (ed / 3.7) ^ 1.11)) ^ 2
        '      fr3=0.046/(3000.0)^0.2
        If (ref < 2000.0#) Then FF = 16.0# / ref
        If (ref > 3000.0#) Then FF = 1.0# / (3.6 * Log10(6.9 / ref + (ed / 3.7) ^ 1.11)) ^ 2
        '      if(ref > 3000.0) ff=0.046/(ref)^0.2
        If (ref >= 2000.0# And ref <= 3000.0#) Then FF = fr2 + (fr3 - fr2) * (ref - 2000.0#) / 1000.0#
        If (rec < 2000.0#) Then fc = 16.0# / rec
        If (rec > 3000.0#) Then fc = 1.0# / (-3.6 * Log10(6.9 / rec + (ed / 3.7) ^ 1.11)) ^ 2
        '      if(rec > 3000.0) fc=0.046/(rec)^0.2
        If (rec >= 2000.0# And rec <= 3000.0#) Then fc = fr2 + (fr3 - fr2) * (rec - 2000.0#) / 1000.0#
        ' interfacial friction factor according to andritsos et al. (1987)
        ' modified by zhang (2001)
        vsgt = 5.0# * Sqrt(dengo / deng)
        Dim FI, fsl As Double
        FI = fc * (1.0# + 15.0# * (2.0# * thf / d) ^ 0.5 * (vsgf / vsgt - 1.0#))
        '      fi=fc*(1.0+14.3*abs(hlf)^0.5*(vsgf/vsgt-1.0))
        ' interfacial friction factor (annular) according to ambrosini et al. (1991)
        '       reg=abs(vc*deng*d/visg)
        '       wed=deng*vc*vc*d/surl
        '       fs=0.046/reg^0.2
        '      shi=fi*deng*(vc-vf)^2/2.0
        '      thfo=thf* sqr (shi*deng)/visg
        '      fra=fs*(1.0+13.8*(thfo-200.0* sqr (deng/denl))*wed^0.2/reg^0.6)
        '      fi=(frs+fra)/2.0
        If (FI < fc) Then FI = fc
        If (FI > 1.0#) Then FI = 1.0#
        '      if(fi < 0.0142) fi=0.0142
        ' calculate film length cf using the combined momentum eqaution
        fsl = denl * (VM - vf) * (VT - vf) - deng * (VM - VC) * (VT - VC)
        abcd = SF * FF * denl * vf * Abs(vf) / 2.0# / af _
             - sc * fc * deng * VC * Abs(VC) / 2.0# / ac _
             - si * FI * deng * (VC - vf) * Abs(VC - vf) / 2.0# * (1.0# / af + 1.0# / ac) _
             + (denl - deng) * const_g * sind(ang)
        Dim abef, cfn, cfm, cfo, dens As Double
        abef = 3.0# * SF * visl * vf / thf / af _
             - SF * FI * deng * (VC - vf) * Abs(VC - vf) / 4.0# / af _
             - sc * fc * deng * VC * Abs(VC) / 2.0# / ac _
             - si * FI * deng * (VC - vf) * Abs(VC - vf) / 2.0# * (1.0# / af + 1.0# / ac) _
             + (denl - deng) * const_g * sind(ang)
        cfn = fsl / abcd
        cfm = fsl / abef
        icon = icon + 1
        If (ref > 3000.0#) Then cfo = cfn
        If (ref < 2000.0#) Then cfo = cfm
        If (ref <= 3000.0# And ref >= 2000.0#) Then cfo = (cfn * (ref - 2000.0#) + cfm * (3000.0# - ref)) / 1000.0#
        If (cfo < 0#) Then
            cfo = 0.5 * cs * (VM / vsl - 1) '-cfo
        End If
        abu = Abs((cfo - cf) / cf)
        If (abu < E1 Or icon > 100) Then GoTo L100
        cf = (cfo + 9.0# * cf) / 10.0#
        GoTo L5
L100:
        cf = cfo
        ' slug unit length
        cu = cf + cs
        '      vf=(cu*vsl-cs*hls*vm)*vt
        '     & /(cf*vt*hls+cu*(vsl-hls*vm))
        '      hlf=(cf*vt*hls+cu*(vsl-hls*vm))/(cf*vt)
        '      vc=(vm-hlf*vf)/(1.0-hlf)
        denm = denl * hls + deng * (1.0# - hls)
        dens = denm + hls * (denl - denm) / 3.0#
        Dim res, fs, fos, dps, fof, dpf As Double
        res = Abs(dens * VM * d / visl)
        fr2 = 16.0# / 2000.0#
        fr3 = 1.0# / (3.6 * Log10(6.9 / 3000.0# + (ed / 3.7) ^ 1.11)) ^ 2
        '      fr3=0.046/(3000.0)^0.2
        If (res < 2000.0#) Then fs = 16.0# / res
        If (res > 3000.0#) Then fs = 1.0# / (3.6 * Log10(6.9 / res + (ed / 3.7) ^ 1.11)) ^ 2
        '      if(res > 3000.0) fs=0.046/(res)^0.2
        If (res >= 2000.0# And res <= 3000.0#) Then fs = fr2 + (fr3 - fr2) * (res - 2000.0#) / 1000.0#
        ' slug freq_Hz
        fqn = VT / cu
        ' slug to slug unit length ratio
        rsu = cs / cu
        ' pressure gradient in slug
        fos = (denl * (VM - vf) * (VT - vf) * hlf + deng * (VM - VC) * (VT - VC) * (1.0# - hlf)) / cs
        dps = -fs * dens * VM * VM * 2.0# / d - denm * const_g * sind(ang) - fos
        ' pressure gradient in film
        fof = fos * cs / cf
        dpf = fof - SF * FF * denl * vf * Abs(vf) / (2.0# * axp) _
         - sc * fc * deng * VC * Abs(VC) / (2.0# * axp) _
         - (denl * hlf + deng * (1.0# - hlf)) * const_g * sind(ang)
        ' total pressure gradient
        pgt = (dps * cs + dpf * cf) / cu
        ' total pressure gradient due to friction
        pgf = -((fs * dens * VM * VM * 2.0# / d) * cs _
         + (SF * FF * denl * vf * Abs(vf) / (2.0# * axp) _
         + sc * fc * deng * VC * Abs(VC) / (2.0# * axp)) * cf) / cu
        pgg = pgt - pgf
        pga = 0
        ' overall liquid holdup
        Hl = (cf * hlf + hls * cs) / cu
        If (Abs(cf) < d) Then '
            If (res < 2000) Then
                fpt = "bub"
            Else
                fpt = "d-b"
            End If
        End If
        'If (Abs(cf) < d) Then fpt = "d-b"
        If (Abs(cf) > 199.0# * cs Or hlf > hls) Then
            fpt = "ann"
            If (vf < 0#) Then vf = vsl / hlf
        End If
L280:
        Exit Sub
    End Sub

    ' stratified and/or annular flow calculation
    Private Sub saflow(d#, ed#, ang#, vsl#, vsg#, denl#, deng#, dengo#, visl#, visg#,
                        surl#, Hl#, FF#, pgt#, pgf#, fpt As String, p#, hlf#, vf#, SF#, thf#, icon#, hlc#, VC#, pgg#, pga#)

        Dim E1, fec, hlsc, g, surw, axp, web, fro, an1, FI As Double ', vdb, VM, hls, hlso, feo, hlfo, cs, CC,
        Dim resl, resg, ccc, fe, af, ac, an2, tha, an, th0, Th As Double 'VT,
        Dim sc, Ab, si, df, DC, denc, ref, rec, fr2, fr3, fc, vsgt As Double
        Dim abcd, abu As Double ', vcn, abu, dpex, remx, FM, dpsl, ad
        'Dim fpt As String * 3
        icon = 0
        ' tolerances for iterations
        E1 = 0.0001
        ' limitation for liquid entraiment fraction in gas core
        fec = 0.75
        ' limitation for liquid holdup in slug body
        hlsc = 0.36
        ' constants
        'const_Pi = 3.1415926
        g = 9.81
        ' surface tension of water against air
        surw = 0.0731
        ' density of air at atmospheric pressure
        dengo = 1.2
        ' cross sectional area of the pipe
        axp = const_Pi * d * d / 4.0#
        FI = 0.0142
        ' entrainment fraction according to oliemans et al"s (1986) correlation
        resg = Abs(deng * vsg * d / visg)
        web = Abs(deng * vsg * vsg * d / surl)
        fro = Abs(Sqrt(g * d) / vsg)
        resl = Abs(denl * vsl * d / visl)
        ccc = 0.003 * web ^ 1.8 * fro ^ 0.92 * resl ^ 0.7 _
         * (denl / deng) ^ 0.38 * (visl / visg) ^ 0.97 / resg ^ 1.24
        fe = ccc / (1.0# + ccc)
        If (fe > 1.0#) Then fe = 1.0#
        ' guess a film velocity
        vf = vsl
        hlf = vsl / (vsl + vsg)
        an1 = 0.7 * const_Pi
L5:
        Dim hlfn As Double
        hlfn = vsl * (1.0# - fe) / vf
        If (hlfn <= 0#) Then hlfn = Abs(hlfn)
        If (hlfn >= 1.0#) Then hlfn = 1.0# - 1.0# / hlfn
        hlf = (hlfn + 4.0# * hlf) / 5.0#
        '      hlf=hlfn
        VC = (vsg + fe * vsl) / (1.0# - hlf)
        af = hlf * axp
        ac = (1.0# - hlf) * axp
        hlc = vsl * fe / VC
        If (hlc < 0#) Then hlc = 0#
        denc = (denl * hlc + deng * (1.0# - hlf - hlc)) / (1.0# - hlf)
        ' wetted angle assuming flat film surface
        ' calculated using newton"s method
        If (af < axp) Then
L10:
            an2 = an1 - 0.5 * (8.0# * af / d / d + Sin(an1) - an1) / (Cos(an1) - 1.0#)
            If (an2 > 2.0# * const_Pi) Then an2 = 1.75 * const_Pi
            If (an2 < 0#) Then an2 = 0.25 * const_Pi
            tha = Abs((an2 - an1) / an1)
            If (tha > E1) Then
                an1 = an2
                GoTo L10
            Else
                an = an2
            End If
        Else
            an = 2.0# * const_Pi
            af = axp
        End If
        ' wetted wall fraction according to grolman et al., aiche (1996)
        If (Abs(ang) < 85.0#) Then
            th0 = an / (2.0# * const_Pi)
            Th = th0 * (surw / surl) ^ 0.15 + deng * (denl * vsl * vsl * d / surl) ^ 0.25 _
                 * (vsg * vsg / ((1.0# - hlf) ^ 2 * g * d)) ^ 0.8 _
                 / ((denl - deng) * cosd(ang))
            If (Th > 1.0#) Then Th = 1.0#
        Else
            Th = 1.0#
        End If
        If (Th > 0.9) Then
            fpt = "ann"
        Else
            fpt = "str"
        End If
        ' wetted perimeters
        SF = const_Pi * d * Th
        sc = const_Pi * d * (1.0# - Th)
        Ab = d * d * (const_Pi * Th - Sin(2.0# * Th * const_Pi) / 2.0#) / 4.0#
        si = (SF * (Ab - af) + d * Sin(const_Pi * Th) * af) / Ab
        ' the hydraulic diameters
        df = 4.0# * af / SF
        thf = 2.0# * af / (SF + si)
        DC = 4.0# * ac / (sc + si)
        ' reynolds numbers
        ref = Abs(denl * vf * df / visl)
        rec = Abs(deng * VC * DC / visg)
        ' friction factors
        fr2 = 16.0# / 2000.0#
        fr3 = 1.0# / (3.6 * Log10(6.9 / 3000.0# + (ed / 3.7) ^ 1.11)) ^ 2
        '      fr3=0.046/(3000.0)^0.2
        If (ref < 2000.0#) Then FF = 16.0# / ref
        If (ref > 3000.0#) Then FF = 1.0# / (3.6 * Log10(6.9 / ref + (ed / 3.7) ^ 1.11)) ^ 2
        '      if(ref > 3000.0) ff=0.046/(ref)^0.2
        If (ref >= 2000.0# And ref <= 3000.0#) Then FF = fr2 + (fr3 - fr2) * (ref - 2000.0#) / 1000.0#
        If (rec < 2000.0#) Then fc = 16.0# / rec
        If (rec > 3000.0#) Then fc = 1.0# / (3.6 * Log10(6.9 / rec + (ed / 3.7) ^ 1.11)) ^ 2
        '      if(rec > 3000.0) fc=0.046/(rec)^0.2
        If (rec >= 2000.0# And rec <= 3000.0#) Then fc = fr2 + (fr3 - fr2) * (rec - 2000.0#) / 1000.0#
        ' interfacial friction factor (stratified) according to andritsos et al. (1987)
        ' modified by zhang (2001)
        vsgt = 5.0# * Sqrt(dengo / deng)
        Dim frs As Double
        frs = fc * (1.0# + 15.0# * (2.0# * thf / d) ^ 0.5 * (vsg / vsgt - 1.0#))
        '      frs=fc*(1.0+14.3*abs(hlf)^0.5*(vsg/vsgt-1.0))
        '      if(frs < fc) frs=fc
        ' interfacial friction factor (annular) according to ambrosini et al. (1991)
        Dim reg, wed, shi, thfo, fra, fs, vfn, vfm, vfo As Double
        reg = Abs(VC * deng * d / visg)
        wed = deng * VC * VC * d / surl
        fs = 0.046 / reg ^ 0.2
        shi = Abs(FI * deng * (VC - vf) ^ 2 / 2.0#)
        thfo = thf * Sqrt(shi * deng) / visg
        fra = fs * (1.0# + 13.8 * (thfo - 200.0# * Sqrt(deng / denl)) * wed ^ 0.2 / reg ^ 0.6)
        If (rec > 10000.0#) Then
            FI = (10000.0# * frs / rec + fra) / (1.0# + 10000.0# / rec)
        Else
            FI = (fra + frs) / 2.0#
        End If
        If (fpt = "ann") Then FI = fra
        If (FI < fc) Then FI = fc
        If (FI > 1.0#) Then FI = 1.0#
        ' iterations
        abcd = (sc * fc * deng * VC * Abs(VC) / (2.0# * ac) + si * FI * deng * (VC - vf) * Abs(VC - vf) * (1.0# / af + 1.0# / ac) / 2.0# _
             - (denl - denc) * g * sind(ang)) * af * 2.0# / (SF * FF * denl)
        If (abcd > 0#) Then
            vfn = Sqrt(abcd)
        Else
            vfn = -Sqrt(-abcd)
        End If
        vfm = (sc * fc * deng * VC * Abs(VC) / (2.0# * ac) _
             + si * FI * deng * (VC - vf) * Abs(VC - vf) * (1.0# / af + 1.0# / ac) / 2.0# _
              - (denl - denc) * g * sind(ang) _
             + SF * FI * deng * (VC - vf) * Abs(VC - vf) / (4.0# * af)) * af * thf / (3.0# * SF * visl)
        icon = icon + 1
        If (ref > 3000.0#) Then vfo = vfn
        If (ref < 2000.0#) Then vfo = vfm
        If (ref <= 3000.0# And ref >= 2000.0#) Then vfo = (vfn * (ref - 2000.0#) + vfm * (3000.0# - ref)) / 1000.0#
L20:
        abu = Abs((vfo - vf) / vf)
        If (abu < E1 Or icon > 10000) Then GoTo L100
        vf = (vfo + 9.0# * vf) / 10.0#
        GoTo L5
L100:
        vf = vfo
        hlfn = vsl * (1.0# - fe) / vf
        If (hlfn <= 0#) Then hlfn = Abs(hlfn)
        If (hlfn >= 1.0#) Then hlfn = 1.0# - 1.0# / hlfn
        hlf = (hlfn + 4.0# * hlf) / 5.0#
        VC = (vsg + fe * vsl) / (1.0# - hlf)
        af = hlf * axp
        ac = (1.0# - hlf) * axp
        hlc = vsl * fe / VC
        If (hlc < 0#) Then hlc = 0#
        ' frictional factors
        denc = (denl * hlc + deng * (1.0# - hlf - hlc)) / (1.0# - hlf)
        ' pressure gradient
        Dim dpf As Double
        If (fpt = "ann") Then
            dpf = -si * FI * deng * (VC - vf) * Abs(VC - vf) / (2.0# * ac) - denc * g * sind(ang)
        Else
            dpf = -SF * FF * denl * vf * Abs(vf) / (2.0# * axp) _
                 - sc * fc * deng * VC * Abs(VC) / (2.0# * axp) _
                 - (denl * hlf + denc * (1.0# - hlf)) * g * sind(ang)
        End If
        ' total pressure gradient
        pgt = dpf / (1.0# - deng * VC * vsg / (p * (1.0# - hlf)))
        ' total pressure gradient due to friction
        pgf = dpf + (denl * hlf + denc * (1.0# - hlf)) * g * sind(ang)
        pgg = -(denl * hlf + denc * (1.0# - hlf)) * g * sind(ang)
        pga = pgt - pgf - pgg
        ' liquid holdup
        Hl = hlf + hlc
        Exit Sub
    End Sub
End Module

'=======================================================================================
'Unifloc 7.25  coronav                                          khabibullin.ra@gubkin.ru
'Petroleum engineering calculations modules (macroses)
'2000 - 2019
'
'=======================================================================================
' class factory functions
' модуль для функций поддержки работы классов
' содержит функции для того чтобы можно было создавать экземпляры классов из других файлов

Option Explicit On

Public Module u7_factory_class
    ' функция генерация трубы из стандартного набора данных
    ' включает и параметры потока в трубе
    ' нужна для упрощения пользовательских функций
    Public Function new_pipeline_with_stream(
                     ByVal qliq_sm3day As Double,
                     ByVal fw_perc As Double,
                     ByVal h_list_m(,) As Double,
                     ByVal t_calc_from_C As Double,
                     ByVal calc_flow_direction As Integer,
                     ByVal str_PVT As String,
                     ByVal diam_list_mm(,) As Double,
            Optional ByVal hydr_corr As H_CORRELATION = 0,
            Optional ByVal t_val(,) As Double = Nothing,
            Optional ByVal temp_method As TEMP_CALC_METHOD = TEMP_CALC_METHOD.StartEndTemp,
            Optional ByVal c_calibr() As Double = Nothing,
            Optional ByVal roughness_m As Double = 0.0001,
            Optional ByVal q_gas_sm3day As Double = 0,
            Optional ByVal znlf As Boolean = False) As CPipe

        Dim pipe As New CPipe
        Dim PVT As New CPVT
        'Dim PTcalc As PTtype
        'Dim TM As TEMP_CALC_METHOD
        ' Dim angle As Double
        Dim tr As New CPipeTrajectory
        Dim amb As New CAmbientFormation
        Dim temp_crv As New CInterpolation
        'Dim prm As PARAMCALC
        Dim c_calibr_grav As Double
        Dim c_calibr_fric As Double


        Try
            Call tr.init_from_vert_range(h_list_m, diam_list_mm)

            Call pipe.init_pipe_constr_by_trajectory(tr)

            PVT = PVT_decode_string(str_PVT) ' initialize PVT properties

            PVT.qliq_sm3day = qliq_sm3day ' set liquid rate and watercut
            PVT.Fw_perc = fw_perc
            PVT.q_gas_free_sm3day = q_gas_sm3day
            pipe.fluid = PVT

            pipe.param = Set_calc_flow_param(calc_along_coord:=calc_flow_direction \ 10 = 1,
                                         flow_along_coord:=calc_flow_direction Mod 10 = 1,
                                         hcor:=hydr_corr,
                                         temp_method:=TEMP_CALC_METHOD.StartEndTemp)
            If znlf Then
                Call pipe.set_ZNLF()
            End If

            Call pipe.InitT(t_calc_from_C, t_val, calc_flow_direction, temp_method)


            ' set calibration properties
            If c_calibr Is Nothing Then
                c_calibr = {1}
            End If
            c_calibr_grav = c_calibr(0)
            If c_calibr.GetUpperBound(0) >= 1 Then
                c_calibr_fric = c_calibr(1)
            Else
                c_calibr_fric = 1
            End If
            pipe.c_calibr_grav = c_calibr_grav
            pipe.c_calibr_fric = c_calibr_fric

            new_pipeline_with_stream = pipe

            Exit Function
        Catch ex As Exception
            'new_pipeline_with_stream = {-1, "error"}
            Dim errmsg As String
            errmsg = "Error:new_pipeline_with_stream:"
            Throw New ApplicationException(errmsg)
        End Try

    End Function

    ' функция генерация трубы из стандартного набора данных
    ' включает и параметры потока в трубе
    ' нужна для упрощения пользовательских функций
    Public Function new_pipe_with_stream(
                ByVal qliq_sm3day As Double,
                ByVal fw_perc As Double,
                ByVal length_m As Double,
                ByVal calc_flow_direction As Integer,
                Optional ByVal str_PVT As String = PVT_DEFAULT,
                Optional ByVal theta_deg As Double = 90,
                Optional ByVal d_mm As Double = 60,
                Optional ByVal hydr_corr As H_CORRELATION = 0,
                Optional ByVal t_calc_from_C As Double = 50,
                Optional ByVal t_calc_to_C As Double = -1,
                Optional ByVal c_calibr() As Double = Nothing,
                Optional ByVal roughness_m As Double = 0.0001,
                Optional ByVal q_gas_sm3day As Double = 0
                             ) As CPipe

        Dim pipe As New CPipe
        Dim PVT As New CPVT
        'Dim PTcalc As PTtype
        'Dim PTin As PTtype
        'Dim PTout As PTtype
        'Dim TM As TEMP_CALC_METHOD
        'Dim out, out_desc
        'Dim out_curves_type As CALC_RESULTS
        'Dim res
        Dim c_calibr_grav As Double
        Dim c_calibr_fric As Double

        ' initialize stream properties
        PVT = PVT_decode_string(str_PVT)    ' create atream object from given string
        PVT.qliq_sm3day = qliq_sm3day           ' set liquid rate to stream
        PVT.Fw_perc = fw_perc                   ' set watercut - fraction of water in stream
        PVT.q_gas_free_sm3day = q_gas_sm3day    ' set gas rate if given. additional gas to main stream
        pipe.fluid = PVT                    ' assign stream to pipe

        ' initialize pipe geometry
        Call pipe.init_pipe(d_mm, length_m, theta_deg, roughness_m)
        ' Pcalc and Tcalc position depends on calc_along_flow
        pipe.param = Set_calc_flow_param(calc_along_coord:=calc_flow_direction \ 10 = 1,
                                         flow_along_coord:=calc_flow_direction Mod 10 = 1,
                                         hcor:=hydr_corr,
                                         temp_method:=TEMP_CALC_METHOD.StartEndTemp)

        If pipe.fluid.qliq_sm3day <= const_ZNLF_rate Then
            Call pipe.set_ZNLF()
        End If


        ' check temp distribution. if second temp not given - set uniform
        If t_calc_to_C < 0 Then t_calc_to_C = t_calc_from_C
        ' temperature initialisation depend on calc direction because initialisation procedure depends on coord direction
        ' check flow direction

        pipe.InitTlinearSmart(t_calc_from_C, t_calc_to_C, calc_flow_direction)

        If c_calibr Is Nothing Then
            c_calibr = {1}
        End If
        ' set calibration properties
        c_calibr_grav = c_calibr(0)
        If c_calibr.GetUpperBound(0) >= 1 Then
            c_calibr_fric = c_calibr(1)
        Else
            c_calibr_fric = 1
        End If
        pipe.c_calibr_grav = c_calibr_grav
        pipe.c_calibr_fric = c_calibr_fric

        new_pipe_with_stream = pipe

    End Function
End Module

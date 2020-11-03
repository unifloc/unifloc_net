Public Class ESP_dict
    Public Property ID As String
    Public Property source As String
    Public Property manufacturer As String
    Public Property name As String
    Public Property stages_max As Integer
    Public Property rate_nom_sm3day As Double
    Public Property rate_opt_min_sm3day As Double
    Public Property rate_opt_max_sm3day As Double
    Public Property rate_max_sm3day As Double
    Public Property slip_nom_rpm As Double
    Public Property freq_Hz As Double
    Public Property eff_max As Double
    Public Property height_stage_m As Double
    Public Property Series As String
    Public Property d_od_mm As Double
    Public Property d_cas_min_mm As Double
    Public Property d_shaft_mm As Double
    Public Property area_shaft_mm As Double
    Public Property power_limit_shaft_kW As Double
    Public Property power_limit_shaft_high_kW As Double
    Public Property power_limit_shaft_max_kW As Double
    Public Property pressure_limit_housing_atma As Double
    Public Property d_motor_od_mm As Double
    Public Property rate_points As Double()
    Public Property head_points As Double()
    Public Property power_points As Double()
    Public Property eff_points As Double()


    Public Property num_stages As Integer
    Public Property head_nom_m As Double
    Public Property gas_correct As Double
    Public Property c_calibr_head As Double
    Public Property c_calibr_rate As Double
    Public Property c_calibr_power As Double
    Public Property dnum_stages_integrate As Integer

End Class
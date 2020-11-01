Public Class ESP_dict
    Public Property ID As String
    Public Property source As String
    Public Property manufacturer As String
    Public Property name As String
    Public Property stages_max As Integer
    Public Property rate_nom_sm3day As Integer
    Public Property rate_opt_min_sm3day As Integer
    Public Property rate_opt_max_sm3day As Integer
    Public Property rate_max_sm3day As Integer
    Public Property slip_nom_rpm As Integer
    Public Property freq_Hz As Integer
    Public Property eff_max As Double
    Public Property height_stage_m As Double
    Public Property Series As Integer
    Public Property d_od_mm As Integer
    Public Property d_cas_min_mm As Integer
    Public Property d_shaft_mm As Integer
    Public Property area_shaft_mm As Integer
    Public Property power_limit_shaft_kW As Integer
    Public Property power_limit_shaft_high_kW As Integer
    Public Property power_limit_shaft_max_kW As Integer
    Public Property pressure_limit_housing_atma As Integer
    Public Property d_motor_od_mm As Integer
    Public Property rate_points() As Integer()
    Public Property head_points() As Double()
    Public Property power_points() As Double()
    Public Property eff_points() As Double()


    Public Property num_stages As Integer
    Public Property head_nom_m As Double
    Public Property gas_correct As Double
    Public Property c_calibr_head As Double
    Public Property c_calibr_rate As Double
    Public Property c_calibr_power As Double
    Public Property dnum_stages_integrate As Integer
End Class
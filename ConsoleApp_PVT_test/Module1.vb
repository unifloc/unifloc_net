Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module Module1

    Sub Main()
        'Dim fluid As New UnfClassLibrary.CPVT With {
        '.gamma_g = 0.6,
        '.gamma_o = 0.86,
        '.Rsb_m3m3 = 100
        '}

        'fluid.Calc_PVT(4, 23)

        'Console.WriteLine("fluid Pb" + CStr(fluid.Pb_calc_atma))
        'Console.ReadKey(True)


        'Dim test2 As JObject = JObject.Parse(pvt_str)
        'Dim myitem = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(pvt_str)
        'Console.WriteLine(myitem.Item("gamma_oil"))
        'Console.ReadKey(True)

        'Dim test_pvt
        'test_pvt = u7_excel.u7_Excel_function_servise.PVT_decode_string(pvt_str)
        'Dim pvt_str, rhg
        'pvt_str = u7_excel.PVT_encode_string(gamma_gas:=0.8, gamma_oil:=0.86, gamma_wat:=1.1, rsb_m3m3:=80, rp_m3m3:=80, pb_atma:=125, t_res_C:=100, bob_m3m3:=1.2, muob_cP:=1)
        'rhg = u7_excel.PVT_rho_gas_kgm3(p_atma:=260, t_C:=80, gamma_gas:=0.8, gamma_oil:=0.86, gamma_wat:=1.1, rsb_m3m3:=80, rp_m3m3:=80, pb_atma:=125, bob_m3m3:=1.2, muob_cP:=1, t_res_C:=100)
        'Console.WriteLine("test: " + CStr(rhg))
        'Console.ReadKey(True)

        Dim test_esp
        'Dim test As String
        test_esp = u7_excel.u7_Excel_functions_ESP.ESP_head_m(qliq_m3day:=30, num_stages:=100, freq_Hz:=50, pump_id:=750, mu_cSt:=0.5)
        'test = JsonConvert.SerializeObject(test_choke)
        Console.WriteLine("test: " + CStr(test_esp))
        Console.ReadKey(True)
    End Sub
End Module
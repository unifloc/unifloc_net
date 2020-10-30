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
        Dim pvt_str, bg, bg_str
        pvt_str = u7_excel.PVT_encode_string(gamma_gas:=0.8, gamma_oil:=0.86, gamma_wat:=1.1, rsb_m3m3:=80, rp_m3m3:=80, pb_atma:=125, t_res_C:=100, bob_m3m3:=1.2, muob_cP:=1)
        bg = u7_excel.PVT_bg_m3m3(p_atma:=260, t_C:=80, gamma_gas:=0.8, gamma_oil:=0.86, gamma_wat:=1.1, rsb_m3m3:=80, rp_m3m3:=80, pb_atma:=125, bob_m3m3:=1.2, muob_cP:=1, tres_C:=100)
        bg_str = JsonConvert.SerializeObject(bg)
        Console.WriteLine(bg_str)

        Dim test_choke
        Dim test As String
        test_choke = u7_excel.u7_Excel_functions_MF.MF_calibr_choke_fast(qliq_sm3day:=50, fw_perc:=20, d_choke_mm:=15, p_in_atma:=60, p_out_atma:=50, str_PVT:=pvt_str)
        test = JsonConvert.SerializeObject(test_choke)
        Console.WriteLine("test: " + test)
        Console.ReadKey(True)
    End Sub
End Module

Imports Newtonsoft.Json

Module Module1

    Sub Main()
        Dim fluid As New UnfClassLibrary.CPVT With {
            .gamma_g = 0.6,
            .gamma_o = 0.86,
            .Rsb_m3m3 = 100
        }

        fluid.Calc_PVT(4, 23)

        Console.WriteLine("fluid Pb" + CStr(fluid.Pb_calc_atma))
        Console.ReadKey(True)

        Dim pvt_str
        pvt_str = u7_excel.PVT_encode_string(gamma_gas:=0.5, gamma_oil:=0.86, gamma_wat:=1)
        Console.WriteLine(pvt_str)
        Console.ReadKey(True)

        Dim test_run
        test_run = JsonConvert.DeserializeObject(pvt_str)
        pvt_str = JsonConvert.SerializeObject(test_run)
        Console.WriteLine(test_run)
        Console.WriteLine(pvt_str)
        Console.ReadKey(True)


        Dim test_next
        test_next = u7_excel.PVT_decode_string(pvt_str)
    End Sub

End Module

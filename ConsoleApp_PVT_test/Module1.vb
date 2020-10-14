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

        Dim new_pvt_str As UnfClassLibrary.CPVT

    End Sub

End Module

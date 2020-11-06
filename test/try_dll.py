import sys 
import clr
import DLLs
from ctypes import *
import os
path = os.path.abspath('..')
path = path + '\\u7_excel\\bin\\Debug'
print(path)
#from System.Collections import *


#import System

#new_lib = CDLL(r'C:\Users\isaev\Downloads\unifloc_net-dev_3\unifloc_net-dev_3\u7_excel\bin\Debug\u7_excel.dll')
path = r'C:\Users\isaev\Downloads\unifloc_net-dev_3\unifloc_net-dev_3\u7_excel\bin\Debug'
sys.path.append(f'{path}')
clr.AddReference('alglibnet2')
clr.AddReference('UnfClassLibrary')
clr.AddReference('u7_excel')

import UnfClassLibrary 
import u7_excel

#x = u7_excel.u7PVT.HelloDna(5.8)
#print(x)

# y = u7_excel.u7PVT.PVT_bg_m3m3(2,23)
# print(f'Вызов 1 - {y}')

# def test_function () :
#     fluid = UnfClassLibrary.CPVT 
#     fluid.gamma_g = 0.6
#     fluid.gamma_o = 0.86
#     fluid.Rsb_m3m3 = 100

#     fluid.Calc_PVT(4.0,23.0)

#     print(f'fluid Pb = {fluid.Pb_calc_atma}')


#from System import Double

#s = String.Overloads[Char, Int32]('A', 10)

#arg = (POINTER(c_double) * 1)()
# fluid = UnfClassLibrary.CPVT()
#
# fluid.gamma_g = 0.6
# fluid.gamma_o = 0.86
# fluid.Rsb_m3m3 = 100
#
# fluid.Calc_PVT(4.0,23.0)
# print(fluid.Pb_calc_atma())

PVT_str = u7_excel.u7_Excel_function_servise.PVT_encode_string(gamma_gas = 0.8, 
                                                               gamma_oil = 0.86, 
                                                               gamma_wat = 1.1, 
                                                               rsb_m3m3= 80, 
                                                               rp_m3m3 = 80, 
                                                               pb_atma = 125, 
                                                               t_res_C = 100, 
                                                               bob_m3m3 = 1.2, 
                                                               muob_cP = 1)
print(PVT_str)
#PVT_str = u7_excel.u7_Excel_function_servise.PVT_encode_string(0.5, 0.78, 98)
c_calibr = u7_excel.u7_Excel_functions_MF.MF_calibr_choke_fast(qliq_sm3day = 50, 
                                                         fw_perc = 20, 
                                                         d_choke_mm = 5, 
                                                         p_in_atma = 60,
                                                         p_out_atma = 50,
                                                         str_PVT = PVT_str)
print(c_calibr[0])


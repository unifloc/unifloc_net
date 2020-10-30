import sys 
import clr
import DLLs
from ctypes import *
import os
path = os.path.abspath('..')
path = path + '\\u7_excel\\bin\\Debug'
print(path)
#import System


#import System

#new_lib = CDLL(r'D:\isaev_downloads\unifloc_net-rnt\unifloc_net-rnt\u7_excel\bin\Debug.UnfClassLibrary.dll')
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

print(u7_excel.u7_Excel_function_servise.PVT_encode_string(0.5, 0.78, 98))


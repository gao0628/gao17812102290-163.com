import pandas as pd
import numpy as np
import geatpy as ea
import xlrd
'''需求：根据停车场的车辆数，动态控制信号灯，一次放行12辆车进入停车场'''
import win32com.client as com  # VISSIM COM



Vissim = com.Dispatch("Vissim.Vissim")
dir = "E:\\vissim标定程序\\1.inpx"
Vissim.LoadNet(dir)
# Define Simulation Configurations
# Vissim.Graphics.CurrentNetworkWindow.SetAttValue("QuickMode", 1)## 设为 不可见 提高效率

Net = Vissim.Net
Sim = Vissim.Simulation
park_num = 12##停车场数量

for i in range(1,61):
    '''大于12，24，36变红，否则变绿，61为车辆总数除以停车场数量'''
    for j in range (1,36000):
        '''36000仿真时长'''
        Sim.RunSingleStep()##单步仿真
        Vehs1 = Net.DataCollectionMeasurements.ItemByKey(1).AttValue('Vehs(Current,1,All)')##提取检测器1的车辆数
        Vehs2 = Net.DataCollectionMeasurements.ItemByKey(2).AttValue('Vehs(Current,1,All)')##提取检测器2的车辆数

        if Vehs1 >= 12*i:##如果大于12的倍数，则变红
            Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(1).SetAttValue('SigState',1)##显示红灯
            # t1 = Vehs1
            if Vehs2 >= 12*i:##如果检测器1大于12，且检测器2大于12，则变为12
                Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(1).SetAttValue('SigState', 3)##显示绿灯
                break##跳出本次循环
        else:
            Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(1).SetAttValue('SigState', 3)##显示绿灯





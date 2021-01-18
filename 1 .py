import win32com.client as com  # VISSIM COM



Vissim = com.Dispatch("Vissim.Vissim")
dir = "D:\\北京城建院\\2020年度\\01 星火仿真\\07 星火仿真\\04 仿真\\20200423\\原方案\\仿真-11.5修改\\仿真\\-11.5.inpx"
Vissim.LoadNet(dir)
# Define Simulation Configurations
# Vissim.Graphics.CurrentNetworkWindow.SetAttValue("QuickMode", 1)## 设为 不可见 提高效率

Net = Vissim.Net
Sim = Vissim.Simulation

#volume = Net.VehicleInput.ItemByKey(1).AttValue('Volume(0)')
for i in range(1,3600):
    Sim.RunSingleStep()
    Vehs1 = Net.DataCollectionMeasurements.ItemByKey(1).AttValue('Vehs(Current,1,All)')##
    Vehs2 = Net.DataCollectionMeasurements.ItemByKey(2).AttValue('Vehs(Current,1,All)')
    Vehs3 = Net.DataCollectionMeasurements.ItemByKey(3).AttValue('Vehs(Current,1,All)')
    Vehs4 = Net.DataCollectionMeasurements.ItemByKey(4).AttValue('Vehs(Current,1,All)')
    
    if Vehs1%10 == 0:
        Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(1).SetAttValue('SigState',1) #红灯
        # t1 = Vehs1
        if Vehs2%10 ==0:
            Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(1).SetAttValue('SigState', 3)
    else:
        Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(1).SetAttValue('SigState', 3)
        
    if Vehs3%10 == 0:
        Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(2).SetAttValue('SigState',1)
                    # t1 = Vehs1
        if Vehs4%10 == 0:
            Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(2).SetAttValue('SigState', 3)         
    else:       
        Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(2).SetAttValue('SigState', 3)
    
    
    

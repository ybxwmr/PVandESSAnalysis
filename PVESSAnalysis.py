import os
import math
import pandas as pd
from PyQt5.QtWidgets import *
import sys
import matplotlib.pyplot as plt

plt.rcParams['font.sans-serif']=['KaiTi','SimHei']#汉字字体
plt.rcParams['font.size']= 10 #字体大小
plt.rcParams['axes.unicode_minus']=False#正常显示负号


def PVandESSAnalysis1(df1,df2):
    df_new = df1.copy()
    ongrid = df2.copy()
    try:
        ration = input('请输入光伏功率折减比例: ')
        df_new['光伏功率(kW)'] = df_new['光伏功率(kW)']* float(ration)
        df_new['上网功率(kW)'] = df_new['光伏功率(kW)'] - df_new['负荷功率(kW)']
        dischargetime = input('请输入上午时段储能充放电时段(8,10表示8点到10点段): ')
        for row in range(len(df_new)):
            if df_new.values[row][2] > 0:
                ongrid.loc[ongrid.index == str(df_new.index[row]).split(' ')[0],'上网电量(kWh)'] = ongrid.loc[ongrid.index == str(df_new.index[row]).split(' ')[0],'上网电量(kWh)'] \
                                                                                    + df_new.values[row][2]

            if df_new.index[row].hour >= int(dischargetime.split(',')[0]) and df_new.index[row].hour < int(dischargetime.split(',')[1]) and df_new.values[row][2] < 0:
                ongrid.loc[ongrid.index == str(df_new.index[row]).split(' ')[0],'上午可放电电量(kWh)'] = ongrid.loc[ongrid.index == str(df_new.index[row]).split(' ')[0],'上午可放电电量(kWh)'] \
                                                                                    - df_new.values[row][2]
        ongrid['上网电量-上午可放电电量'] = ongrid['上网电量(kWh)'] - ongrid['上午可放电电量(kWh)']
        print(f"{ongrid['上网电量-上午可放电电量'].idxmax()}全天上网电量减去上午储能可放电电量最大,最大值为:{max(ongrid['上网电量-上午可放电电量'])}kWh")
        output_select = input("是否输出结果(是输入y,否输入n): ")
        if output_select == 'y':
            ongrid.to_excel('每日光伏上午电量及储能上午可放电量策略1.xlsx')
            print(f'储能pcs功率不小于:{max(df_new['上网功率(kW)'])}kW,储能容量不小于:{max(ongrid['上网电量(kWh)'])}kWh')
            df_new.plot()
            ongrid.plot()
            plt.show()
    except ValueError:
            # 如果输入不是数值，打印错误信息
            print("请输入一个有效的数值。")

def PVandESSAnalysis2(df1,df2):
    df_new = df1.copy()
    df_raw = df1.copy()
    ongrid = df2.copy()
    dischargetime = input('请输入上午时段储能充放电时段(8,10表示8点到10点段): ')
    ongrid['上网电量-上午可放电电量'] = ongrid['上网电量(kWh)'] - ongrid['上午可放电电量(kWh)']
    ongrid['光伏限发因子'] = 1.0
    while max(ongrid['上网电量-上午可放电电量']) >= 0:
        df_new['上网功率(kW)'] = df_new['光伏功率(kW)'] - df_new['负荷功率(kW)']
        ongrid.loc[:,['上网电量(kWh)','上午可放电电量(kWh)','上网电量-上午可放电电量']] = 0
        for row in range(len(df_new)):
            if df_new.iloc[row,2] > 0:
                ongrid.loc[ongrid.index == str(df_new.index[row]).split(' ')[0],'上网电量(kWh)'] = ongrid.loc[ongrid.index == str(df_new.index[row]).split(' ')[0],'上网电量(kWh)'] \
                                                                                    + df_new.iloc[row,2]

            if df_new.index[row].hour >= int(dischargetime.split(',')[0]) and df_new.index[row].hour < int(dischargetime.split(',')[1]) and df_new.iloc[row,2] < 0:
                ongrid.loc[ongrid.index == str(df_new.index[row]).split(' ')[0],'上午可放电电量(kWh)'] = ongrid.loc[ongrid.index == str(df_new.index[row]).split(' ')[0],'上午可放电电量(kWh)'] \
                                                                                    - df_new.values[row][2]
        ongrid['上网电量-上午可放电电量'] = ongrid['上网电量(kWh)'] - ongrid['上午可放电电量(kWh)']
        for row in range(len(ongrid)):
            if ongrid.iloc[row,2] > 0:
                ongrid.iloc[row,3] = ongrid.iloc[row,3] - 0.05
            
    
        for row in range(len(df_new)):
            df_new.iloc[row,0] = df_raw.iloc[row,0] * ongrid.loc[ongrid.index == str(df_new.index[row]).split(' ')[0],'光伏限发因子']
    
    output_select = input("是否输出结果(是输入y,否输入n): ")
    if output_select == 'y':
        ongrid.to_excel('每日光伏上午电量及储能上午可放电量策略2.xlsx')
        print(f'储能pcs功率不小于:{max(df_new['上网功率(kW)'])}kW,储能容量不小于:{max(ongrid['上网电量(kWh)'])}kWh')
        print(f'不限功率下光伏发电量:{df_raw['光伏功率(kW)'].sum()}kWh,限功率下光伏发电量:{df_new['光伏功率(kW)'].sum()}kWh\
              \n光伏年损失电量{df_raw['光伏功率(kW)'].sum()-df_new['光伏功率(kW)'].sum()}kWh')
        
        plt.figure()
        plt.plot(df_new[:], marker='none',label=['光伏功率(kW)','负荷功率(kW)','上网功率(kW)'])
        plt.title('光伏负荷及上网功率年曲线图')
        plt.legend()

        plt.figure()
        plt.plot(ongrid['光伏限发因子'], marker='o',label='光伏限发因子')
        plt.title('光伏限发因子')
        plt.legend()
        

        plt.figure()
        plt.plot(ongrid[['上网电量(kWh)','上午可放电电量(kWh)','上网电量-上午可放电电量']],marker='none',\
                 label=['上网电量(kWh)','上午可放电电量(kWh)','上网电量-上午可放电电量'])
        plt.title('全年每日上网及储能可放电电量')
        plt.legend()
        plt.show()
            


if __name__ == '__main__':
    app = QApplication(sys.argv)
    print('PVESSAnalysis软件主要根据光伏及负荷曲线进行分析,以光伏电量不上网考虑，计算光伏配置容量及储能配置容量')
    print('/**********************************************说明***********************************************************/')
    print('程序分为两种策略,策略1需要输入原始的年分时光伏功率和负荷功率曲线文件,格式为xlsx、xls或csv,上午指定单个储能可放电')
    print('时间段,折减因子，程序根据输入文件中的光伏功率乘以折减因子计算白天上网部分电量和上午指定时间段储能可以对外放电量，')
    print('当上网部分电量大于储能上午指定时间段的可放电量时，说明光伏装机容量较大，储能无法做到日两充两放，需要减小折减因子，直到')
    print('日最大上网部分电量减去储能上午指定时间段的可放电量小于或接近0为最佳,最终输出储能配置建议及光伏和负荷功率曲线，日上网电量')
    print('及上午指定时段储能可放电电量曲线。策略2输入文件为年分时光伏功率和负荷功率曲线文件,格式为xlsx、xls或csv,上午指定单个')
    print('储能可放电时间段，固定原始装机容量不变，计算每日上网电量及上午储能可放电量，当上网电量大于储能可放电量时，计算当日的光伏限发')
    print('因子，最终获得年每日光伏限发因子，光伏限发电量损失，储能最优配置容量。以上两种策略目的通过配储的方式，解决光伏发电量不上网')
    print('/***********************************************************************************************************/')
    print('作者:yb\n日期:2024.7.4')
    print('/***********************************************************************************************************/')
    file_path,_ = QFileDialog.getOpenFileName(None,'选择光伏负荷功率文件')
    if file_path != 0:
        file_extension = os.path.splitext(file_path)[1]
        if file_extension.lower() in ['.xls','.xlsx']:
            df = pd.read_excel(file_path,index_col=0)
        elif file_extension.lower() == '.csv':
            df = pd.read_csv(file_path,index_col=0,encoding='ANSI')
        else:
            print("unknow file type!")
    df.fillna(0)
    time_seq = pd.date_range(start='2023/1/1',end='2024/1/1',freq='1h',unit='s')
    time_seq2 = pd.date_range(start='2023/1/1',end='2023/12/31',freq='d',unit='s')
    togrid = pd.DataFrame()
    togrid['time'] = time_seq2
    togrid['上网电量(kWh)'] = 0.0
    togrid['上午可放电电量(kWh)'] = 0.0
    togrid = togrid.set_index('time')
    df['time'] = time_seq[0:-1]
    df_origin = df.set_index('time')
    df_origin.columns = ['光伏功率(kW)','负荷功率(kW)']
    strategy = input('策略选择(光伏装机可变化输入1,光伏装机固定功率变化输入2): ')
    if strategy == '1':
        print('你选择了策略1')
        while True:       
            PVandESSAnalysis1(df_origin,togrid) 
    else:
        print('你选择了策略2')
        while True:       
            PVandESSAnalysis2(df_origin,togrid)
    
             
        
        

    








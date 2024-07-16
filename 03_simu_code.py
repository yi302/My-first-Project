import numpy as np
import pandas as pd
import os
import math
import xlrd
# 定义常量
P_pv_max =1000 #光伏系统最大装机容量（kW）
N_b_max = 100 #最大可配置电动车数量
N_EV_max = 100 #最大可配置动力电池数量
N_CP_max = 50 #最大可安装充电桩数量
k1 = 5000  # 光伏系统单位投资成本 元/KW
k2 = 3000  # 充换电站投资成本，元/KWh
k3 = 10000  # 站外运输车辆静态投资成本系数，元/KWh
k4 = 10000 # 充电桩静态投资系数，元/个
k5 = 800 #储能子系统静态投资与装机容量比例系数（元/KWh）
k6 = 700 #储能子系统静态投资与装机功率比例系数（元/KW）
k7 = 500  # 光伏系统运营和维护成本，元/kw
k8 = 300  # 充换电站运营和维护成本，元/kwh
k9 = 1000  # 站外车辆运营和维护成本，元/KWh
K10 = 200 # 充电桩运营和维护成本，元/个
k11 = 300 #储能子系统运营和维护成本，元/KWh
re1 = 0.04  # 光伏残值率
re2 = 0.02  # 充换电站残值率
re3 = 0.2  # 电动运输残值率
re4 = 0.1 # 充电桩残值率
re5 = 0.05 #储能子系统残值率
k12 = 30000 #每辆电动车量的载货量（kg）
k13 = 5  #电动车续航里程与动力电池关系 km/kwh
V_b= 40 #动力电池的储能容量（KWh）
Y = 10  # 运营期十年
r = 0.08  # 贴现率为8%
alpha = 0.5  # 电费费率
P_charging = 20  # KW,充电桩最大充电功率
M_total = 150000  # kg，每日总站运输量
k14 = 2  # 每辆车每天可以来回运输的次数，即最大可换电次数
D_range = k13 * V_b #单车的运输距离
T = 24 / k14  #充电周期
E_E = 0.9 #储能系统充放电效率
grid_price = [0] * 24
for i in range(24):
    if 0 <= i < 1 or 7 <= i < 12 or 14 <= i < 16:
        grid_price[i] = 0.672
    elif 1 <= i < 7 or 12 <= i < 14:
        grid_price[i] = 0.323
    elif 16 <= i < 18 or 20 <= i < 24:
        grid_price[i] = 1.001
    elif 18 <= i < 20:
        grid_price[i] = 1.344
#将grid_price 扩展为8760项
grid_price = grid_price * 365

#函数判定是否满足换电需求
def able(T,N_CP,N_EV,P_charging,V_b):
    # 检查T的长度，并计算换电需求
    a = len(T)
    if a == 1:
        b = 24 * P_charging * N_CP / V_b - N_EV
    elif a == 2:
        b1 = (T[1] - T[0]) * P_charging * N_CP / V_b - N_EV
        b2 = (24 + T[0] - T[1]) * P_charging * N_CP / V_b - N_EV
    elif a == 3:
        b1 = (T[1] - T[0]) * P_charging * N_CP / V_b - N_EV
        b2 = (T[2] - T[1]) * P_charging * N_CP / V_b - N_EV
        b3 = (24 + T[0] - T[2]) * P_charging * N_CP / V_b - N_EV
    else:
        raise ValueError("T列表的长度必须是1, 2或3。")

    # 检查每个时间段的换电需求是否满足
    if a == 1:
        return b >= 0
    elif a == 2:
        return b1 >= 0 and b2 >= 0
    elif a == 3:
        return b1 >= 0 and b2 >= 0 and b3 >= 0

# 计算贴现值函数
def tiexian(NUM, Year, r):
    return sum(NUM / ((1 + r) ** i) for i in range(1, Year + 1))

# 目标函数
def objective(P_pv, P_bat, H_bat, N_CP, N_b, N_EV,data_input,T):
    data_df = data_input
    # 计算系统静态投资
    C_inv = k1 * P_pv + k2 * N_b * V_b + k3 * V_b * N_EV + k4 * N_CP

    # 计算储能子系统静态投资
    C_bat = k5 * H_bat +k6 * P_bat

    # 计算每年运营和维护成本
    C_m = k7 * P_pv + k8 * N_b * V_b + k9 * V_b * N_EV + K10 * N_CP + k11 * H_bat

    # 计算系统残值
    C_re = re1 * k1 * P_pv + re2 * k2 * N_b * V_b + re3 * k3 * V_b * N_EV + re4 *k4 * N_CP + re5 * C_bat

    # 计算每年运营和维护成本的贴现值
    C_m_tiexian = tiexian(C_m, Y, r)

    # 计算每天电网补电量，i_charging, PV_data为dataFrame，需要调整i_charging为365 * 24h
    data_df["output"] =  P_pv * data_df["output"]  #dataFrame数据，8760
    data_df["储能运行"] = 0
    data_df["储能系统容量"] = 0
    data_df_e = data_df["储能运行"].copy()
    #以下代码计算充电功率
    H_b = 0.4 * H_bat #定义储能系统的初容量
    EOD = 0.2
    P_CP_max = N_CP * P_charging
    for i in range(8760):
        a = max(0, min(((H_b - EOD * H_bat)*E_E), P_bat))  # a是i小时，储能最大放电功率；
        b = min(((H_bat-H_b)/E_E), P_bat) #b是i小时，储能最大充电功率；
        # 如果a + i时刻输入总功率能够进入充电桩的充电功率区间；
        # 储能运行策略，当充电桩充电功率不足时，补充到最大功率；
        # 按照最大可补电功率测算，即此时补电功率完全消纳；
        # 通过储能减小的弃电包括两部分，一是储能直接转移，二是由于储能补充功率，使得原有功率
        # 能够被利用，减少弃电
        if data_df["output"][i] >= P_CP_max:
            #光伏发电大于最大充电桩功率，此时储能电池为充电状态；
            data_df_e.iloc[i] = min((data_df["output"][i]-P_CP_max), b)
        else:
            #此时储能系统为充电桩供电，处于放电状态
            data_df_e.iloc[i] = -min(a, (P_CP_max-data_df["output"][i]))

        # 需要更新储能剩余容量，同时对储能功率进行修正
        if data_df_e.iloc[i] > 0: #充电效率
            H_b = H_b + data_df_e[i] * E_E
        else: #放电效率
            H_b = H_b + data_df_e[i] / E_E

        if H_b > H_bat:
            data_df_e.iloc[i] = (H_bat + data_df_e[i] * E_E - H_b)/E_E
            H_b = H_bat
        if H_b <  EOD * H_bat:
            data_df_e.iloc[i] = (data_df_e[i]/E_E - (H_b - H_bat * EOD)) * E_E
            H_b = EOD * H_bat

        data_df["储能系统容量"].iloc[i] = H_b

    data_df["储能运行"] = data_df_e
    #计算充电桩充电功率曲线：
    #首先是计算光伏参与充电桩的功率曲线，随后是电网补电参与充电桩，使其满足充电电池的用电需求
    data_df["光伏直接供电功率曲线"] = np.where(
        (data_df["output"] > 0) & (data_df["output"] <= P_CP_max),
        data_df["output"],
        np.where(data_df["output"] > P_CP_max, P_CP_max, 0)
    )
    #计算储能放电补充功率曲线
    data_df["储能放电补充功率曲线"] = data_df["储能运行"].loc[data_df["储能运行"] < 0]
    data_df["储能放电补充功率曲线"].fillna(value=0, inplace=True)
    data_df["光伏充电桩功率曲线"] = data_df["光伏直接供电功率曲线"]-data_df["储能放电补充功率曲线"]
    #计算电网补充下电功率曲线,根据充换电的要求进行计算；
    #根据每日调度要求在中午十二点和下午六点进行电池更换；
    #假定每个电池的初始的容量为：（1-EOD）* V_b；
    #定义待充电池N_b_a，在充电池N_b_b=N_CP，和充满电池为N_b_f
    #定义待充电池曲线，
    V_b_min = EOD * V_b
    data_df["待充电池数量曲线"]=N_b-N_CP
    data_df["在充电池数量曲线"]= N_CP
    data_df["充满电池数量曲线"] = 0
    data_df["在充电池充电功率"] = data_df["光伏充电桩功率曲线"] / N_CP
    data_df["在充电池容量"] = V_b_min
    data_df["在充电池容量"].iloc[0] = data_df["在充电池充电功率"].iloc[0] + V_b_min
    data_df["电网补电量"] = 0
    for i in range(1, 8760):
        if data_df["在充电池容量"].iloc[i] < V_b:
            data_df["在充电池容量"].iloc[i] = data_df["在充电池充电功率"].iloc[i] + data_df["在充电池容量"].iloc[i - 1]
            data_df["充满电池数量曲线"].iloc[i] = data_df["充满电池数量曲线"].iloc[i - 1]
            data_df["待充电池数量曲线"].iloc[i] = data_df["待充电池数量曲线"].iloc[i - 1]
        if data_df["在充电池容量"].iloc[i] >= V_b:
            data_df["充满电池数量曲线"].iloc[i] = data_df["充满电池数量曲线"].iloc[i - 1] + \
                                                  data_df["在充电池数量曲线"].iloc[i - 1]
            data_df["待充电池数量曲线"].iloc[i] = data_df["待充电池数量曲线"].iloc[i - 1] - \
                                                  data_df["在充电池数量曲线"].iloc[i - 1]
            data_df["在充电池容量"].iloc[i] = V_b_min + data_df["在充电池容量"].iloc[i] - V_b
        if (data_df["待充电池数量曲线"].iloc[i] < 0) and (-N_CP <= data_df["待充电池数量曲线"].iloc[i]):
            data_df["在充电池数量曲线"].iloc[i] = N_CP + data_df["待充电池数量曲线"].iloc[i]
            data_df["充满电池数量曲线"].iloc[i] = N_b - data_df["在充电池数量曲线"].iloc[i]
            data_df["待充电池数量曲线"].iloc[i] = 0
            data_df["在充电池充电功率"].iloc[i] = min(
                data_df["光伏充电桩功率曲线"].iloc[i] / data_df["在充电池数量曲线"].iloc[i], P_charging)
        if i % 24 == T[0]-1 or i % 24 == T[1]-1:  # 在每天的12点和下午六点进行电池更换
            data_df["充满电池数量曲线"].iloc[i] = data_df["充满电池数量曲线"].iloc[i] - N_EV
            data_df["待充电池数量曲线"].iloc[i] = data_df["待充电池数量曲线"].iloc[i] + N_EV
            if data_df["充满电池数量曲线"].iloc[i] < 0:  # 不满足换电要求，则需要在两次换电之间通过电网补电加大充电功率，每次补电都到达满功率
                # 不需要设计每一时刻的补电功率，直接计算该时段的电网补电量
                data_df["电网补电量"].iloc[i] = -data_df["充满电池数量曲线"].iloc[i] * V_b
                data_df["充满电池数量曲线"].iloc[i] = 0
                data_df["待充电池数量曲线"].iloc[i] = N_b - N_CP
        if data_df["充满电池数量曲线"].iloc[i]+ data_df["在充电池数量曲线"].iloc[i] >= N_b:
            data_df["在充电池数量曲线"].iloc[i] = N_b-data_df["充满电池数量曲线"].iloc[i]
        else:
            data_df["在充电池数量曲线"].iloc[i]= N_CP
            data_df["待充电池数量曲线"].iloc[i] = N_b - data_df["充满电池数量曲线"].iloc[i]- data_df["在充电池数量曲线"].iloc[i]

    P_grid_data = data_df["电网补电量"].sum()

    # 计算每年的电费投入
    C_ch = sum(P_grid_data * grid_price)
    C_ch_tiexian = tiexian(C_ch, Y, r)
    # 计算设备残值的贴现
    C_re_tiexian = C_re / ((1 + r) ** Y)
    # 计算总投入
    C_total = C_inv + C_m_tiexian + C_ch_tiexian - C_re_tiexian
    # 计算运输量
    T_EV = 365 * k14 * N_EV * D_range * k12
    T_EV_tiexian = tiexian(T_EV, Y, r)
    # 计算目标函数
    F = C_total / T_EV_tiexian

    return F, data_df

def plot_h(data_df):
    length = len(data_df)
    data_x = '['
    data_total = '['
    data_e = '['
    data_h = '['
    data_w = '['
    for i in range(length):
        data_x += '\'{}\','.format(data_df.index[i])
        data_total += '{:.2f},'.format(data_df["output"][i])
        data_e += '{:.2f},'.format(data_df["储能运行"][i])
        data_h += '{:.2f},'.format(data_df["光伏充电桩功率曲线"][i])
        data_w += '{:.2f},'.format(data_df["电网补电量"][i])
    data_x = data_x[:-1] + ']\n'
    data_total = data_total[:-1] + ']\n'
    data_e = data_e[:-1] + ']\n'
    data_h = data_h[:-1] + ']\n'
    data_w = data_w[:-1] + ']\n'
    message = '''
        <!DOCTYPE html>
        <html lang="zh-CN" style="height: 100%">
        <head>
          <meta charset="utf-8">
        </head>
        <body style="height: 100%; margin: 0">
          <div id="container" style="height: 100%"></div>


          <script type="text/javascript" src="https://fastly.jsdelivr.net/npm/echarts@5.4.0/dist/echarts.min.js"></script>

          <script type="text/javascript">
            var dom = document.getElementById('container');
            var myChart = echarts.init(dom, null, {
              renderer: 'canvas',
              useDirtyRect: false
            });
            var app = {};

            var option;

            option = {
          title: {
            text: '光储充换电站仿真运行曲线'
          },
          tooltip: {
            trigger: 'axis'
          },
          legend: {
            show:true
          },
          grid: {
            left: '3%',
            right: '4%',
            bottom: 50,
            containLabel: true,
            show: true,
            borderColor: "rgba(0, 0, 0, 1)"
          },
          toolbox: {
            show: true,
            feature: {
              dataZoom: {
                yAxisIndex: "none"
              },
              dataView: {
                readOnly: false
              },
              magicType: {
                type: ["line", "bar"]
              },
              restore: {}
            }
          },
          tooltip: {
                trigger: 'axis',
                axisPointer: {
                  type: 'cross'
                },
                backgroundColor: 'rgba(255, 255, 255, 0.8)'
              },
              axisPointer: {
                link: { xAxisIndex: 'all' },
                label: {
                  backgroundColor: '#777'
                }
              },
          dataZoom: [
            {
              show: true,
              realtime: true,
            },
            {
              type: 'inside',
              realtime: true,
            }
          ],
          xAxis: {
            type: 'category',
            boundaryGap: false,
            axisLine: {
              lineStyle: {
                color: "black"
              }
            },
            axisTick: {
              show: true,
              inside: true
            },
            data: '''
    message += data_x
    message += '''
            },
          yAxis: {
            type: 'value',
            name: "功率（kW）",
            nameTextStyle: {
              fontSize: 14
            },
            axisLine: {
                show:true,
              lineStyle: {
                color: "black"
              }
            },
            axisTick: {
              show: true,
              inside: true
            }
          },
          series: [
            {
              name: '光伏发电功率',
              type: 'line',
              smooth:true,
              markLine: {
              silent: false,
                  lineStyle: {
                    color: '#333'
                  },
                  label: {
                    position:'middle',
                    formatter: '平均{b}:{c} KW',
                    fontSize: 16
                  },
              data: [{
                type: "average",
                name:"光伏发电功率"
              }]
            },
              data: '''
    message += data_total
    message += '''},
            {
              name: '储能运行功率',
              type: 'line',
              smooth:true,
              data: '''
    message += data_e
    message += '''},
            {
              name: '充电桩运行功率',
              type: 'line',
              smooth:true,
              markLine: {
              silent: false,
                  lineStyle: {
                    color: '#333'
                  },
                  label: {
                    position:'middle',
                    formatter: '平均{b}:{c} KW',
                    fontSize: 14
                  },
              data: [{
                type: "average",
                name:"充电桩功率"
              }]
            },
              data: '''
    message += data_h
    message += '''},
            {
              name: '电网补电量（换电间隔）',
              type: 'bar',
              markLine: {
              silent: false,
                  lineStyle: {
                    color: '#333'
                  },
                  label: {
                    position:'middle',
                    formatter: '平均{b}:{c} KW',
                    fontSize: 14
                  },
              data: [{
                type: "average",
                name:"电网补电量（换电间隔）"
              }]
            },
              data: '''
    message += data_w
    message += '''}
          ]
        };

            if (option && typeof option === 'object') {
              myChart.setOption(option);
            }

            window.addEventListener('resize', myChart.resize);
          </script>
        </body>
        </html>
        '''
    path_tem = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates/')
    with open(path_tem + 'plot_res_hour.html', 'w', encoding="utf-8") as f:
        f.write(message)
    f.close()

def plot_bat(data_df): #绘制在站电池变化图
    length = len(data_df)
    data_x = '['
    data_total = '['
    data_e = '['
    data_h = '['
    for i in range(length):
        data_x += '\'{}\','.format(data_df.index[i])
        data_total += '{:.0f},'.format(data_df["待充电池数量曲线"][i])
        data_e += '{:.0f},'.format(data_df["在充电池数量曲线"][i])
        data_h += '{:.0f},'.format(data_df["充满电池数量曲线"][i])
    data_x = data_x[:-1] + ']\n'
    data_total = data_total[:-1] + ']\n'
    data_e = data_e[:-1] + ']\n'
    data_h = data_h[:-1] + ']\n'
    message = '''
        <!DOCTYPE html>
        <html lang="zh-CN" style="height: 100%">
        <head>
          <meta charset="utf-8">
        </head>
        <body style="height: 100%; margin: 0">
          <div id="container" style="height: 100%"></div>

          <script type="text/javascript" src="https://fastly.jsdelivr.net/npm/echarts@5.4.0/dist/echarts.min.js"></script>

          <script type="text/javascript">
            var dom = document.getElementById('container');
            var myChart = echarts.init(dom, null, {
              renderer: 'canvas',
              useDirtyRect: false
            });
            var app = {};

            var option;

            option = {
          title: {
            text: '光储充换电站在站动力电池仿真曲线'
          },
          tooltip: {
            trigger: 'axis'
          },
          legend: {
            show:true
          },
          grid: {
            left: '3%',
            right: '4%',
            bottom: 50,
            containLabel: true,
            show: true,
            borderColor: "rgba(0, 0, 0, 1)"
          },
          toolbox: {
            show: true,
            feature: {
              dataZoom: {
                yAxisIndex: "none"
              },
              dataView: {
                readOnly: false
              },
              magicType: {
                type: ["line", "bar"]
              },
              restore: {}
            }
          },
          tooltip: {
                trigger: 'axis',
                axisPointer: {
                  type: 'cross'
                },
                backgroundColor: 'rgba(255, 255, 255, 0.8)'
              },
              axisPointer: {
                link: { xAxisIndex: 'all' },
                label: {
                  backgroundColor: '#777'
                }
              },
          dataZoom: [
            {
              show: true,
              realtime: true,
            },
            {
              type: 'inside',
              realtime: true,
            }
          ],
          xAxis: {
            type: 'category',
            boundaryGap: false,
            axisLine: {
              lineStyle: {
                color: "black"
              }
            },
            axisTick: {
              show: true,
              inside: true
            },
            data: '''
    message += data_x
    message += '''
            },
          yAxis: {
            type: 'value',
            name: "电池数量（个）",
            nameTextStyle: {
              fontSize: 14
            },
            axisLine: {
                show:true,
              lineStyle: {
                color: "black"
              }
            },
            axisTick: {
              show: true,
              inside: true
            }
          },
          series: [
            {
              name: '待充电池数量',
              type: 'bar',
              markLine: {
              silent: false,
                  lineStyle: {
                    color: '#333'
                  },
                  label: {
                    position:'middle',
                    formatter: '平均{b}:{c} 个',
                    fontSize: 16
                  },
              data: [{
                type: "average",
                name:"待充电池数量"
              }]
            },
              data: '''
    message += data_total
    message += '''},
            {
              name: '在充电池数量',
              type: 'bar',
              data: '''
    message += data_e
    message += '''},
            {
              name: '充满电池数量',
              type: 'bar',
              markLine: {
              silent: false,
                  lineStyle: {
                    color: '#333'
                  },
                  label: {
                    position:'middle',
                    formatter: '平均{b}:{c} 个',
                    fontSize: 14
                  },
              data: [{
                type: "average",
                name:"充满电池数量"
              }]
            },
              data: '''
    message += data_h
    message += '''}
          ]
        };

            if (option && typeof option === 'object') {
              myChart.setOption(option);
            }

            window.addEventListener('resize', myChart.resize);
          </script>
        </body>
        </html>
        '''
    path_tem = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates/')
    with open(path_tem + 'plot_bat_hour.html', 'w', encoding="utf-8") as f:
        f.write(message)
    f.close()

def plot_d(data_df, name): #热力图绘制
    length = len(data_df)
    message = '''
        <!DOCTYPE html>
        <html lang="zh-CN" style="height: 100%">
        <head>
          <meta charset="utf-8">
        </head>
        <body style="height: 100%; margin: 0">
          <div id="container" style="height: 100%"></div>                 
          <script type="text/javascript" src="https://fastly.jsdelivr.net/npm/echarts@5.4.0/dist/echarts.min.js"></script>        
          
          <script type="text/javascript">
            var dom = document.getElementById('container');
            var myChart = echarts.init(dom, null, {
              renderer: 'canvas',
              useDirtyRect: false
            });
            var app = {};

            var option;
            '''
    message += 'const d=['
    j = 0
    sum = 0
    d = []
    for i in range(length):
        if name == "日历图：充电桩功率":
            sum += data_df["光伏充电桩功率曲线"][i]
        elif name == "日历图：光伏出力":
            sum += data_df["output"][i]
        elif name == "日历图：储能充电量":
            if data_df["储能运行"][i] < 0:
                sum += -data_df["储能运行"][i]
        elif name == "日历图：储能放电量":
            if data_df["储能运行"][i] > 0:
                sum += data_df["储能运行"][i]
        j += 1
        if j == 24:
            if name in ["日历图：储能放电量","日历图：储能充电量"]:
                d.append(sum)
                message += "{:.2f},".format(sum)
            else:
                d.append(sum / 24)
                message += "{:.2f},".format(sum / 24)
            j = 0
            sum = 0
    message = message[:-1] + "]\n"
    message += '''
            const d_max=Math.floor(Math.max(...d)/10)*10;
            const d_min=Math.floor(Math.min(...d)/10)*10;
            function getVirtualData(year) {
            const date = +echarts.time.parse(year + '-01-01');
            const end = +echarts.time.parse(+year + 1 + '-01-01');
            const dayTime = 3600 * 24 * 1000; 
            const data = [];
            var j=0;
            for (let time = date; time < end; time += dayTime) {
              data.push([
                echarts.time.format(time, '{yyyy}-{MM}-{dd}', false),d[j]]);
              j=j+1;
              }
              return data;
            }
        option = {
          title: {
            top: 6,
            left: 'center',
            text: 'Daily Step Count'
          },
          tooltip: {
            trigger: "item"
          },
          visualMap: {
            min:d_min,
            max:d_max,
            minOpen: true,
            maxOpen:true,
            inRange:{
              color: ["#313695", "#4575b4", "#74add1", "#abd9e9", "#e0f3f8", "#ffffbf", "#fee090", "#fdae61", "#f46d43", "#d73027", "#a50026"]
            },
            type: 'piecewise',
            orient: 'horizontal',
            left: 'center',
            top: 35
          },
          calendar: {
            top: 80,
            left: 30,
            right: 30,
            range: '2023',
            itemStyle: {
              borderWidth: 1
            },
            yearLabel: { show: false }
          },
          series: {'''
    if name in ['日历图：储能充电量','日历图：储能充电量']:
        message += "name:\"{} KWh\",\n".format(name)
    else:
        message += "name:\"{} KW\",\n".format(name)
    message += '''type: 'heatmap',
            coordinateSystem: 'calendar',
            data: getVirtualData('2023')
          }
        };

            if (option && typeof option === 'object') {
              myChart.setOption(option);
            }

            window.addEventListener('resize', myChart.resize);
          </script>
        </body>
        </html>
        '''
    path_tem = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates/')
    if name == "日历图：充电桩功率":
        filename = 'plot_res_day_h.html'
    elif name == "日历图：光伏出力":
        filename = 'plot_res_day_total.html'
    elif name == "日历图：储能放电量":
        filename = 'plot_res_day_e_dis.html'
    elif name == "日历图：储能充电量":
        filename = 'plot_res_day_e_ch.html'
    with open(path_tem + filename, 'w', encoding="utf-8") as f:
        f.write(message)
    f.close()

def data_save(data_df, save_path):
    # 将data_df所有列的数据输出到xlsx文件中：
    data_df.to_excel(save_path, index=False)

if __name__ =="__main__":
    pvfile = r"C:\00-工作资料20240303\05-污水处理厂综合能源管理平台\07-专利写作\06-一种光储充换电站及电动重卡容量优化配置方法\02-pv-data\02_WUhan_xinzhou_1KW.xlsx"
    save_path = r"C:\00-工作资料20240303\05-污水处理厂综合能源管理平台\07-专利写作\06-一种光储充换电站及电动重卡容量优化配置方法\04-code_result\01_results.xlsx"
    # 从 Excel 文件中读取光伏发电数据
    data_input = pd.read_excel(pvfile, sheet_name='Sheet1')  # 修改为实际 sheet 名称
    P_pv=100
    P_bat=25
    H_bat=50
    N_CP=3
    N_b=10
    N_EV=5 # 光伏100 KW， 储能25KW， 50KWh， 充电桩3个，动力电池10个，电动车辆5辆
    T=[12,18] #换电时间输入
    able_ = able(T, N_CP, N_EV, P_charging, V_b)
    if not able_:
        print("该配置不能满足换电需求，请提高充电桩充电功率、增加充电桩数量或调整充电时间安排")
    else:
        F,data_result = objective(P_pv, P_bat, H_bat, N_CP, N_b, N_EV, data_input,T)
        print("仿真计算已经完成，该配置下全生命周期单位运输成本为{:.3f}".format(F))
        plot_h(data_result)
        plot_bat(data_result)
        data_save(data_result, save_path) #保存DataFrame到xlsx文件
        plot_d(data_result, "日历图：充电桩功率")
        plot_d(data_result, "日历图：光伏出力")
        plot_d(data_result, "日历图：储能充电量")
        plot_d(data_result, "日历图：储能放电量")
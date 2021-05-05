import xlwt
import xlrd
import os
import time
import math
import itertools

# rederence:
# https://blog.csdn.net/YeChao3/article/details/83660615

# https://blog.csdn.net/liao392781/article/details/85338491
# Python3学习（三十八）：如何遍历一个目录下的同类型文件（csv、log等）

# https://blog.csdn.net/u013300049/article/details/79313979?depth_1-utm_source=distribute.pc_relevant_right.none-task&utm_source=distribute.pc_relevant_right.none-task
# itertools

# 读取课表信息，存于tab的列表当中
tab = []# 储存每个人的课表信息
info = []# 储存对应的 姓名_学号 信息

files = os.listdir('./course')
xls = list(filter(lambda x:x[-4:]=='.xls', files))
xlsx = list(filter(lambda x:x[-5:]=='.xlsx', files))

print("读取到以下表格文件：\n{}{}\n数量是：{}".format(xls,xlsx,len(xls)+len(xlsx)))

# 打开文件
for i in range(len(xls)):
    data = xlrd.open_workbook(filename="./course/"+xls[i])
    tab.append(data.sheets()[0])
    info.append(data.sheets()[0].cell(0,0).value[7:-7]+xls[i][7:-4])
    # print(info)
for i in range(len(xlsx)):
    data = xlrd.open_workbook("./course/"+xlsx[i])
    tab.append(data.sheets()[0])
    info.append(data.sheets()[0].cell(0,0).value[7:-7]+xlsx[i][7:-5])
    # print(info)
    # xlrd.open_workbook(filename=,encoding_override="utf-8")
'''amount = eval(input("请输入文件的数量"))
for i in range(1,amount+1):
    #data = xlrd.open_workbook("path of file")
    data = xlrd.open_workbook("arrange (" + str(i) + ").xls")
    tab.append(data.sheets()[0])
'''
# print输出人员信息，检查是否有误
print("\n经处理一共有{}张课表信息，人员如下：".format(len(tab)))
for i in range(len(tab)):
    print(info[i])

# 将所有成员信息写入convert当中的convert.xls内
# wb_c表格ws_c表，设置姓名班级对应的特殊昵称
wb_c = xlwt.Workbook(encoding='ascii')
ws_c = wb_c.add_sheet('Convert')
print('请在convert文件夹里的convert.xls里编辑姓名学号后一格对应的昵称,或者直接新建一个convert0.xls的文件，格式与前一个相同\n')
for i in range(len(tab)):
    ws_c.write(i,0,label=info[i])
wb_c.save('./convert/convert.xls')

if_nickname = input("是否需要对应昵称？\n1  是\n2  否\n")

if if_nickname == '1':
    input("请打开目录convert文件中的表格，在B列姓名右方写出对应昵称，另存为文件为convert0.xls后按回车键继续")
    # 需要更改昵称的话，在这里写代码
    # 打开convert0.xls表格
    try:
        data_conv = xlrd.open_workbook('./convert/convert0.xls')
    except:
        try:
            data_conv = xlrd.open_workbook('./convert/convert0.xlsx')
        except:
            print("格式错误")
    t_c = data_conv.sheets()[0]
    for i in range(len(info)):
        for j in range(t_c.nrows):
            # print(t_c.cell(j,0).value)
            if t_c.cell(j,0).value == info[i]:
                info[i] = t_c.cell(j,1).value
    print(info)

# 读取课程表周数[['6'], ['1-5', '7-14'], ['13-15']，['4', '6', '12', '14', '16']]
# 返回值是一个列表，包含单个周数信息，比如 [1, 2, 3, 4, 5, 6, 7, 8, 13, 14, 15]
def read(a):
    free = []
    if a == ' ' or a == '':
        None
    else:
        a = a.split('\n\n')
        for i_a in range(len(a)):
            a[i_a] = a[i_a].split('\n')
        a[0].remove('')
        for i_a in range(len(a)):
            try:
                if a[i_a] != ['']:
                    a[i_a][2] = a[i_a][2].replace('[周]','')
                    a[i_a][2] = a[i_a][2].replace('[单周]','')
                    a[i_a][2] = a[i_a][2].replace('[双周]','')
                    free.append(a[i_a][2].split(','))
            except:
                print("出现错误，请将此状况报告给开发者，谢谢")
    free2 = []
    if free == []:
        None
    else:
        for i_2 in range(len(free)):
            for j_2 in range(len(free[i_2])):
                if '-' in free[i_2][j_2]:
                    list_free2 = free[i_2][j_2].split('-')
                    for i_33 in range(eval(list_free2[0]),eval(list_free2[-1])+1):
                        free2.append(i_33)
                else:
                    free2.append(eval(free[i_2][j_2]))
    #print(free2)
    return free2


# 有课的课表
# 先建立一个框架，空白的
lis_busy = []
lis_free = []
week_l = 0
week_r = 19
for i_l in range(week_r):
    lis_busy.append([])
    lis_free.append([])
    for i_ll in range(7):
        lis_busy[i_l].append([])
        lis_free[i_l].append([])
        for i_lll in range(6):
            lis_busy[i_l][i_ll].append([])
            lis_free[i_l][i_ll].append([])


# worksheet可以通过tab[]的索引
# 注意，为了方便起见，这里不在从0开始计数
for i in range(len(tab)):
    #
    for row in range(3, 9):
        for col in range(1, 8):
            fr = read(tab[i].cell(row, col).value)
            for i_week in range(len(lis_busy)):
                if i_week in fr:
                    lis_busy[i_week][col-1][row-3].append(info[i])
                else:
                    lis_free[i_week][col-1][row-3].append(info[i])
            # print('第四周星期二{}'.format(lis[4][2]))
# print(lis_busy)
# print(lis_free)
# 整合各列表存于wbd当中，同时储存在lis里
wb_busy = xlwt.Workbook(encoding='ascii')
'''[week1[Mon[[0102][0304][0506][0708][0910][1112]]Tue[]Wed[]Thu[]Fri[]Sat[]Sun[]
          ]
    week2[]
    ......]'''


# wb_busy.save('./busytime/busytime.xls')
# print(lis)
# 到此为止，已经建立了一个list，有没时间的人都写了进去，开心，任务完成了一半
'''
lis_busy 有课时间表
lis_free 无课时间表
格式：有没空的人 = lis_busy[第一周][星期八][78节课]
info     人员名单
'''

'''
x = itertools.combinations_with_replacement(info, 100)
print(list(x)[30:100])
'''


# 安排傍晚值班的程序
'''
def arrange_evening(bgwk,bgdy,edwk,eddy,people):
    print("step1 计算人均工作最大次数：", end='')
    days = (edwk-bgwk)*7+eddy-bgdy+1
    average_timesum = math.ceil(days * people / len(info))
    print(average_timesum)
    weekday_cal = []
    wk = bgwk
    dy = bgdy
    for i in range(days):
        temp_weekday_cal = (wk,dy)#week,day
        if dy == 6:
            dy = 0
            wk += 1
        else:
            dy += 1
        weekday_cal.append(temp_weekday_cal)
    print("step2 新建列表")
    #ws_wholeday.write(0, 0, label='')
    list_wholeday = []# 表格
    list_pre = []
    for i in range(days):
        list_wholeday.append([])
        list_pre.append({})
    list_used = {}# 统计已经安排过的人的次数以及时间
    for i in info:
        list_used[i] = 0
    # 预统计
    for i in range(len(list_wholeday)):
        for j in range(block):
            temp_dict_peo = {}
            for k in info:
                temp_dict_peo[k]=suit_wholeday_value(k,j,weekday_cal[i][0],weekday_cal[i][1])
            list_pre[i][j] = temp_dict_peo
    print("step3 开始第一次安排")
    # 开始遍历每一天
    for i in range(len(list_wholeday)):
        # 遍历每一天当中的时间段
        for j in range(block):
            temp_dict = {}
            temp_peo = []
            if_arrange = 0
            # 遍历预统计当中的人
            for p in range(people):

                for k in sorted(list_pre[i][j]):
                    # 判断100的人以及未安排到的人
                    if list_pre[i][j][k] >= 60 and list_used[k] < average_timesum and not k in temp_peo and not if_in_list(k,list_wholeday[i]):
                        temp_peo.append(k)
                        list_used[k] += 1
                        if_arrange = 1
                        break
                if if_arrange == 0:
                    temp_peo.append(' ')
                if_arrange = 0
            list_wholeday[i].append(temp_peo)
    #print(list_wholeday)
    print("step4 开始第二次安排")

    print("step5 开始写入信息")
    wb_wholeday = xlwt.Workbook(encoding='ascii')
    ws_wholeday = wb_wholeday.add_sheet('值班表')
    ws_times = wb_wholeday.add_sheet("值班次数统计")
    for i in range(len(tab)):
        ws_times.write(i, 0, label=info[i])
        ws_times.write(i, 1, label=list_used[info[i]])
    ws_wholeday.write(0, 0, label='周数星期')
    ws_wholeday.write(0, 1, label='时间段')
    ws_wholeday.write(0, 2, label='人员')
    for i in range(days):
        ws_wholeday.write(i * block + 1, 0, label='第'+str(weekday_cal[i][0]+1)+'周 星期'+str(weekday_cal[i][1]+1))
        ws_wholeday.write(i * block + 1, 1, label='8:00-10:00')
        ws_wholeday.write(i * block + 2, 1, label='10:00-12:00')
        ws_wholeday.write(i * block + 3, 1, label='12:00-下午课前')
        ws_wholeday.write(i * block + 4, 1, label='下午56节课时间段')
        ws_wholeday.write(i * block + 5, 1, label='下午78节课时间段')
        if block == 6:
            ws_wholeday.write(i * block + 6, 1, label='下午78节课后-傍晚')
        for j in range(len(list_wholeday[i])):
            for k in range(len(list_wholeday[i][j])):
                ws_wholeday.write(i * block + 1 + j, k+2, label=list_wholeday[i][j][k])

    filename = input("step6 请输入保存的文件名:")
    wb_wholeday.save('./'+filename+'.xls')
    print("Succcessful")
'''

# 计算全天值班的合适概率
def suit_wholeday_value(mem, num, week, day):
    output = 100
    if num == 0:#8-10
        if mem in lis_busy[week][day][0]:
            output = 0
        elif mem in lis_busy[week][day][1]:
            output -= 10
    elif num == 1:#10-12
        if mem in lis_busy[week][day][1]:
            output = 0
        elif mem in lis_busy[week][day][0]:
            output -= 10
    elif num == 2:#12-before class 56
        if mem in lis_busy[week][day][1]:
            output -= 10
        if mem in lis_busy[week][day][2]:
            output -= 10
    elif num == 3:#class5-6
        if mem in lis_busy[week][day][2]:
            output = 0
        elif mem in lis_busy[week][day][3]:
            output -= 10
    elif num == 4:#class7-8
        if mem in lis_busy[week][day][3]:
            output = 0
        elif mem in lis_busy[week][day][2]:
            output -= 10
    elif num == 5:#after class78
        if mem in lis_busy[week][day][3]:
            output -= 10
        if mem in lis_busy[week][day][4]:
            output -= 10
    return output


# 判断子列表是否含有某元素
def if_in_list(k, lis):
    state = 0
    for i in lis:
        if k in i:
            state = 1
    if state == 1:
        return True
    else:
        return False


# 安排整天的程序
def arrange_wholeday(bgwk,bgdy,edwk,eddy,people,ifdawn):
    print("开始安排")
    if ifdawn == 1:
        block = 6
    else:
        block = 5

    print("step1 计算人均工作最大次数：", end='')
    days = (edwk-bgwk)*7+eddy-bgdy+1
    average_timesum = math.ceil(block * days * people / len(info))
    print(average_timesum)
    weekday_cal = []
    wk = bgwk
    dy = bgdy
    for i in range(days):
        temp_weekday_cal = (wk,dy)#week,day
        if dy == 6:
            dy = 0
            wk += 1
        else:
            dy += 1
        weekday_cal.append(temp_weekday_cal)
    print("step2 新建列表")
    #ws_wholeday.write(0, 0, label='')
    list_wholeday = []# 表格
    list_pre = []
    for i in range(days):
        list_wholeday.append([])
        list_pre.append({})
    list_used = {}# 统计已经安排过的人的次数以及时间
    for i in info:
        list_used[i] = 0
    # 预统计
    for i in range(len(list_wholeday)):
        for j in range(block):
            temp_dict_peo = {}
            for k in info:
                temp_dict_peo[k]=suit_wholeday_value(k,j,weekday_cal[i][0],weekday_cal[i][1])
            list_pre[i][j] = temp_dict_peo
    print("step3 开始第一次安排")
    # 开始遍历每一天
    for i in range(len(list_wholeday)):
        # 遍历每一天当中的时间段
        for j in range(block):
            temp_dict = {}
            temp_peo = []
            if_arrange = 0
            # 遍历预统计当中的人
            for p in range(people):

                for k in sorted(list_pre[i][j]):
                    # 判断100的人以及未安排到的人
                    if list_pre[i][j][k] >= 60 and list_used[k] < average_timesum \
                            and not k in temp_peo and not if_in_list(k,list_wholeday[i]):
                        temp_peo.append(k)
                        list_used[k] += 1
                        if_arrange = 1
                        break
                if if_arrange == 0:
                    temp_peo.append(' ')
                if_arrange = 0
            list_wholeday[i].append(temp_peo)
    #print(list_wholeday)
    print("step4 开始第二次安排")

    print("step5 开始写入信息")
    wb_wholeday = xlwt.Workbook(encoding='ascii')
    ws_wholeday = wb_wholeday.add_sheet('值班表')
    ws_times = wb_wholeday.add_sheet("值班次数统计")
    for i in range(len(tab)):
        ws_times.write(i, 0, label=info[i])
        ws_times.write(i, 1, label=list_used[info[i]])
    ws_wholeday.write(0, 0, label='周数星期')
    ws_wholeday.write(0, 1, label='时间段')
    ws_wholeday.write(0, 2, label='人员')
    for i in range(days):
        ws_wholeday.write(i * block + 1, 0, label='第'+str(weekday_cal[i][0]+1)+'周 星期'+str(weekday_cal[i][1]+1))
        ws_wholeday.write(i * block + 1, 1, label='8:00-10:00')
        ws_wholeday.write(i * block + 2, 1, label='10:00-12:00')
        ws_wholeday.write(i * block + 3, 1, label='12:00-下午课前')
        ws_wholeday.write(i * block + 4, 1, label='下午56节课时间段')
        ws_wholeday.write(i * block + 5, 1, label='下午78节课时间段')
        if block == 6:
            ws_wholeday.write(i * block + 6, 1, label='下午78节课后-傍晚')
        for j in range(len(list_wholeday[i])):
            for k in range(len(list_wholeday[i][j])):
                ws_wholeday.write(i * block + 1 + j, k+2, label=list_wholeday[i][j][k])

    filename = input("step6 请输入保存的文件名:")
    wb_wholeday.save('./'+filename+'.xls')
    print("Succcessful")


# 转换4个输入时间为输出的四个字符
def time_convert1():
    try:
        temp = input("请以空格分隔来输入，输入完成后回车\n开始周数 开始星期数 结束周数 结束星期数 单次人数\n请输入:")
        temp = temp.split(" ")
        a = eval(temp[0])-1
        b = eval(temp[1])-1
        c = eval(temp[2])-1
        d = eval(temp[3])-1
        e = eval(temp[4])
        return a, b, c, d, e
    except:
        print("输入错误")

# 转换4个输入时间为输出的四个字符
def time_convert2():
    try:
        temp = input("请以空格分隔来输入，输入完成后回车\n开始周数 开始星期数 结束周数 结束星期数 单次人数 是否需要下午课以后（是1 否0）\n请输入:")
        temp = temp.split(" ")
        a = eval(temp[0])-1
        b = eval(temp[1])-1
        c = eval(temp[2])-1
        d = eval(temp[3])-1
        e = eval(temp[4])
        f = eval(temp[5])
        return a, b, c, d, e, f
    except:
        print("输入错误")


def main():
    choice = input("请输入数字后回车进行选择\n1 安排傍晚值班\n2 安排全天值班\n3 退出\n请输入:")
    if choice == '1':
        print("安排傍晚值班")
        #(a,b,c,d,e)=time_convert1()
        #arrange_evening(a,b,c,d,e)
        print("此功能还在开发中，敬请期待")
    elif choice == '2':
        print("安排全天值班")
        (a,b,c,d,e,f)=time_convert2()
        arrange_wholeday(a,b,c,d,e,f)
    elif choice == '3':
        quit()


main()
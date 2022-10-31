#coding=utf8
import xlrd
import numpy as np

# 数据集路径
DATASET_PATH = ''
# 载入数据集文件
DataFile = xlrd.open_workbook(DATASET_PATH)
# 载入第一个表格
table = DataFile.sheets()[0]

dataset = []
# 属性和其离散取值数量（连续属性用二分法进行离散取值）
# attributesdict = {'色泽':['青绿', '乌黑', '浅白'], '根蒂':['蜷缩', '稍蜷', '硬挺'], '敲声':['浊响', '沉闷', '清脆'],\
#                   '纹理':['清晰', '稍糊', '模糊'], '脐部':['凹陷', '稍凹', '平坦'], '触感':['硬滑', '软粘'],\
#                   '密度':['0', '1'], '含糖率':['0', '1']}
attributesdict = {'色泽':['青绿', '乌黑', '浅白'], '根蒂':['蜷缩', '稍蜷', '硬挺'], '敲声':['浊响', '沉闷', '清脆'],\
                  '纹理':['清晰', '稍糊', '模糊'], '脐部':['凹陷', '稍凹', '平坦'], '触感':['硬滑', '软粘']}

# 将表格信息转化为字典
def Read_Excel(excel):
    for rows in range(1, excel.nrows):
        dict_ = {'编号':'', '色泽':'', '根蒂':'', '敲声':'', '纹理':'',\
                 '脐部':'', '触感':'', '密度':'', '含糖率':'', '好瓜':''}
        dict_['编号'] = table.cell_value(rows, 0)
        dict_['色泽'] = table.cell_value(rows, 1)
        dict_['根蒂'] = table.cell_value(rows, 2)
        dict_['敲声'] = table.cell_value(rows, 3)
        dict_['纹理'] = table.cell_value(rows, 4)
        dict_['脐部'] = table.cell_value(rows, 5)
        dict_['触感'] = table.cell_value(rows, 6)
        dict_['密度'] = table.cell_value(rows, 7)
        dict_['含糖率'] = table.cell_value(rows, 8)
        dict_['好瓜'] = table.cell_value(rows, 9)
        dataset.append(dict_)
    return dataset

# 对数据进行排序并计算其中值
def Get_MidValue(data, class_name):
    sort = []
    midvalue = []
    # 取出所有密度值
    for i in range(len(data)):
        sort.append(data[i][class_name])
    # 冒泡排序
    for x in range(len(sort)):
        for y in range(x, len(sort)):
            if sort[x] > sort[y]:
                temp = sort[x]
                sort[x] = sort[y]
                sort[y] = temp
    # 求中间值
    for i in range(len(sort)-1):
        midvalue.append(round(((sort[i]+sort[i+1])/2), 3))
    return midvalue


# 计算信息熵
def Calculate_Entropy(GOODMELON_NUM, BADMELON_NUM, TOTAL_NUM):
    if TOTAL_NUM != 0:
        GOODMELON_RATIO = GOODMELON_NUM / TOTAL_NUM
        BADMELON_RATIO = BADMELON_NUM / TOTAL_NUM
        if GOODMELON_RATIO == 0:
            Ent = -(BADMELON_RATIO*np.log2(BADMELON_RATIO))
        if BADMELON_RATIO == 0:
            Ent = -(GOODMELON_RATIO*np.log2(GOODMELON_RATIO))
        if GOODMELON_RATIO != 0 and BADMELON_RATIO != 0:
            Ent = -(GOODMELON_RATIO*np.log2(GOODMELON_RATIO) + BADMELON_RATIO*np.log2(BADMELON_RATIO))
    else:
        Ent = 0
    return Ent


# 计算信息增益，输入数据与希望计算信息增益的属性名称
def Calculate_InformationGain(data, class_name):
    # 若属性为连续取值
    if class_name == '密度' or class_name == '含糖率':
        Gains = []
        # 获取所有候选划分点
        midvalues = Get_MidValue(data, class_name)
        for midvalue in midvalues:
            GOODMELON_NUM = 0
            BADMELON_NUM = 0
            TOTAL_NUM = 0
            CLASS1_GOODMELON_NUM = 0
            CLASS1_BADMELON_NUM = 0
            CLASS1_NUM = 0
            CLASS2_GOODMELON_NUM = 0
            CLASS2_BADMELON_NUM = 0
            CLASS2_NUM = 0
            for i in data:
                TOTAL_NUM += 1
                if i[class_name] > midvalue:
                    CLASS1_NUM += 1
                    if i['好瓜'] == '是':
                        GOODMELON_NUM += 1
                        CLASS1_GOODMELON_NUM += 1
                    else:
                        BADMELON_NUM += 1
                        CLASS1_BADMELON_NUM += 1
                else:
                    CLASS2_NUM += 1
                    if i['好瓜'] == '是':
                        GOODMELON_NUM += 1
                        CLASS2_GOODMELON_NUM += 1
                    else:
                        BADMELON_NUM += 1
                        CLASS2_BADMELON_NUM += 1
            # 计算总信息熵
            GOODMELON_RATIO = GOODMELON_NUM / TOTAL_NUM
            BADMELON_RATIO = BADMELON_NUM / TOTAL_NUM
            if GOODMELON_RATIO == 0:
                Ent_D = -(BADMELON_RATIO * np.log2(BADMELON_RATIO))
            if BADMELON_RATIO == 0:
                Ent_D = -(GOODMELON_RATIO * np.log2(GOODMELON_RATIO))
            if GOODMELON_RATIO != 0 and BADMELON_RATIO != 0:
                Ent_D = -(GOODMELON_RATIO * np.log2(GOODMELON_RATIO) + BADMELON_RATIO*np.log2(BADMELON_RATIO))
            # 计算属性为第一种取值的信息熵
            Ent_D1 = Calculate_Entropy(CLASS1_GOODMELON_NUM, CLASS1_BADMELON_NUM, CLASS1_NUM)
            # 计算第二种取值的信息熵
            Ent_D2 = Calculate_Entropy(CLASS2_GOODMELON_NUM, CLASS2_BADMELON_NUM, CLASS2_NUM)
            # 计算信息增益
            Gain = Ent_D-((CLASS1_NUM/TOTAL_NUM) * Ent_D1 + (CLASS2_NUM/TOTAL_NUM) * Ent_D2)
            # 记录信息增益
            Gains.append(round(Gain, 3))
        temp = 0
        for i in range(len(Gains)):
            if Gains[i] > temp:
                temp = Gains[i]
                mark = i
        return GOODMELON_NUM, BADMELON_NUM, Gains[mark], midvalues[mark]
    # 若属性为离散取值
    else:
        if class_name in list(attributesdict.keys()):
            GOODMELON_NUM = 0
            BADMELON_NUM = 0
            TOTAL_NUM = 0
            CLASS1_GOODMELON_NUM = 0
            CLASS1_BADMELON_NUM = 0
            CLASS1_NUM = 0
            CLASS2_GOODMELON_NUM = 0
            CLASS2_BADMELON_NUM = 0
            CLASS2_NUM = 0
            CLASS3_GOODMELON_NUM = 0
            CLASS3_BADMELON_NUM = 0
            CLASS3_NUM = 0
            for i in data:
                TOTAL_NUM += 1
                if i[class_name] == attributesdict[class_name][0]:
                    CLASS1_NUM += 1
                    if i['好瓜'] == '是':
                        GOODMELON_NUM += 1
                        CLASS1_GOODMELON_NUM += 1
                    else:
                        BADMELON_NUM += 1
                        CLASS1_BADMELON_NUM += 1
                elif i[class_name] == attributesdict[class_name][1]:
                    CLASS2_NUM += 1
                    if i['好瓜'] == '是':
                        GOODMELON_NUM += 1
                        CLASS2_GOODMELON_NUM += 1
                    else:
                        BADMELON_NUM += 1
                        CLASS2_BADMELON_NUM += 1
                elif len(attributesdict[class_name]) == 3:
                    if i[class_name] == attributesdict[class_name][2]:
                        CLASS3_NUM += 1
                        if i['好瓜'] == '是':
                            GOODMELON_NUM += 1
                            CLASS3_GOODMELON_NUM += 1
                        else:
                            BADMELON_NUM += 1
                            CLASS3_BADMELON_NUM += 1
            # 计算总信息熵
            GOODMELON_RATIO = GOODMELON_NUM / TOTAL_NUM
            BADMELON_RATIO = BADMELON_NUM / TOTAL_NUM
            if GOODMELON_RATIO == 0 and BADMELON_RATIO == 0:
                return GOODMELON_NUM, BADMELON_NUM, 0
            elif BADMELON_RATIO == 0 and GOODMELON_RATIO != 0:
                Ent_D = -(GOODMELON_RATIO*np.log2(GOODMELON_RATIO))
            elif GOODMELON_RATIO == 0 and BADMELON_RATIO != 0:
                Ent_D = -(BADMELON_RATIO*np.log2(BADMELON_RATIO))
            else:
                Ent_D = -(GOODMELON_RATIO*np.log2(GOODMELON_RATIO) + BADMELON_RATIO*np.log2(BADMELON_RATIO))
            # 计算属性为第一种取值的信息熵
            Ent_D1 = Calculate_Entropy(CLASS1_GOODMELON_NUM, CLASS1_BADMELON_NUM, CLASS1_NUM)
            # 计算第二种取值的信息熵
            Ent_D2 = Calculate_Entropy(CLASS2_GOODMELON_NUM, CLASS2_BADMELON_NUM, CLASS2_NUM)
            # 计算第三种取值的信息熵
            Ent_D3 = Calculate_Entropy(CLASS3_GOODMELON_NUM, CLASS3_BADMELON_NUM, CLASS3_NUM)
            # 计算信息增益
            Gain = Ent_D - ((CLASS1_NUM/TOTAL_NUM)*Ent_D1+(CLASS2_NUM/TOTAL_NUM)*Ent_D2+(CLASS3_NUM/TOTAL_NUM)*Ent_D3)
            return GOODMELON_NUM, BADMELON_NUM, round(Gain, 3)


# 生成决策树
def TreeGenerate(goodmelon_num, badmelon_num, data, attributes_dict):
    new_attributes_dict = attributes_dict
    # 生成结点
    DecisionTree = {'node':''}
    SIMILARMARK_NUM = 0 # 属于同一类别（标签）的样本的数量
    SIMILARCOLOR_NUM = 0 # 属于同一颜色的样本的数量
    SIMILARROOT_NUM = 0 # 属于同一根蒂的样本的数量
    SIMILARVOICE_NUM = 0 # 属于同一敲声的样本的数量
    SIMILARTEXTURE_NUM = 0 # 属于同一纹理的样本的数量
    SIMILARUMBILICAL_NUM = 0 # 属于同一脐部的样本的数量
    SIMILARTOUCH_NUM = 0 # 属于同一触感的样本的数量
    SIMILARDENSITY_NUM = 0 # 属于同一密度的样本的数量
    SIMILARSUGAR_NUM = 0 # 属于同一含糖率的样本的数量
    # 若只有一个数据，则直接标记为叶结点
    if len(data) == 1:
        if data[0]['好瓜'] == '是':
            return '好瓜'
        else:
            return '坏瓜'
    for i in range(len(data)-1):
        # 若所有样本属于同一类别，则将结点标记为叶结点
        if data[i]['好瓜'] == data[i+1]['好瓜']:
            SIMILARMARK_NUM += 1
            if SIMILARMARK_NUM == len(data)-1:
                if data[i]['好瓜'] == '是':
                    return '好瓜'
                else:
                    return '坏瓜'
        if '色泽' in list(data[i].keys()) and data[i]['色泽'] == data[i+1]['色泽']:
            SIMILARCOLOR_NUM += 1
        if '根蒂' in list(data[i].keys()) and data[i]['根蒂'] == data[i+1]['根蒂']:
            SIMILARROOT_NUM += 1
        if '敲声' in list(data[i].keys()) and data[i]['敲声'] == data[i+1]['敲声']:
            SIMILARVOICE_NUM += 1
        if '纹理' in list(data[i].keys()) and data[i]['纹理'] == data[i+1]['纹理']:
            SIMILARTEXTURE_NUM += 1
        if '脐部' in list(data[i].keys()) and data[i]['脐部'] == data[i+1]['脐部']:
            SIMILARUMBILICAL_NUM += 1
        if '触感' in list(data[i].keys()) and data[i]['触感'] == data[i+1]['触感']:
            SIMILARTOUCH_NUM += 1
        if '密度' in list(data[i].keys()) and data[i]['密度'] == data[i+1]['密度']:
            SIMILARDENSITY_NUM += 1
        if '含糖率' in list(data[i].keys()) and data[i]['含糖率'] == data[i+1]['含糖率']:
            SIMILARSUGAR_NUM += 1
    # 若样本在所有属性上取值相同，则将结点标记为叶结点
    if SIMILARCOLOR_NUM == SIMILARROOT_NUM == SIMILARVOICE_NUM == SIMILARTEXTURE_NUM ==\
       SIMILARUMBILICAL_NUM == SIMILARTOUCH_NUM == SIMILARDENSITY_NUM == SIMILARSUGAR_NUM == len(data)-1:
        if goodmelon_num < badmelon_num:
            return '坏瓜'
        else:
            return '好瓜'

    # 选择最优划分属性
    InformationGains = {}
    for attribute in list(new_attributes_dict.keys()):
        InformationGains[attribute] = Calculate_InformationGain(data, attribute)
    temp = 0
    # 找到信息增益最大的那个属性
    for i in list(InformationGains.keys()):
        if InformationGains[i][2] > temp:
            temp = InformationGains[i][2]
            best_att = i
    DecisionTree[best_att] = DecisionTree.pop('node')
    DecisionTree[best_att] = {}
    
    # 二分法更新连续值属性的取值
    if best_att == '密度' or best_att == '含糖率':
        for i in range(len(data)):
            if data[i][best_att] > InformationGains[best_att][3]:
                data[i][best_att] = '0'
            else:
                data[i][best_att] = '1'
        # 构造新的分枝
        for value in new_attributes_dict[best_att]:
            if value == '0':
                value = '大于' + str(InformationGains[best_att][3])
            else:
                value = '小于' + str(InformationGains[best_att][3])
            DecisionTree[best_att][best_att+value] = ''
    # 离散取值变量只需要新建分枝
    else:
        for value in new_attributes_dict[best_att]:
            DecisionTree[best_att][best_att+value] = ''
    # 获取属性的可能取值
    for value in list(DecisionTree[best_att].keys()):
        value = value.replace(best_att, '')
        # 取出在该属性下的特定取值的数据
        new_data = []
        for item in data:
            # 连续属性
            if best_att == '密度' or best_att == '含糖率':
                if (value == '大于' + str(InformationGains[best_att][3]) and item[best_att] == '0') or\
                value == '小于' + str(InformationGains[best_att][3]) and item[best_att] == '1':
                    new_data.append(item)
            # 离散属性
            elif item[best_att] == value:
                new_data.append(item)
        # 若样本在某可能取值下为空，则标记为叶结点，其类别为D中样本最多的类
        if new_data == []:
            if goodmelon_num < badmelon_num:
                DecisionTree[best_att][best_att+value] = '坏瓜'
            else:
                DecisionTree[best_att][best_att+value] = '好瓜'
        # 数据中含有该属性的可能取值，即对该取值进行展开
        else:
            if best_att in list(new_attributes_dict.keys()):
                new_attributes_dict.pop(best_att)
            # 若所有属性都已划分
            if new_attributes_dict == {}:
                if goodmelon_num < badmelon_num:
                    DecisionTree[best_att][best_att+value] = '坏瓜'
                else:
                    DecisionTree[best_att][best_att+value] = '好瓜'
            # 以符合属性取值的新数据与新的属性表构造分枝
            DecisionTree[best_att][best_att+value] = TreeGenerate(8, 9, new_data, new_attributes_dict)
            # 重新补充属性，防止后续for循环缺失属性
            if best_att == '色泽':
                new_attributes_dict['色泽'] = ['青绿', '乌黑', '浅白']
            if best_att == '根蒂':
                new_attributes_dict['根蒂'] = ['蜷缩', '稍蜷', '硬挺']
            if best_att == '敲声':
                new_attributes_dict['敲声'] = ['浊响', '沉闷', '清脆']
            if best_att == '纹理':
                new_attributes_dict['纹理'] = ['清晰', '稍糊', '模糊']
            if best_att == '脐部':
                new_attributes_dict['脐部'] = ['凹陷', '稍凹', '平坦']
            if best_att == '触感':
                new_attributes_dict['触感'] = ['硬滑', '软粘']
    return DecisionTree[best_att]


print(Read_Excel(table))
print(TreeGenerate(8, 9, dataset, attributesdict))

import pulp
import numpy as np
import pandas as pd

sheet1 = pd.read_excel('datas/附件2.xlsx', sheet_name='2023年统计的相关数据')  # 土地类型上的亩产量数据
sheet1 = sheet1.dropna(subset=['地块类型'])

sheet2 = pd.read_excel('datas/附件2.xlsx', sheet_name='2023年的农作物种植情况')  # 作物在不同地块上的种植面积数据
sheet3 = pd.read_excel('datas/附件1.xlsx', sheet_name='乡村的现有耕地')  # 地块编号对应的土地类型

#获取预测销售量
# 1. 将 sheet2 和 sheet3 合并，得到每个地块的土地类型
land_types = pd.merge(sheet2, sheet3, left_on='种植地块', right_on='地块名称')
land_types = land_types.drop(columns=['说明 ','地块名称','地块面积/亩','作物编号'])
sheet11 = sheet1.drop(columns=['序号','作物编号'])

print(land_types)
print(sheet11)

# 2. 将合并后的数据与 sheet1 合并，得到每个作物在每种土地类型上的亩产量
yield_data = pd.merge(land_types, sheet11, left_on=['地块类型','作物名称','种植季次'], right_on=['地块类型','作物名称','种植季次'])
# yield_data = yield_data.drop(columns=['销售单价/(元/斤)','种植成本/(元/亩)'])
print(yield_data)

# 3. 计算每种作物的产量（种植面积 * 亩产量）
yield_data['产量'] = yield_data['种植面积/亩'] * yield_data['亩产量/斤']

# 4. 汇总每种作物的总产量作为预测销售量
expected_sales_volume = yield_data.groupby(['作物编号','作物名称_x','种植季次_x'])['产量'].sum().reset_index()
expected_sales_volume = expected_sales_volume.rename(columns={'作物名称_x': '作物名称','产量':'预期销售量'})

# print(expected_sales_volume)

#线性规划模型设计

# 约束条件留存
        # (1) 平旱地、梯田和山坡地每年都只能种植一季粮食类作物 1
        # 设置变量时已满足。


        # # (2) 水浇地每年可以种植一季水稻或两季蔬菜 1
        # for season in ['第一季', '第二季']:
        #     model += pulp.lpSum([
        #         decision_vars.get(('水浇地', crop, season), 0)
        #         for crop in fields[(fields['地块类型'] == '水浇地') & (fields['季次'] == season)]['作物名称']
        #     ]) <= 1
        #
        # # (3) 普通大棚每年种植一季蔬菜和一季食用菌 1
        # for season in ['第一季', '第二季']:
        #     model += pulp.lpSum([
        #         decision_vars.get(('普通大棚', crop, season), 0)
        #         for crop in fields[(fields['地块类型'] == '普通大棚') & (fields['季次'] == season)]['作物名称']
        #     ]) <= 2  # 至少一季蔬菜和一季食用菌
        #
        # # (4) 智慧大棚每年种植两季蔬菜（大白菜、白萝卜和红萝卜除外） 1
        # for season in ['第一季', '第二季']:
        #     model += pulp.lpSum([
        #         decision_vars.get(('智慧大棚', crop, season), 0)
        #         for crop in fields[(fields['地块类型'] == '智慧大棚') & (fields['季次'] == season)]['作物名称']
        #         if crop not in ['大白菜', '白萝卜', '红萝卜']
        #     ]) <= 2  # 每季最多种植两种蔬菜
        #
        # (8) 大白菜、白萝卜和红萝卜只能在水浇地的第二季种植 1
        # for crop in ['大白菜', '白萝卜', '红萝卜']:
        #     for season in ['第一季','单季']:
        #         model += pulp.lpSum([
        #             decision_vars.get(('水浇地', crop, season,year), 0)
        #         ]) == 0

        # # (9) 普通大棚每年种植两季作物，第一季可种植多种蔬菜（大白菜、白萝卜和红萝卜除外），第二季只能种植食用菌
        # for season in ['第一季', '第二季']:
        #     model += pulp.lpSum([
        #         decision_vars.get(('普通大棚', crop, season), 0)
        #         for crop in fields[(fields['地块类型'] == '普通大棚') & (fields['季次'] == season)]['作物名称']
        #         if season == '第一季' and crop not in ['大白菜', '白萝卜', '红萝卜']
        #     ]) >= 1  # 第一季至少种植一种蔬菜
        #     model += pulp.lpSum([
        #         decision_vars.get(('普通大棚', crop, season), 0)
        #         for crop in fields[(fields['地块类型'] == '普通大棚') & (fields['季次'] == season)]['作物名称']
        #         if season == '第二季' and crop not in ['食用菌']
        #     ]) == 0  # 第二季只能种植食用菌
# -*- coding: utf-8 -*-
import pandas as pd

df_type2 = pd.read_excel('AI_Coding模型消费看板.xlsx', sheet_name=8, header=0)
df_type2.columns = ['姓名', '岗位名称', '周期', '使用工具', '使用反馈', 'token消耗', 'AI建议采纳率', 'AI建议采纳数', 'AI建议被采纳数', 'AI建议接受率', 'AI建议被接受数', 'AI建议被不接受数', '部门路径', '报销申请时间']

def get_dept2(path):
    if pd.isna(path):
        return '未知'
    parts = str(path).split('>')
    return parts[1].strip() if len(parts) >= 2 else '未知'

df_type2['部门'] = df_type2['部门路径'].apply(get_dept2)

# 查看"未知"的行
print("部门为'未知'的行:")
print(df_type2[df_type2['部门'] == '未知'][['姓名', '部门路径', '部门']])

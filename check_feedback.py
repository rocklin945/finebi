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

# 过滤掉错误的表头行
df_type2 = df_type2[(df_type2['姓名'] != '姓名') & (df_type2['部门路径'] != '部门路径')]

# 过滤审批中的
df_feedback = df_type2[~df_type2['使用反馈'].astype(str).str.contains('审批中', na=False)].copy()

# 按部门分组
for dept, group in df_feedback.groupby('部门'):
    print(f"\n=== {dept} ===")
    for _, row in group.iterrows():
        if pd.notna(row['使用反馈']) and str(row['使用反馈']).strip():
            fb = str(row['使用反馈'])[:50].replace('\n', ' ')
            print(f"  {row['姓名']}：{fb}...")
        else:
            print(f"  {row['姓名']}：(空反馈)")

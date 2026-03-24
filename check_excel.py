# -*- coding: utf-8 -*-
import pandas as pd

# 读取生成的Excel
df_pending = pd.read_excel('output/2.3_审批中尚未反馈人员.xlsx')
df_feedback = pd.read_excel('output/2.4_反馈意见汇总.xlsx')

print(f"审批中Excel行数: {len(df_pending)}")
print(f"反馈汇总Excel行数: {len(df_feedback)}")

print("\n=== 审批中Excel ===")
for _, row in df_pending.iterrows():
    print(f"{row['姓名']}\t{row['岗位名称']}\t{row['部门']}\t{row['使用反馈']}")

print("\n=== 反馈汇总Excel ===")
for _, row in df_feedback.iterrows():
    print(f"\n{row['部门']}:")
    feedbacks = str(row['反馈意见汇总']).split('\n')
    for fb in feedbacks:
        if fb.strip():
            print(f"  {fb[:80]}")

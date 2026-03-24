# -*- coding: utf-8 -*-
import pandas as pd

df = pd.read_excel('output/2.4_反馈意见汇总.xlsx')
for _, row in df.iterrows():
    print(f"\n=== {row['部门']} ===")
    feedbacks = str(row['反馈意见汇总']).split('\n')
    print(f"反馈条数: {len(feedbacks)}")
    for fb in feedbacks:
        if fb.strip():
            print(f"  {fb[:60]}")

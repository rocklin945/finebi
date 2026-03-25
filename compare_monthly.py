# -*- coding: utf-8 -*-
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import numpy as np
import os

# 设置中文字体
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
matplotlib.rcParams['axes.unicode_minus'] = False

# 输出目录
output_dir = 'output'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# ========== 读取数据 ==========
df_type1 = pd.read_excel('AI_Coding模型消费看板.xlsx', sheet_name=7)
df_type1.columns = ['花名', '岗位名称', '使用时间', '使用模型', 'token消耗', '输入token', '输出token', '花费', '部门路径']

df_type2 = pd.read_excel('AI_Coding模型消费看板.xlsx', sheet_name=8, header=None)
col_names = ['花名', '岗位名称', '周期', '使用工具', '使用反馈', 'token消耗', '代码采纳率', 'AI建议采纳数', 'AI建议被采纳数', 'AI建议接受率', 'AI建议被接受数', 'AI建议被不接受数', '部门路径', '报销申请时间']
df_type2.columns = col_names
df_type2 = df_type2[df_type2['花名'] != '花名'].copy()

# 读取满意度数据
df_satisfaction = pd.read_csv('querydata.csv')

def get_dept(path):
    if pd.isna(path):
        return '未知'
    parts = str(path).split('>')
    return parts[1].strip() if len(parts) >= 2 else '未知'

df_type1['部门'] = df_type1['部门路径'].apply(get_dept)
df_type2['部门'] = df_type2['部门路径'].apply(get_dept)

# 转换日期
df_type1['使用时间'] = pd.to_datetime(df_type1['使用时间'])
df_type1['月份'] = df_type1['使用时间'].dt.month

# 转换数值列
for col in ['token消耗', '花费', '代码采纳率', 'AI建议采纳数', 'AI建议被采纳数', 'AI建议接受率', 'AI建议被接受数']:
    if col in df_type2.columns:
        df_type2[col] = pd.to_numeric(df_type2[col], errors='coerce').fillna(0)

# ========== 读取部门人数 ==========
df_dept = pd.read_excel('产研人员分布情况（职能维度）.xlsx', sheet_name=0)
df_dept.columns = ['团队', '姓名', '岗位职级', '岗位名称', '职能序列', '一级部门', '部门', 'HC']
df_dept['团队'] = df_dept['团队'].ffill()
df_dept_valid = df_dept[df_dept['姓名'].notna()].copy()
dept_person_count = df_dept_valid.groupby('部门').size().reset_index(name='部门人数')
total_person_count = len(df_dept_valid)

# ========== 2月份和3月份数据 ==========
df_type1_feb = df_type1[df_type1['月份'] == 2].copy()
df_type1_mar = df_type1[df_type1['月份'] == 3].copy()

# AI-IDE订阅数据用周期开始日期的月份来判断
def get_cycle_start_month(cycle):
    try:
        start = str(cycle).split('至')[0]
        return int(start.split('-')[1])
    except:
        return 0

df_type2['开始月份'] = df_type2['周期'].apply(get_cycle_start_month)
df_type2_feb = df_type2[df_type2['开始月份'] == 2].copy()
df_type2_mar = df_type2[df_type2['开始月份'] == 3].copy()

# ========== 1. Token消耗量对比 ==========
feb_token = df_type1_feb.groupby('部门')['token消耗'].sum().reset_index(name='2月Token消耗')
mar_token = df_type1_mar.groupby('部门')['token消耗'].sum().reset_index(name='3月Token消耗')

token_compare = pd.merge(feb_token, mar_token, on='部门', how='outer').fillna(0)
token_compare = pd.merge(dept_person_count, token_compare, on='部门', how='left').fillna(0)
token_compare['2月Token消耗'] = token_compare['2月Token消耗'].astype(int)
token_compare['3月Token消耗'] = token_compare['3月Token消耗'].astype(int)
token_compare['环比增长率(%)'] = (((token_compare['3月Token消耗'] - token_compare['2月Token消耗']) / token_compare['2月Token消耗'].replace(0, 1)) * 100).round(1)

feb_total = int(df_type1_feb['token消耗'].sum())
mar_total = int(df_type1_mar['token消耗'].sum())
total_row = pd.DataFrame([{
    '部门': '总计',
    '部门人数': total_person_count,
    '2月Token消耗': feb_total,
    '3月Token消耗': mar_total,
    '环比增长率(%)': round(((mar_total - feb_total) / feb_total * 100), 1) if feb_total > 0 else 0
}])
token_compare = pd.concat([total_row, token_compare], ignore_index=True)

fig, ax1 = plt.subplots(figsize=(14, 8))
x = range(len(token_compare))
width = 0.35
bars1 = ax1.bar([i - width/2 for i in x], token_compare['2月Token消耗'], width, label='2月', color='#6CB0F5', edgecolor='white')
bars2 = ax1.bar([i + width/2 for i in x], token_compare['3月Token消耗'], width, label='3月', color='#E8893C', edgecolor='white')
ax1.set_ylabel('Token消耗量', fontsize=11)
ax1.set_title('各部门Token消耗量对比（2月 vs 3月）', fontsize=14, fontweight='bold', pad=20)
ax1.legend(loc='upper left', frameon=False)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('white')
for i, (idx, row) in enumerate(token_compare.iterrows()):
    ax1.text(i - width/2, row['2月Token消耗'] + 5000000, f'{int(row["2月Token消耗"]/1000000):,}M', ha='center', va='bottom', fontsize=7, color='#6CB0F5')
    ax1.text(i + width/2, row['3月Token消耗'] + 5000000, f'{int(row["3月Token消耗"]/1000000):,}M', ha='center', va='bottom', fontsize=7, color='#E8893C')

ax2 = ax1.twinx()
pcts = token_compare['环比增长率(%)'].values
ax2.plot(x, pcts, 'o-', color='black', linewidth=2, markersize=8, markerfacecolor='white', markeredgecolor='black', markeredgewidth=2)
ax2.set_ylabel('环比增长率(%)', fontsize=11)
ax2.spines['top'].set_visible(False)
ax2.spines['left'].set_visible(False)
ax2.set_facecolor('none')
for i, pct in enumerate(pcts):
    ax2.annotate(f'{pct}%', (i, pct), textcoords="offset points", xytext=(0, 15), ha='center', fontsize=10, color='black', fontweight='bold')
ax2.axhline(y=0, color='gray', linestyle='--', linewidth=0.8)
ax1.set_xticks(x)
ax1.set_xticklabels(token_compare['部门'], rotation=45, ha='right', fontsize=10)
plt.tight_layout()
plt.subplots_adjust(bottom=0.15, top=0.88)
plt.savefig(f'{output_dir}/3.1_Token消耗量对比.png', dpi=150, facecolor='white')
plt.close()

# ========== 2. Token花费对比 ==========
feb_cost = df_type1_feb.groupby('部门')['花费'].sum().reset_index(name='2月Token花费($)')
mar_cost = df_type1_mar.groupby('部门')['花费'].sum().reset_index(name='3月Token花费($)')

cost_compare = pd.merge(feb_cost, mar_cost, on='部门', how='outer').fillna(0)
cost_compare = pd.merge(dept_person_count, cost_compare, on='部门', how='left').fillna(0)
cost_compare['2月Token花费($)'] = cost_compare['2月Token花费($)'].round(2)
cost_compare['3月Token花费($)'] = cost_compare['3月Token花费($)'].round(2)
cost_compare['环比增长率(%)'] = (((cost_compare['3月Token花费($)'] - cost_compare['2月Token花费($)']) / cost_compare['2月Token花费($)'].replace(0, 1)) * 100).round(1)

feb_cost_total = round(df_type1_feb['花费'].sum(), 2)
mar_cost_total = round(df_type1_mar['花费'].sum(), 2)
total_row = pd.DataFrame([{
    '部门': '总计',
    '部门人数': total_person_count,
    '2月Token花费($)': feb_cost_total,
    '3月Token花费($)': mar_cost_total,
    '环比增长率(%)': round(((mar_cost_total - feb_cost_total) / feb_cost_total * 100), 1) if feb_cost_total > 0 else 0
}])
cost_compare = pd.concat([total_row, cost_compare], ignore_index=True)

fig, ax1 = plt.subplots(figsize=(14, 8))
x = range(len(cost_compare))
bars1 = ax1.bar([i - width/2 for i in x], cost_compare['2月Token花费($)'], width, label='2月', color='#6CB0F5', edgecolor='white')
bars2 = ax1.bar([i + width/2 for i in x], cost_compare['3月Token花费($)'], width, label='3月', color='#E8893C', edgecolor='white')
ax1.set_ylabel('Token花费 ($)', fontsize=11)
ax1.set_title('各部门Token花费对比（2月 vs 3月）', fontsize=14, fontweight='bold', pad=20)
ax1.legend(loc='upper left', frameon=False)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('white')
for i, (idx, row) in enumerate(cost_compare.iterrows()):
    ax1.text(i - width/2, row['2月Token花费($)'] + 1, f'${row["2月Token花费($)"]:.1f}', ha='center', va='bottom', fontsize=8, color='#6CB0F5')
    ax1.text(i + width/2, row['3月Token花费($)'] + 1, f'${row["3月Token花费($)"]:.1f}', ha='center', va='bottom', fontsize=8, color='#E8893C')

ax2 = ax1.twinx()
pcts = cost_compare['环比增长率(%)'].values
ax2.plot(x, pcts, 'o-', color='black', linewidth=2, markersize=8, markerfacecolor='white', markeredgecolor='black', markeredgewidth=2)
ax2.set_ylabel('环比增长率(%)', fontsize=11)
ax2.spines['top'].set_visible(False)
ax2.spines['left'].set_visible(False)
ax2.set_facecolor('none')
for i, pct in enumerate(pcts):
    ax2.annotate(f'{pct}%', (i, pct), textcoords="offset points", xytext=(0, 15), ha='center', fontsize=10, color='black', fontweight='bold')
ax2.axhline(y=0, color='gray', linestyle='--', linewidth=0.8)
ax1.set_xticks(x)
ax1.set_xticklabels(cost_compare['部门'], rotation=45, ha='right', fontsize=10)
plt.tight_layout()
plt.subplots_adjust(bottom=0.15, top=0.88)
plt.savefig(f'{output_dir}/3.2_Token花费对比.png', dpi=150, facecolor='white')
plt.close()

# ========== 3. Token使用人数对比 ==========
feb_users = df_type1_feb[df_type1_feb['token消耗'] > 0].groupby('部门')['花名'].nunique().reset_index(name='2月Token使用人数')
mar_users = df_type1_mar[df_type1_mar['token消耗'] > 0].groupby('部门')['花名'].nunique().reset_index(name='3月Token使用人数')

users_compare = pd.merge(feb_users, mar_users, on='部门', how='outer').fillna(0)
users_compare = pd.merge(dept_person_count, users_compare, on='部门', how='left').fillna(0)
users_compare['2月Token使用人数'] = users_compare['2月Token使用人数'].astype(int)
users_compare['3月Token使用人数'] = users_compare['3月Token使用人数'].astype(int)
users_compare['环比增长率(%)'] = (((users_compare['3月Token使用人数'] - users_compare['2月Token使用人数']) / users_compare['2月Token使用人数'].replace(0, 1)) * 100).round(1)

feb_users_total = int(df_type1_feb[df_type1_feb['token消耗'] > 0]['花名'].nunique())
mar_users_total = int(df_type1_mar[df_type1_mar['token消耗'] > 0]['花名'].nunique())
total_row = pd.DataFrame([{
    '部门': '总计',
    '部门人数': total_person_count,
    '2月Token使用人数': feb_users_total,
    '3月Token使用人数': mar_users_total,
    '环比增长率(%)': round(((mar_users_total - feb_users_total) / feb_users_total * 100), 1) if feb_users_total > 0 else 0
}])
users_compare = pd.concat([total_row, users_compare], ignore_index=True)

fig, ax1 = plt.subplots(figsize=(14, 8))
x = range(len(users_compare))
bars1 = ax1.bar([i - width/2 for i in x], users_compare['2月Token使用人数'], width, label='2月', color='#6CB0F5', edgecolor='white')
bars2 = ax1.bar([i + width/2 for i in x], users_compare['3月Token使用人数'], width, label='3月', color='#E8893C', edgecolor='white')
ax1.set_ylabel('使用人数', fontsize=11)
ax1.set_title('各部门Token使用人数对比（2月 vs 3月）', fontsize=14, fontweight='bold', pad=20)
ax1.legend(loc='upper left', frameon=False)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('white')
for i, (idx, row) in enumerate(users_compare.iterrows()):
    ax1.text(i - width/2, row['2月Token使用人数'] + 0.5, f'{int(row["2月Token使用人数"])}', ha='center', va='bottom', fontsize=9, color='#6CB0F5')
    ax1.text(i + width/2, row['3月Token使用人数'] + 0.5, f'{int(row["3月Token使用人数"])}', ha='center', va='bottom', fontsize=9, color='#E8893C')

ax2 = ax1.twinx()
pcts = users_compare['环比增长率(%)'].values
ax2.plot(x, pcts, 'o-', color='black', linewidth=2, markersize=8, markerfacecolor='white', markeredgecolor='black', markeredgewidth=2)
ax2.set_ylabel('环比增长率(%)', fontsize=11)
ax2.spines['top'].set_visible(False)
ax2.spines['left'].set_visible(False)
ax2.set_facecolor('none')
for i, pct in enumerate(pcts):
    ax2.annotate(f'{pct}%', (i, pct), textcoords="offset points", xytext=(0, 15), ha='center', fontsize=10, color='black', fontweight='bold')
ax2.axhline(y=0, color='gray', linestyle='--', linewidth=0.8)
ax1.set_xticks(x)
ax1.set_xticklabels(users_compare['部门'], rotation=45, ha='right', fontsize=10)
plt.tight_layout()
plt.subplots_adjust(bottom=0.15, top=0.88)
plt.savefig(f'{output_dir}/3.3_Token使用人数对比.png', dpi=150, facecolor='white')
plt.close()

# ========== 4. AI-IDE订阅人数对比 ==========
feb_ai_users = df_type2_feb.groupby('部门')['花名'].nunique().reset_index(name='2月AI-IDE订阅人数')
mar_ai_users = df_type2_mar.groupby('部门')['花名'].nunique().reset_index(name='3月AI-IDE订阅人数')

ai_users_compare = pd.merge(feb_ai_users, mar_ai_users, on='部门', how='outer').fillna(0)
ai_users_compare = pd.merge(dept_person_count, ai_users_compare, on='部门', how='left').fillna(0)
ai_users_compare['2月AI-IDE订阅人数'] = ai_users_compare['2月AI-IDE订阅人数'].astype(int)
ai_users_compare['3月AI-IDE订阅人数'] = ai_users_compare['3月AI-IDE订阅人数'].astype(int)
ai_users_compare['环比增长率(%)'] = (((ai_users_compare['3月AI-IDE订阅人数'] - ai_users_compare['2月AI-IDE订阅人数']) / ai_users_compare['2月AI-IDE订阅人数'].replace(0, 1)) * 100).round(1)

feb_ai_total = int(df_type2_feb['花名'].nunique())
mar_ai_total = int(df_type2_mar['花名'].nunique())
total_row = pd.DataFrame([{
    '部门': '总计',
    '部门人数': total_person_count,
    '2月AI-IDE订阅人数': feb_ai_total,
    '3月AI-IDE订阅人数': mar_ai_total,
    '环比增长率(%)': round(((mar_ai_total - feb_ai_total) / feb_ai_total * 100), 1) if feb_ai_total > 0 else 0
}])
ai_users_compare = pd.concat([total_row, ai_users_compare], ignore_index=True)

fig, ax1 = plt.subplots(figsize=(14, 8))
x = range(len(ai_users_compare))
bars1 = ax1.bar([i - width/2 for i in x], ai_users_compare['2月AI-IDE订阅人数'], width, label='2月', color='#6CB0F5', edgecolor='white')
bars2 = ax1.bar([i + width/2 for i in x], ai_users_compare['3月AI-IDE订阅人数'], width, label='3月', color='#E8893C', edgecolor='white')
ax1.set_ylabel('订阅人数', fontsize=11)
ax1.set_title('各部门AI-IDE订阅人数对比（2月 vs 3月）', fontsize=14, fontweight='bold', pad=20)
ax1.legend(loc='upper left', frameon=False)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('white')
for i, (idx, row) in enumerate(ai_users_compare.iterrows()):
    ax1.text(i - width/2, row['2月AI-IDE订阅人数'] + 0.3, f'{int(row["2月AI-IDE订阅人数"])}', ha='center', va='bottom', fontsize=9, color='#6CB0F5')
    ax1.text(i + width/2, row['3月AI-IDE订阅人数'] + 0.3, f'{int(row["3月AI-IDE订阅人数"])}', ha='center', va='bottom', fontsize=9, color='#E8893C')

ax2 = ax1.twinx()
pcts = ai_users_compare['环比增长率(%)'].values
ax2.plot(x, pcts, 'o-', color='black', linewidth=2, markersize=8, markerfacecolor='white', markeredgecolor='black', markeredgewidth=2)
ax2.set_ylabel('环比增长率(%)', fontsize=11)
ax2.spines['top'].set_visible(False)
ax2.spines['left'].set_visible(False)
ax2.set_facecolor('none')
for i, pct in enumerate(pcts):
    ax2.annotate(f'{pct}%', (i, pct), textcoords="offset points", xytext=(0, 15), ha='center', fontsize=10, color='black', fontweight='bold')
ax2.axhline(y=0, color='gray', linestyle='--', linewidth=0.8)
ax1.set_xticks(x)
ax1.set_xticklabels(ai_users_compare['部门'], rotation=45, ha='right', fontsize=10)
plt.tight_layout()
plt.subplots_adjust(bottom=0.15, top=0.88)
plt.savefig(f'{output_dir}/3.4_AI-IDE订阅人数对比.png', dpi=150, facecolor='white')
plt.close()

# ========== 5. AI-IDE平均满意度对比 ==========
# 通过人名从querydata.csv获取满意度分数，然后按部门汇总
# 先获取2月和3月的订阅用户
feb_users_set = set(df_type2_feb['花名'].unique())
mar_users_set = set(df_type2_mar['花名'].unique())

# 关联满意度数据
df_satisfaction_feb = df_satisfaction[df_satisfaction['name'].isin(feb_users_set)].copy()
df_satisfaction_mar = df_satisfaction[df_satisfaction['name'].isin(mar_users_set)].copy()

# 按部门汇总满意度
def get_user_dept(name, df_type2):
    user_data = df_type2[df_type2['花名'] == name]
    if len(user_data) > 0:
        return user_data.iloc[0]['部门']
    return '未知'

df_satisfaction_feb['部门'] = df_satisfaction_feb['name'].apply(lambda x: get_user_dept(x, df_type2_feb))
df_satisfaction_mar['部门'] = df_satisfaction_mar['name'].apply(lambda x: get_user_dept(x, df_type2_mar))

# 计算各部门平均满意度
feb_sat = df_satisfaction_feb.groupby('部门')['satisfaction_score'].mean().reset_index(name='2月满意度')
mar_sat = df_satisfaction_mar.groupby('部门')['satisfaction_score'].mean().reset_index(name='3月满意度')

satisfaction_compare = pd.merge(feb_sat, mar_sat, on='部门', how='outer').fillna(0)
satisfaction_compare = pd.merge(dept_person_count, satisfaction_compare, on='部门', how='left').fillna(0)
satisfaction_compare['2月满意度'] = satisfaction_compare['2月满意度'].round(1)
satisfaction_compare['3月满意度'] = satisfaction_compare['3月满意度'].round(1)
# 只有2月和3月都有数据的部门才计算变化率
satisfaction_compare['满意度变化率(%)'] = satisfaction_compare.apply(
    lambda row: round(((row['3月满意度'] - row['2月满意度']) / row['2月满意度'] * 100), 1) if row['2月满意度'] > 0 and row['3月满意度'] > 0 else 0, axis=1
)

# 添加总计行
feb_total_sat = df_satisfaction_feb['satisfaction_score'].mean()
mar_total_sat = df_satisfaction_mar['satisfaction_score'].mean() if len(df_satisfaction_mar) > 0 else float('nan')

# 计算变化率
if pd.notna(feb_total_sat) and pd.notna(mar_total_sat) and feb_total_sat > 0 and mar_total_sat > 0:
    change_rate = round(((mar_total_sat - feb_total_sat) / feb_total_sat * 100), 1)
else:
    change_rate = 0.0

total_row = pd.DataFrame([{
    '部门': '总计',
    '部门人数': total_person_count,
    '2月满意度': round(feb_total_sat, 1) if pd.notna(feb_total_sat) else 0,
    '3月满意度': round(mar_total_sat, 1) if pd.notna(mar_total_sat) else 0,
    '满意度变化率(%)': change_rate
}])
satisfaction_compare = pd.concat([total_row, satisfaction_compare], ignore_index=True)

fig, ax1 = plt.subplots(figsize=(14, 8))
x = range(len(satisfaction_compare))
bars1 = ax1.bar([i - width/2 for i in x], satisfaction_compare['2月满意度'], width, label='2月', color='#6CB0F5', edgecolor='white')
bars2 = ax1.bar([i + width/2 for i in x], satisfaction_compare['3月满意度'], width, label='3月', color='#E8893C', edgecolor='white')
ax1.set_ylabel('满意度得分', fontsize=11)
ax1.set_title('各部门AI-IDE平均满意度对比（2月 vs 3月）', fontsize=14, fontweight='bold', pad=20)
ax1.legend(loc='upper left', frameon=False)
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.set_facecolor('#FAFAFA')
ax1.set_ylim(0, 6)
fig.patch.set_facecolor('white')
for i, (idx, row) in enumerate(satisfaction_compare.iterrows()):
    val2 = row['2月满意度'] if pd.notna(row['2月满意度']) else 0
    val3 = row['3月满意度'] if pd.notna(row['3月满意度']) else 0
    ax1.text(i - width/2, val2 + 0.1, f'{val2:.1f}', ha='center', va='bottom', fontsize=9, color='#6CB0F5')
    ax1.text(i + width/2, val3 + 0.1, f'{val3:.1f}', ha='center', va='bottom', fontsize=9, color='#E8893C')

ax2 = ax1.twinx()
pcts = satisfaction_compare['满意度变化率(%)'].fillna(0).values
ax2.plot(x, pcts, 'o-', color='black', linewidth=2, markersize=8, markerfacecolor='white', markeredgecolor='black', markeredgewidth=2)
ax2.set_ylabel('满意度变化率(%)', fontsize=11)
ax2.spines['top'].set_visible(False)
ax2.spines['left'].set_visible(False)
ax2.set_facecolor('none')
for i, pct in enumerate(pcts):
    if pd.notna(pct):
        ax2.annotate(f'{pct}%', (i, pct), textcoords="offset points", xytext=(0, 15), ha='center', fontsize=10, color='black', fontweight='bold')
ax2.axhline(y=0, color='gray', linestyle='--', linewidth=0.8)
ax1.set_xticks(x)
ax1.set_xticklabels(satisfaction_compare['部门'], rotation=45, ha='right', fontsize=10)
plt.tight_layout()
plt.subplots_adjust(bottom=0.15, top=0.88)
plt.savefig(f'{output_dir}/3.5_AI-IDE平均满意度对比.png', dpi=150, facecolor='white')
plt.close()

# ========== 6. 场景发布饼图 ==========
# 从querydata.csv获取场景分布
df_scenario = df_satisfaction[df_satisfaction['scenario_distribution'].notna()].copy()

# 解析场景（逗号分隔）
scenario_counts = {}
for scenarios in df_scenario['scenario_distribution']:
    for scenario in str(scenarios).split(','):
        scenario = scenario.strip()
        if scenario:
            scenario_counts[scenario] = scenario_counts.get(scenario, 0) + 1

scenario_df = pd.DataFrame(list(scenario_counts.items()), columns=['场景', '人数'])
scenario_df = scenario_df.sort_values('人数', ascending=False)

print("场景发布统计:")
print(scenario_df)

fig, ax = plt.subplots(figsize=(10, 8))
colors = ['#4A90D9', '#E8893C', '#5BA0E8', '#D07030', '#6CB0F5', '#F5A060', '#8DC0FF', '#E8A060']
wedges, texts, autotexts = ax.pie(scenario_df['人数'], labels=scenario_df['场景'], autopct='%1.1f%%',
                                   colors=colors[:len(scenario_df)], startangle=90,
                                   wedgeprops={'edgecolor': 'white', 'linewidth': 2})
for text in texts:
    text.set_fontsize(12)
for autotext in autotexts:
    autotext.set_fontsize(10)
    autotext.set_color('white')
    autotext.set_fontweight('bold')
ax.set_title('AI-IDE场景发布分布', fontsize=14, fontweight='bold', pad=20)
plt.tight_layout()
plt.savefig(f'{output_dir}/3.6_场景发布分布.png', dpi=150, facecolor='white')
plt.close()

# ========== 保存汇总数据到Excel ==========
with pd.ExcelWriter(f'{output_dir}/3_月度对比汇总.xlsx') as writer:
    token_compare.to_excel(writer, sheet_name='Token消耗量', index=False)
    cost_compare.to_excel(writer, sheet_name='Token花费', index=False)
    users_compare.to_excel(writer, sheet_name='Token使用人数', index=False)
    ai_users_compare.to_excel(writer, sheet_name='AI-IDE订阅人数', index=False)
    satisfaction_compare.to_excel(writer, sheet_name='AI-IDE满意度', index=False)
    scenario_df.to_excel(writer, sheet_name='场景发布', index=False)

print("\n=== 所有对比图表生成完成 ===")
print(f"输出目录: {output_dir}")

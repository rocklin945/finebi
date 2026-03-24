# -*- coding: utf-8 -*-
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
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
df_type2.columns = ['花名', '岗位名称', '周期', '使用工具', '使用反馈', 'Unnamed', 'token消耗', '代码采纳率', 'AI建议采纳数', 'AI建议被采纳数', 'AI建议接受率', 'AI建议被接受数', 'AI建议被不接受数', '部门路径'][:len(df_type2.columns)]
df_type2 = df_type2[df_type2['花名'] != '花名'].copy()

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
for col in ['token消耗', '花费', '代码采纳率', 'AI建议采纳数', 'AI建议被采纳数', 'AI建议接受率', 'AI建议被接受数', 'AI建议被不接受数']:
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
df_type2_feb = df_type2[df_type2['周期'].astype(str).str.contains('2026-02', na=False)].copy()
df_type2_mar = df_type2[df_type2['周期'].astype(str).str.contains('2026-03', na=False)].copy()

# ========== 1. Token消耗量对比 ==========
feb_token = df_type1_feb.groupby('部门')['token消耗'].sum().reset_index(name='2月Token消耗')
mar_token = df_type1_mar.groupby('部门')['token消耗'].sum().reset_index(name='3月Token消耗')

token_compare = pd.merge(feb_token, mar_token, on='部门', how='outer').fillna(0)
token_compare = pd.merge(dept_person_count, token_compare, on='部门', how='left').fillna(0)
token_compare['2月Token消耗'] = token_compare['2月Token消耗'].astype(int)
token_compare['3月Token消耗'] = token_compare['3月Token消耗'].astype(int)
token_compare['环比增长'] = token_compare['3月Token消耗'] - token_compare['2月Token消耗']
token_compare['环比增长率(%)'] = ((token_compare['3月Token消耗'] / token_compare['2月Token消耗'].replace(0, 1)) * 100).round(1)

# 添加总计行
feb_total = int(df_type1_feb['token消耗'].sum())
mar_total = int(df_type1_mar['token消耗'].sum())
total_row = pd.DataFrame([{
    '部门': '总计',
    '部门人数': total_person_count,
    '2月Token消耗': feb_total,
    '3月Token消耗': mar_total,
    '环比增长': mar_total - feb_total,
    '环比增长率(%)': round((mar_total / feb_total * 100), 1) if feb_total > 0 else 0
}])
token_compare = pd.concat([token_compare, total_row], ignore_index=True)
token_compare = token_compare.sort_values('3月Token消耗', ascending=False)

# 柱状图
fig, ax = plt.subplots(figsize=(14, 6))
x = range(len(token_compare))
width = 0.35
bars1 = ax.bar([i - width/2 for i in x], token_compare['2月Token消耗'], width, label='2月', color='#6CB0F5', edgecolor='white')
bars2 = ax.bar([i + width/2 for i in x], token_compare['3月Token消耗'], width, label='3月', color='#E8893C', edgecolor='white')
ax.set_xticks(x)
ax.set_xticklabels(token_compare['部门'], rotation=45, ha='right', fontsize=10)
ax.set_ylabel('Token消耗量', fontsize=11)
ax.set_title('各部门Token消耗量对比（2月 vs 3月）', fontsize=14, fontweight='bold', pad=15)
ax.legend(loc='upper right', frameon=False)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('white')
# 添加数值和百分比标签
for i, (idx, row) in enumerate(token_compare.iterrows()):
    ax.text(i - width/2, row['2月Token消耗'] + 5000000, f'{int(row["2月Token消耗"]/1000000):,}M', ha='center', va='bottom', fontsize=7, color='#6CB0F5')
    ax.text(i + width/2, row['3月Token消耗'] + 5000000, f'{int(row["3月Token消耗"]/1000000):,}M', ha='center', va='bottom', fontsize=7, color='#E8893C')
    # 百分比变化标注在3月柱子上方
    pct = row['环比增长率(%)']
    if pct > 0:
        ax.text(i + width/2, row['3月Token消耗'] + 20000000, f'▲{pct}%', ha='center', va='bottom', fontsize=8, color='#28A745', fontweight='bold')
    elif pct < 0:
        ax.text(i + width/2, row['3月Token消耗'] + 20000000, f'▼{abs(pct)}%', ha='center', va='bottom', fontsize=8, color='#DC3545', fontweight='bold')
plt.tight_layout()
plt.savefig(f'{output_dir}/3.1_Token消耗量对比.png', dpi=150, facecolor='white')
plt.close()

# ========== 2. Token花费对比 ==========
feb_cost = df_type1_feb.groupby('部门')['花费'].sum().reset_index(name='2月Token花费($)')
mar_cost = df_type1_mar.groupby('部门')['花费'].sum().reset_index(name='3月Token花费($)')

cost_compare = pd.merge(feb_cost, mar_cost, on='部门', how='outer').fillna(0)
cost_compare = pd.merge(dept_person_count, cost_compare, on='部门', how='left').fillna(0)
cost_compare['2月Token花费($)'] = cost_compare['2月Token花费($)'].round(2)
cost_compare['3月Token花费($)'] = cost_compare['3月Token花费($)'].round(2)
cost_compare['环比增长($)'] = (cost_compare['3月Token花费($)'] - cost_compare['2月Token花费($)']).round(2)
cost_compare['环比增长率(%)'] = ((cost_compare['3月Token花费($)'] / cost_compare['2月Token花费($)'].replace(0, 1)) * 100).round(1)

# 添加总计行
feb_cost_total = round(df_type1_feb['花费'].sum(), 2)
mar_cost_total = round(df_type1_mar['花费'].sum(), 2)
total_row = pd.DataFrame([{
    '部门': '总计',
    '部门人数': total_person_count,
    '2月Token花费($)': feb_cost_total,
    '3月Token花费($)': mar_cost_total,
    '环比增长($)': round(mar_cost_total - feb_cost_total, 2),
    '环比增长率(%)': round((mar_cost_total / feb_cost_total * 100), 1) if feb_cost_total > 0 else 0
}])
cost_compare = pd.concat([cost_compare, total_row], ignore_index=True)
cost_compare = cost_compare.sort_values('3月Token花费($)', ascending=False)

fig, ax = plt.subplots(figsize=(14, 6))
x = range(len(cost_compare))
bars1 = ax.bar([i - width/2 for i in x], cost_compare['2月Token花费($)'], width, label='2月', color='#6CB0F5', edgecolor='white')
bars2 = ax.bar([i + width/2 for i in x], cost_compare['3月Token花费($)'], width, label='3月', color='#E8893C', edgecolor='white')
ax.set_xticks(x)
ax.set_xticklabels(cost_compare['部门'], rotation=45, ha='right', fontsize=10)
ax.set_ylabel('Token花费 ($)', fontsize=11)
ax.set_title('各部门Token花费对比（2月 vs 3月）', fontsize=14, fontweight='bold', pad=15)
ax.legend(loc='upper right', frameon=False)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('white')
for i, (idx, row) in enumerate(cost_compare.iterrows()):
    ax.text(i - width/2, row['2月Token花费($)'] + 1, f'${row["2月Token花费($)"]:.1f}', ha='center', va='bottom', fontsize=8, color='#6CB0F5')
    ax.text(i + width/2, row['3月Token花费($)'] + 1, f'${row["3月Token花费($)"]:.1f}', ha='center', va='bottom', fontsize=8, color='#E8893C')
    pct = row['环比增长率(%)']
    if pct > 0:
        ax.text(i + width/2, row['3月Token花费($)'] + 10, f'▲{pct}%', ha='center', va='bottom', fontsize=8, color='#28A745', fontweight='bold')
    elif pct < 0:
        ax.text(i + width/2, row['3月Token花费($)'] + 10, f'▼{abs(pct)}%', ha='center', va='bottom', fontsize=8, color='#DC3545', fontweight='bold')
plt.tight_layout()
plt.savefig(f'{output_dir}/3.2_Token花费对比.png', dpi=150, facecolor='white')
plt.close()

# ========== 3. Token使用人数对比 ==========
feb_users = df_type1_feb[df_type1_feb['token消耗'] > 0].groupby('部门')['花名'].nunique().reset_index(name='2月Token使用人数')
mar_users = df_type1_mar[df_type1_mar['token消耗'] > 0].groupby('部门')['花名'].nunique().reset_index(name='3月Token使用人数')

users_compare = pd.merge(feb_users, mar_users, on='部门', how='outer').fillna(0)
users_compare = pd.merge(dept_person_count, users_compare, on='部门', how='left').fillna(0)
users_compare['2月Token使用人数'] = users_compare['2月Token使用人数'].astype(int)
users_compare['3月Token使用人数'] = users_compare['3月Token使用人数'].astype(int)
users_compare['环比增长'] = users_compare['3月Token使用人数'] - users_compare['2月Token使用人数']
users_compare['环比增长率(%)'] = ((users_compare['3月Token使用人数'] / users_compare['2月Token使用人数'].replace(0, 1)) * 100).round(1)

# 添加总计行
feb_users_total = int(df_type1_feb[df_type1_feb['token消耗'] > 0]['花名'].nunique())
mar_users_total = int(df_type1_mar[df_type1_mar['token消耗'] > 0]['花名'].nunique())
total_row = pd.DataFrame([{
    '部门': '总计',
    '部门人数': total_person_count,
    '2月Token使用人数': feb_users_total,
    '3月Token使用人数': mar_users_total,
    '环比增长': mar_users_total - feb_users_total,
    '环比增长率(%)': round((mar_users_total / feb_users_total * 100), 1) if feb_users_total > 0 else 0
}])
users_compare = pd.concat([users_compare, total_row], ignore_index=True)
users_compare = users_compare.sort_values('3月Token使用人数', ascending=False)

fig, ax = plt.subplots(figsize=(14, 6))
x = range(len(users_compare))
bars1 = ax.bar([i - width/2 for i in x], users_compare['2月Token使用人数'], width, label='2月', color='#6CB0F5', edgecolor='white')
bars2 = ax.bar([i + width/2 for i in x], users_compare['3月Token使用人数'], width, label='3月', color='#E8893C', edgecolor='white')
ax.set_xticks(x)
ax.set_xticklabels(users_compare['部门'], rotation=45, ha='right', fontsize=10)
ax.set_ylabel('使用人数', fontsize=11)
ax.set_title('各部门Token使用人数对比（2月 vs 3月）', fontsize=14, fontweight='bold', pad=15)
ax.legend(loc='upper right', frameon=False)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('white')
for i, (idx, row) in enumerate(users_compare.iterrows()):
    ax.text(i - width/2, row['2月Token使用人数'] + 0.5, f'{int(row["2月Token使用人数"])}', ha='center', va='bottom', fontsize=9, color='#6CB0F5')
    ax.text(i + width/2, row['3月Token使用人数'] + 0.5, f'{int(row["3月Token使用人数"])}', ha='center', va='bottom', fontsize=9, color='#E8893C')
    pct = row['环比增长率(%)']
    if pct > 0:
        ax.text(i + width/2, row['3月Token使用人数'] + 3, f'▲{pct}%', ha='center', va='bottom', fontsize=8, color='#28A745', fontweight='bold')
    elif pct < 0:
        ax.text(i + width/2, row['3月Token使用人数'] + 3, f'▼{abs(pct)}%', ha='center', va='bottom', fontsize=8, color='#DC3545', fontweight='bold')
plt.tight_layout()
plt.savefig(f'{output_dir}/3.3_Token使用人数对比.png', dpi=150, facecolor='white')
plt.close()

# ========== 4. AI-IDE订阅人数对比 ==========
feb_ai_users = df_type2_feb.groupby('部门')['花名'].nunique().reset_index(name='2月AI-IDE订阅人数')
mar_ai_users = df_type2_mar.groupby('部门')['花名'].nunique().reset_index(name='3月AI-IDE订阅人数')

ai_users_compare = pd.merge(feb_ai_users, mar_ai_users, on='部门', how='outer').fillna(0)
ai_users_compare = pd.merge(dept_person_count, ai_users_compare, on='部门', how='left').fillna(0)
ai_users_compare['2月AI-IDE订阅人数'] = ai_users_compare['2月AI-IDE订阅人数'].astype(int)
ai_users_compare['3月AI-IDE订阅人数'] = ai_users_compare['3月AI-IDE订阅人数'].astype(int)
ai_users_compare['环比增长'] = ai_users_compare['3月AI-IDE订阅人数'] - ai_users_compare['2月AI-IDE订阅人数']
ai_users_compare['环比增长率(%)'] = ((ai_users_compare['3月AI-IDE订阅人数'] / ai_users_compare['2月AI-IDE订阅人数'].replace(0, 1)) * 100).round(1)

# 添加总计行
feb_ai_total = int(df_type2_feb['花名'].nunique())
mar_ai_total = int(df_type2_mar['花名'].nunique())
total_row = pd.DataFrame([{
    '部门': '总计',
    '部门人数': total_person_count,
    '2月AI-IDE订阅人数': feb_ai_total,
    '3月AI-IDE订阅人数': mar_ai_total,
    '环比增长': mar_ai_total - feb_ai_total,
    '环比增长率(%)': round((mar_ai_total / feb_ai_total * 100), 1) if feb_ai_total > 0 else 0
}])
ai_users_compare = pd.concat([ai_users_compare, total_row], ignore_index=True)
ai_users_compare = ai_users_compare.sort_values('3月AI-IDE订阅人数', ascending=False)

fig, ax = plt.subplots(figsize=(14, 6))
x = range(len(ai_users_compare))
bars1 = ax.bar([i - width/2 for i in x], ai_users_compare['2月AI-IDE订阅人数'], width, label='2月', color='#6CB0F5', edgecolor='white')
bars2 = ax.bar([i + width/2 for i in x], ai_users_compare['3月AI-IDE订阅人数'], width, label='3月', color='#E8893C', edgecolor='white')
ax.set_xticks(x)
ax.set_xticklabels(ai_users_compare['部门'], rotation=45, ha='right', fontsize=10)
ax.set_ylabel('订阅人数', fontsize=11)
ax.set_title('各部门AI-IDE订阅人数对比（2月 vs 3月）', fontsize=14, fontweight='bold', pad=15)
ax.legend(loc='upper right', frameon=False)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('white')
for i, (idx, row) in enumerate(ai_users_compare.iterrows()):
    ax.text(i - width/2, row['2月AI-IDE订阅人数'] + 0.3, f'{int(row["2月AI-IDE订阅人数"])}', ha='center', va='bottom', fontsize=9, color='#6CB0F5')
    ax.text(i + width/2, row['3月AI-IDE订阅人数'] + 0.3, f'{int(row["3月AI-IDE订阅人数"])}', ha='center', va='bottom', fontsize=9, color='#E8893C')
    pct = row['环比增长率(%)']
    if pct > 0:
        ax.text(i + width/2, row['3月AI-IDE订阅人数'] + 2, f'▲{pct}%', ha='center', va='bottom', fontsize=8, color='#28A745', fontweight='bold')
    elif pct < 0:
        ax.text(i + width/2, row['3月AI-IDE订阅人数'] + 2, f'▼{abs(pct)}%', ha='center', va='bottom', fontsize=8, color='#DC3545', fontweight='bold')
plt.tight_layout()
plt.savefig(f'{output_dir}/3.4_AI-IDE订阅人数对比.png', dpi=150, facecolor='white')
plt.close()

# ========== 5. AI-IDE平均满意度对比 ==========
feb_satisfaction = df_type2_feb.groupby('部门').agg({
    '代码采纳率': 'mean',
    'AI建议接受率': 'mean'
}).reset_index()
feb_satisfaction['2月满意度'] = ((feb_satisfaction['代码采纳率'] + feb_satisfaction['AI建议接受率']) * 50).round(1)
feb_satisfaction = feb_satisfaction[['部门', '2月满意度']]

mar_satisfaction = df_type2_mar.groupby('部门').agg({
    '代码采纳率': 'mean',
    'AI建议接受率': 'mean'
}).reset_index()
mar_satisfaction['3月满意度'] = ((mar_satisfaction['代码采纳率'] + mar_satisfaction['AI建议接受率']) * 50).round(1)
mar_satisfaction = mar_satisfaction[['部门', '3月满意度']]

satisfaction_compare = pd.merge(feb_satisfaction, mar_satisfaction, on='部门', how='outer').fillna(0)
satisfaction_compare = pd.merge(dept_person_count, satisfaction_compare, on='部门', how='left').fillna(0)
satisfaction_compare['2月满意度'] = satisfaction_compare['2月满意度'].round(1)
satisfaction_compare['3月满意度'] = satisfaction_compare['3月满意度'].round(1)
satisfaction_compare['满意度变化'] = (satisfaction_compare['3月满意度'] - satisfaction_compare['2月满意度']).round(1)
satisfaction_compare['满意度变化率(%)'] = ((satisfaction_compare['3月满意度'] / satisfaction_compare['2月满意度'].replace(0, 1)) * 100 - 100).round(1)

# 添加总计行
feb_total_sat = ((df_type2_feb['代码采纳率'].mean() + df_type2_feb['AI建议接受率'].mean()) * 50)
mar_total_sat = ((df_type2_mar['代码采纳率'].mean() + df_type2_mar['AI建议接受率'].mean()) * 50)
total_row = pd.DataFrame([{
    '部门': '总计',
    '部门人数': total_person_count,
    '2月满意度': round(feb_total_sat, 1),
    '3月满意度': round(mar_total_sat, 1),
    '满意度变化': round(mar_total_sat - feb_total_sat, 1),
    '满意度变化率(%)': round(((mar_total_sat / feb_total_sat) * 100 - 100), 1) if feb_total_sat > 0 else 0
}])
satisfaction_compare = pd.concat([satisfaction_compare, total_row], ignore_index=True)
satisfaction_compare = satisfaction_compare.sort_values('3月满意度', ascending=False)

fig, ax = plt.subplots(figsize=(14, 6))
x = range(len(satisfaction_compare))
bars1 = ax.bar([i - width/2 for i in x], satisfaction_compare['2月满意度'], width, label='2月', color='#6CB0F5', edgecolor='white')
bars2 = ax.bar([i + width/2 for i in x], satisfaction_compare['3月满意度'], width, label='3月', color='#E8893C', edgecolor='white')
ax.set_xticks(x)
ax.set_xticklabels(satisfaction_compare['部门'], rotation=45, ha='right', fontsize=10)
ax.set_ylabel('满意度得分', fontsize=11)
ax.set_title('各部门AI-IDE平均满意度对比（2月 vs 3月）', fontsize=14, fontweight='bold', pad=15)
ax.legend(loc='upper right', frameon=False)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('white')
ax.set_ylim(0, 100)
for i, (idx, row) in enumerate(satisfaction_compare.iterrows()):
    ax.text(i - width/2, row['2月满意度'] + 1, f'{row["2月满意度"]:.1f}', ha='center', va='bottom', fontsize=9, color='#6CB0F5')
    ax.text(i + width/2, row['3月满意度'] + 1, f'{row["3月满意度"]:.1f}', ha='center', va='bottom', fontsize=9, color='#E8893C')
    pct = row['满意度变化率(%)']
    if pct > 0:
        ax.text(i + width/2, row['3月满意度'] + 5, f'▲{pct}%', ha='center', va='bottom', fontsize=8, color='#28A745', fontweight='bold')
    elif pct < 0:
        ax.text(i + width/2, row['3月满意度'] + 5, f'▼{abs(pct)}%', ha='center', va='bottom', fontsize=8, color='#DC3545', fontweight='bold')
plt.tight_layout()
plt.savefig(f'{output_dir}/3.5_AI-IDE平均满意度对比.png', dpi=150, facecolor='white')
plt.close()

# ========== 保存汇总数据到Excel ==========
with pd.ExcelWriter(f'{output_dir}/3_月度对比汇总.xlsx') as writer:
    token_compare.to_excel(writer, sheet_name='Token消耗量', index=False)
    cost_compare.to_excel(writer, sheet_name='Token花费', index=False)
    users_compare.to_excel(writer, sheet_name='Token使用人数', index=False)
    ai_users_compare.to_excel(writer, sheet_name='AI-IDE订阅人数', index=False)
    satisfaction_compare.to_excel(writer, sheet_name='AI-IDE满意度', index=False)

print("=== 所有对比图表生成完成 ===")
print(f"输出目录: {output_dir}")

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

# ========== 1. 读取产研人员分布情况 ==========
print("=== 1. 读取产研人员分布情况 ===")
df1 = pd.read_excel('产研人员分布情况（职能维度）.xlsx', sheet_name=0)
df1.columns = ['团队', '姓名', '岗位职级', '岗位名称', '职能序列', '一级部门', '部门', 'HC']
df1['团队'] = df1['团队'].ffill()

df1_valid = df1[df1['姓名'].notna()].copy()
dept_person_count = df1_valid.groupby('部门').size().reset_index(name='部门人数')
dept_person_count = dept_person_count.sort_values('部门人数', ascending=False)
print("部门人数统计:")
print(dept_person_count.to_string())
dept_person_count.to_excel(f'{output_dir}/部门人数统计.xlsx', index=False)

# ========== 2. 读取AI_Coding模型消费看板 ==========
print("\n=== 2. 读取AI_Coding模型消费看板 ===")
df_type1 = pd.read_excel('AI_Coding模型消费看板.xlsx', sheet_name=7)
df_type2 = pd.read_excel('AI_Coding模型消费看板.xlsx', sheet_name=8)

df_type1.columns = ['姓名', '岗位名称', '使用时间', '使用模型', 'token消耗', '输入token', '输出token', '花费', '部门路径']
df_type2.columns = ['姓名', '岗位名称', '周期', '使用工具', '使用反馈', 'token消耗', 'AI建议采纳率', 'AI建议采纳数', 'AI建议被采纳数', 'AI建议接受率', 'AI建议被接受数', 'AI建议被不接受数', '部门路径', '报销申请时间']

def get_dept2(path):
    if pd.isna(path):
        return '未知'
    parts = str(path).split('>')
    return parts[1].strip() if len(parts) >= 2 else '未知'

df_type1['部门'] = df_type1['部门路径'].apply(get_dept2)
df_type2['部门'] = df_type2['部门路径'].apply(get_dept2)

# 过滤掉错误的表头行（姓名或部门路径为表头字符串）
df_type2 = df_type2[(df_type2['姓名'] != '姓名') & (df_type2['部门路径'] != '部门路径')]

# 转换数值列
for col in ['AI建议采纳率', 'AI建议采纳数', 'AI建议被采纳数', 'AI建议接受率', 'AI建议被接受数', 'AI建议被不接受数']:
    df_type2[col] = pd.to_numeric(df_type2[col], errors='coerce').fillna(0)

# ---------- 2.1 token花费人数占部门人数的比例 ----------
print("\n--- 2.1 token花费人数占比 ---")
token_users = df_type1[df_type1['token消耗'] > 0].groupby('部门')['姓名'].nunique().reset_index(name='token花费人数')
ratio1 = pd.merge(dept_person_count, token_users, on='部门', how='left')
ratio1['token花费人数'] = ratio1['token花费人数'].fillna(0).astype(int)
ratio1['占比'] = (ratio1['token花费人数'] / ratio1['部门人数'] * 100).round(1)
ratio1 = ratio1.sort_values('占比', ascending=False)
print(ratio1.to_string())

# 柱状图 - 人数对比
fig, ax = plt.subplots(figsize=(12, 6))
x = range(len(ratio1))
width = 0.35
ax.bar([i - width/2 for i in x], ratio1['部门人数'], width, label='部门人数', color='#B0B0B0', edgecolor='white')
ax.bar([i + width/2 for i in x], ratio1['token花费人数'], width, label='token花费人数', color='#4A90D9', edgecolor='white')
ax.set_xticks(x)
ax.set_xticklabels(ratio1['部门'], rotation=45, ha='right', fontsize=10)
ax.set_ylabel('人数', fontsize=11)
ax.set_title('各部门token花费人数与部门总人数对比', fontsize=14, fontweight='bold', pad=15)
ax.legend(loc='upper right', frameon=False)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('white')
# 添加数值标签
for i, (cnt1, cnt2) in enumerate(zip(ratio1['部门人数'], ratio1['token花费人数'])):
    ax.text(i - width/2, cnt1 + 2, str(cnt1), ha='center', va='bottom', fontsize=9, color='#666666')
    ax.text(i + width/2, cnt2 + 2, str(cnt2), ha='center', va='bottom', fontsize=9, color='#4A90D9', fontweight='bold')
plt.tight_layout()
plt.savefig(f'{output_dir}/2.1a_token人数对比.png', dpi=150, facecolor='white')
plt.close()

# 横向柱状图 - 占比
fig, ax = plt.subplots(figsize=(10, 6))
ratio1_sorted = ratio1.sort_values('占比', ascending=True)
# 生成渐变颜色：占比相同的部门颜色相同
base_colors = ['#D0E8FF', '#C0E0FF', '#B0D9FF', '#A0D0FF', '#8DC0FF', '#6CB0F5', '#5BA0E8', '#4A90D9']
pct_unique = ratio1_sorted['占比'].unique()
pct_to_color = {pct: base_colors[i % len(base_colors)] for i, pct in enumerate(sorted(pct_unique))}
colors = [pct_to_color[pct] for pct in ratio1_sorted['占比']]
n = len(ratio1_sorted)
bars = ax.barh(range(n), ratio1_sorted['占比'], color=colors)
ax.set_yticks(range(n))
ax.set_yticklabels(ratio1_sorted['部门'], fontsize=10)
ax.set_xlabel('占比 (%)', fontsize=11)
ax.set_title('各部门token花费人数占本部门人数比例', fontsize=14, fontweight='bold', pad=15)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('white')
# 添加数值标签
for bar, pct in zip(bars, ratio1_sorted['占比']):
    ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height()/2, f'{pct:.1f}%', ha='left', va='center', fontsize=10)
plt.tight_layout()
plt.savefig(f'{output_dir}/2.1b_token占比.png', dpi=150, facecolor='white')
plt.close()
print(f"已保存: 2.1a_token人数对比.png, 2.1b_token占比.png")

# ---------- 2.2 AI-IDE使用人数占部门人数的比例 ----------
print("\n--- 2.2 AI-IDE使用人数占比 ---")
# 类型Ⅱ中所有订阅报销人员就是AI-IDE使用人员
ai_users = df_type2.groupby('部门')['姓名'].nunique().reset_index(name='AI-IDE使用人数')
ratio2 = pd.merge(dept_person_count, ai_users, on='部门', how='left')
ratio2['AI-IDE使用人数'] = ratio2['AI-IDE使用人数'].fillna(0).astype(int)
ratio2['占比'] = (ratio2['AI-IDE使用人数'] / ratio2['部门人数'] * 100).round(1)
ratio2 = ratio2.sort_values('占比', ascending=False)
print(ratio2.to_string())

# 柱状图 - 人数对比
fig, ax = plt.subplots(figsize=(12, 6))
x = range(len(ratio2))
width = 0.35
ax.bar([i - width/2 for i in x], ratio2['部门人数'], width, label='部门人数', color='#B0B0B0', edgecolor='white')
ax.bar([i + width/2 for i in x], ratio2['AI-IDE使用人数'], width, label='AI-IDE使用人数', color='#E8893C', edgecolor='white')
ax.set_xticks(x)
ax.set_xticklabels(ratio2['部门'], rotation=45, ha='right', fontsize=10)
ax.set_ylabel('人数', fontsize=11)
ax.set_title('各部门AI-IDE使用人数与部门总人数对比', fontsize=14, fontweight='bold', pad=15)
ax.legend(loc='upper right', frameon=False)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('white')
for i, (cnt1, cnt2) in enumerate(zip(ratio2['部门人数'], ratio2['AI-IDE使用人数'])):
    ax.text(i - width/2, cnt1 + 2, str(cnt1), ha='center', va='bottom', fontsize=9, color='#666666')
    ax.text(i + width/2, cnt2 + 2, str(cnt2), ha='center', va='bottom', fontsize=9, color='#E8893C', fontweight='bold')
plt.tight_layout()
plt.savefig(f'{output_dir}/2.2a_AI人数对比.png', dpi=150, facecolor='white')
plt.close()

# 横向柱状图 - 占比
fig, ax = plt.subplots(figsize=(10, 6))
ratio2_sorted = ratio2.sort_values('占比', ascending=True)
# 生成渐变颜色：占比相同的部门颜色相同
base_colors = ['#FFE0C8', '#FFD0B0', '#FAC0A0', '#F5B080', '#F0A060', '#E8893C', '#D07030', '#B05820']
# 按占比分配颜色，相同占比用相同颜色
pct_unique = ratio2_sorted['占比'].unique()
pct_to_color = {pct: base_colors[i % len(base_colors)] for i, pct in enumerate(sorted(pct_unique))}
colors = [pct_to_color[pct] for pct in ratio2_sorted['占比']]
n = len(ratio2_sorted)
bars = ax.barh(range(n), ratio2_sorted['占比'], color=colors)
ax.set_yticks(range(n))
ax.set_yticklabels(ratio2_sorted['部门'], fontsize=10)
ax.set_xlabel('占比 (%)', fontsize=11)
ax.set_title('各部门AI-IDE使用人数占本部门人数比例', fontsize=14, fontweight='bold', pad=15)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('white')
# 添加数值标签
for bar, pct in zip(bars, ratio2_sorted['占比']):
    ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2, f'{pct:.1f}%', ha='left', va='center', fontsize=10)
plt.tight_layout()
plt.savefig(f'{output_dir}/2.2b_AI占比.png', dpi=150, facecolor='white')
plt.close()
print(f"已保存: 2.2a_AI人数对比.png, 2.2b_AI占比.png")

# ---------- 2.3 审批中尚未反馈的人员（Excel） ----------
print("\n--- 2.3 审批中尚未反馈人员 ---")
pending = df_type2[df_type2['使用反馈'].astype(str).str.contains('审批中', na=False)]
pending_detail = pending[['姓名', '岗位名称', '部门', '使用反馈']].sort_values('部门')
pending_detail.to_excel(f'{output_dir}/2.3_审批中尚未反馈人员.xlsx', index=False)
print(f"共 {len(pending)} 人审批中尚未反馈")
print(f"已保存: 2.3_审批中尚未反馈人员.xlsx")

# ---------- 2.4 各部门反馈意见汇总（Excel） ----------
print("\n--- 2.4 各部门反馈意见汇总 ---")
# 过滤掉审批中的，只保留有效的反馈意见
df_feedback = df_type2[~df_type2['使用反馈'].astype(str).str.contains('审批中', na=False)].copy()

# 按部门分组，汇总反馈意见
dept_list = []
feedback_list = []
for dept, group in df_feedback.groupby('部门'):
    dept_list.append(dept)
    feedbacks = []
    for _, row in group.iterrows():
        if pd.notna(row['使用反馈']) and str(row['使用反馈']).strip():
            feedbacks.append(f"{row['姓名']}：{row['使用反馈']}")
    feedback_list.append('\n'.join(feedbacks))

feedback_by_dept = pd.DataFrame({'部门': dept_list, '反馈意见汇总': feedback_list})
print(f"有效反馈记录数: {len(df_feedback)}")
feedback_by_dept.to_excel(f'{output_dir}/2.4_反馈意见汇总.xlsx', index=False)
print(f"已保存: 2.4_反馈意见汇总.xlsx")

# ---------- 2.5 代码采纳率和AI建议接受率 ----------
print("\n--- 2.5 代码采纳率和AI建议接受率 ---")
rates = df_type2.groupby('部门').agg({
    'AI建议采纳率': 'mean',
    'AI建议接受率': 'mean'
}).round(2).reset_index()
rates.columns = ['部门', '代码采纳率(%)', 'AI建议接受率(%)']
rates = rates.sort_values('AI建议接受率(%)', ascending=False)
print(rates.to_string())

# 柱状图 - 比率对比
fig, ax = plt.subplots(figsize=(12, 6))
x = range(len(rates))
width = 0.35
ax.bar([i - width/2 for i in x], rates['代码采纳率(%)'], width, label='代码采纳率', color='#4A90D9', edgecolor='white')
ax.bar([i + width/2 for i in x], rates['AI建议接受率(%)'], width, label='AI建议接受率', color='#E8893C', edgecolor='white')
ax.set_xticks(x)
ax.set_xticklabels(rates['部门'], rotation=45, ha='right', fontsize=10)
ax.set_ylabel('比率 (%)', fontsize=11)
ax.set_title('各部门代码采纳率和AI建议接受率对比', fontsize=14, fontweight='bold', pad=15)
ax.legend(loc='upper right', frameon=False)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('white')
for i, (r1, r2) in enumerate(zip(rates['代码采纳率(%)'], rates['AI建议接受率(%)'])):
    ax.text(i - width/2, r1 + 1, f'{r1:.1f}%', ha='center', va='bottom', fontsize=9, color='#4A90D9')
    ax.text(i + width/2, r2 + 1, f'{r2:.1f}%', ha='center', va='bottom', fontsize=9, color='#E8893C')
plt.tight_layout()
plt.savefig(f'{output_dir}/2.5_采纳接受率对比.png', dpi=150, facecolor='white')
plt.close()
print(f"已保存: 2.5_采纳接受率对比.png")

# ---------- 2.6 各部门Token使用情况汇总表 ----------
print("\n--- 2.6 各部门Token使用情况汇总表 ---")
# 汇总token花费和token消耗
token_cost = df_type1.groupby('部门')['花费'].sum().reset_index(name='Token花费(美元)')
token_cost['Token花费(美元)'] = token_cost['Token花费(美元)'].round(2)
token_consume = df_type1.groupby('部门')['token消耗'].sum().reset_index(name='Token消耗')

# 合并Token相关数据
token_summary = pd.merge(dept_person_count, token_users, on='部门', how='left')
token_summary = pd.merge(token_summary, token_cost, on='部门', how='left')
token_summary = pd.merge(token_summary, token_consume, on='部门', how='left')
token_summary['token花费人数'] = token_summary['token花费人数'].fillna(0).astype(int)
token_summary['Token花费(美元)'] = token_summary['Token花费(美元)'].fillna(0).round(2)
token_summary['Token消耗'] = token_summary['Token消耗'].fillna(0).astype(int)

# 计算占比
token_summary['Token使用人数占比'] = (token_summary['token花费人数'] / token_summary['部门人数'] * 100).round(1)

# 按Token使用人数占比排序
token_summary = token_summary.sort_values('Token使用人数占比', ascending=False)
token_summary = token_summary.reset_index(drop=True)
token_summary.insert(0, '排名', range(1, len(token_summary) + 1))

print(token_summary[['排名', '部门', 'token花费人数', '部门人数', 'Token使用人数占比', 'Token花费(美元)', 'Token消耗']].to_string())

# 保存Token数据到Excel
token_output = token_summary[['排名', '部门', 'token花费人数', '部门人数', 'Token使用人数占比', 'Token花费(美元)', 'Token消耗']]
token_output.columns = ['排名', '部门名', 'Token使用人数', '部门人数', 'Token使用人数占比(%)', 'Token花费(美元)', 'Token消耗']
token_output.to_excel(f'{output_dir}/2.6_各部门Token使用情况汇总.xlsx', index=False)
print(f"已保存: 2.6_各部门Token使用情况汇总.xlsx")

# ---------- 2.7 各部门AI-IDE使用情况汇总表 ----------
print("\n--- 2.7 各部门AI-IDE使用情况汇总表 ---")
# 合并AI-IDE相关数据
ai_summary = pd.merge(dept_person_count, ai_users, on='部门', how='left')
ai_summary['AI-IDE使用人数'] = ai_summary['AI-IDE使用人数'].fillna(0).astype(int)

# 计算占比
ai_summary['AI-IDE使用人数占比'] = (ai_summary['AI-IDE使用人数'] / ai_summary['部门人数'] * 100).round(1)

# 按AI-IDE使用人数占比排序
ai_summary = ai_summary.sort_values('AI-IDE使用人数占比', ascending=False)
ai_summary = ai_summary.reset_index(drop=True)
ai_summary.insert(0, '排名', range(1, len(ai_summary) + 1))

print(ai_summary[['排名', '部门', 'AI-IDE使用人数', '部门人数', 'AI-IDE使用人数占比']].to_string())

# 保存AI-IDE数据到Excel
ai_output = ai_summary[['排名', '部门', 'AI-IDE使用人数', '部门人数', 'AI-IDE使用人数占比']]
ai_output.columns = ['排名', '部门名', 'AI-IDE使用人数', '部门人数', 'AI-IDE使用人数占比(%)']
ai_output.to_excel(f'{output_dir}/2.7_各部门AI-IDE使用情况汇总.xlsx', index=False)
print(f"已保存: 2.7_各部门AI-IDE使用情况汇总.xlsx")

print("\n=== 所有分析完成！===")
print(f"输出目录: {output_dir}")

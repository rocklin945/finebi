# -*- coding: utf-8 -*-
import pandas as pd

# 读取Excel
df_pending = pd.read_excel('output/2.3_审批中尚未反馈人员.xlsx')
df_feedback = pd.read_excel('output/2.4_反馈意见汇总.xlsx')

# 生成审批中表格HTML
pending_html = '<h2>4. 审批中尚未反馈人员</h2>\n<table>\n<colgroup>\n<col />\n<col />\n<col />\n<col />\n</colgroup>\n<tbody>\n<tr>\n<th>姓名</th>\n<th>岗位名称</th>\n<th>部门</th>\n<th>使用反馈</th>\n</tr>\n'
for _, row in df_pending.iterrows():
    pending_html += f'<tr>\n<td>{row["姓名"]}</td>\n<td>{row["岗位名称"]}</td>\n<td>{row["部门"]}</td>\n<td>{row["使用反馈"]}</td>\n</tr>\n'
pending_html += '</tbody>\n</table>\n'

# 生成反馈汇总表格HTML
feedback_html = '<h2>5. 各部门反馈意见汇总</h2>\n<table>\n<colgroup>\n<col style="width:150px"/>\n<col />\n</colgroup>\n<tbody>\n<tr>\n<th>部门</th>\n<th>反馈意见</th>\n</tr>\n'
for _, row in df_feedback.iterrows():
    # 将换行转换为<br/>
    feedback_content = str(row['反馈意见汇总']).replace('\n', '<br/>')
    feedback_html += f'<tr>\n<td>{row["部门"]}</td>\n<td>{feedback_content}</td>\n</tr>\n'
feedback_html += '</tbody>\n</table>\n'

# 输出到文件
with open('.devops-mcp-temp/tables.html', 'w', encoding='utf-8') as f:
    f.write(pending_html)
    f.write(feedback_html)

print("Done")

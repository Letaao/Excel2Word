import pandas as pd
from openpyxl import load_workbook

# 读取Excel文件
file_path = './data/工作簿1.xlsx'
excel_data = pd.read_excel(file_path)

# 统计子过程描述的单元格个数
result = {}
current_function = None

for index, row in excel_data.iterrows():
    if pd.notna(row['功能过程']):
        current_function = row['功能过程']
        if current_function not in result:
            result[current_function] = {'count': 0, 'rows': []}
    if pd.notna(row['子过程描述']):
        result[current_function]['count'] += 1
        result[current_function]['rows'].append(index)

# 添加统计结果到新的列
excel_data['子过程描述数'] = ""

for function, data in result.items():
    if data['rows']:
        excel_data.at[data['rows'][0], '子过程描述数'] = data['count']

excel_data2 = excel_data.drop(labels='子过程描述',axis=1)

excel_data2.dropna(subset=['功能过程'], how='all', inplace=True)


output_file_path = './data/子过程描述统计结果.xlsx'
excel_data2.to_excel(output_file_path, index=False)

output_file_path
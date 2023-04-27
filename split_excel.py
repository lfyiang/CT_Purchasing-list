import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os

# 创建GUI界面并选择原始表格文件
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(title='选择原始表格文件', filetypes=[('Excel files', '*.xlsx;*.xls')])

if not file_path:
    print('未选择任何文件！')
    exit(0)

# 读取原始表格数据
df = pd.read_excel(file_path)

# 将入库日期列转换为字符串类型并指定长度为4
df['入库日期'] = df['入库日期'].apply(lambda x: '{:04d}'.format(x))

# 按入库日期和供货商进行排序
df_sorted = df.sort_values(['入库日期', '供货商'])

# 拆分表格并保存为xls格式
if not os.path.exists('拆分表'):
    os.mkdir('拆分表')
file_list = []
for name, group in df_sorted.groupby(['入库日期', '供货商']):
    file_name = '{}_{}.xls'.format(name[0], name[1])
    file_path = os.path.join('拆分表', file_name)
    writer = pd.ExcelWriter(file_path, engine='openpyxl')
    group.to_excel(writer, index=False)
    writer._save()
    file_list.append(file_name)

# 保存生成的所有文件的文件名到文件列表表格中
pd.DataFrame(file_list, columns=['文件名']).to_excel('文件列表.xlsx', index=False)

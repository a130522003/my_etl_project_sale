import pandas as pd
import numpy as np
# 讀取 Excel 檔案
excel_file = '(DATA) 專業測驗.xlsx'
excel_data = pd.ExcelFile (excel_file)

# 取得所有工作表名稱
sheet_names = excel_data.sheet_names

# 逐一讀取並存為 CSV，並讀取 CSV 檔案
for sheet_name in sheet_names:
    df = excel_data.parse(sheet_name)
    df.to_csv(f'{sheet_name}.csv', index=False)
    
detail = pd.read_csv('detail.csv')
# print(detail.head(3))
sales_info = pd.read_csv('sales_info.csv')

# print(sales_info.head(3))
# 過濾掉測試組別的數據
# detail = detail[~detail['group'].str.contains('Test', na=False)]
# 將date列轉換為datetime類型
detail['date'] = pd.to_datetime(detail['date'])
sales_info['date'] = pd.to_datetime(sales_info['date'])

# 只保留3/1-3/5的數據
start_date = pd.to_datetime('2021-03-01')
end_date = pd.to_datetime('2021-03-05')
detail = detail[(detail['date'] >= start_date) & (detail['date'] <= end_date)]
sales_info = sales_info[(sales_info['date'] >= start_date) & (sales_info['date'] <= end_date)]
# print(sales_info.head(3))

# 合併detail和sales_info數據
merged_data = pd.merge(detail, sales_info, on=['login', 'date'])
merged_data = merged_data[~merged_data['group'].str.contains('Test', na=False)]
# print(merged_data.tail())

# 計算bonus
merged_data['bonus'] = np.where(merged_data['deposit'] > 3000, 1, 0)

# 按銷售和日期分組，計算業績
sales_performance = merged_data.groupby(['sales', 'date']).agg({
    'pnl': 'sum',
    'volume': 'sum',
    'commission': 'sum',
    'deposit': 'sum',
    'bonus': 'sum'
}).reset_index() #取消groupby分組

# 按銷售分組，計算3/1-3/5的總業績
final_performance = sales_performance.groupby('sales').agg({
    'pnl': 'sum',
    'volume': 'sum',
    'commission': 'sum',
    'deposit': 'sum',
    'bonus': 'sum'
}).reset_index()
# 顯示完整欄位與資料
pd.set_option("display.max_columns",None)
pd.set_option("display.max_rows",None)
# print(sales_performance)
# 按pnl降序排序
final_performance = final_performance.sort_values('pnl', ascending=False)
# print(final_performance.info())
#添加日期範圍列
final_performance['date'] = '2021-03-01 to 2021-03-05'

# 重新排列列的順序
columns = ['date', 'sales', 'pnl', 'volume', 'commission', 'deposit', 'bonus']
final_performance = final_performance[columns]

# 輸出結果
print(final_performance.to_string(index=False))

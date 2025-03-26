import pandas as pd
# 使用 openpyxl 引擎讀取 .xlsx 檔案
xlsx = pd.ExcelFile("StockTable.xlsx", engine='openpyxl')
df = xlsx.parse('StockTable')
# 從第 3 列開始（index=2），取出之後的資料
data = df.iloc[2:]

# 初始化儲存股票代號與名稱的列表
stock_ids = []
stock_names = []

# 假設欄位是偶數為股票代號，奇數為名稱，成對處理
for i in range(0, data.shape[1], 2):
    id_col = data.iloc[:, i]
    name_col = data.iloc[:, i+1]
    
    for stock_id, stock_name in zip(id_col, name_col):
        if pd.notna(stock_id) and pd.notna(stock_name):
            stock_ids.append(str(stock_id).strip(' '))  # 移除特殊空白字元
            stock_names.append(str(stock_name).strip(' '))

# 建立新的 DataFrame
clean_df = pd.DataFrame({
    '股票代號': stock_ids,
    '名稱': stock_names
})


# 只儲存股票代號欄位成 txt，每行一個代號，不加索引與標題
clean_df['股票代號'].to_csv('stock_ids.csv', index=False, header=False)

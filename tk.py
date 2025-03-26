import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pandas as pd
import os
import threading
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
# read files from excel
def process_file(filepath, year, month):
    try:
        df = pd.read_excel(filepath, dtype=str)
        name_col = next((col for col in df.columns if "姓" in col), None)
        role_col = next((col for col in df.columns if "身份別" in col), None)
        prev_col = next((col for col in df.columns if "上月實際持有股數" in col), None)
        curr_col = next((col for col in df.columns if "本月實際自有持有股數" in col), None)
        inc_col = "本月增加_自有股數(集中)" if "本月增加_自有股數(集中)" in df.columns else None
        dec_col = "本月減少_自有股數(集中)" if "本月減少_自有股數(集中)" in df.columns else None
        if not all([name_col, role_col, prev_col, curr_col]) or (not inc_col and not dec_col):
            return []
        df[inc_col] = pd.to_numeric(df.get(inc_col, 0), errors="coerce").fillna(0)
        df[dec_col] = pd.to_numeric(df.get(dec_col, 0), errors="coerce").fillna(0)
        changed_df = df[(df[inc_col] != 0) | (df[dec_col] != 0)]
        result = []
        for _, row in changed_df.iterrows():
            company_id = os.path.basename(filepath).split("_")[1]
            result.append({
                "公司檔名": company_id,
                "姓名": row[name_col],
                "身份別": row[role_col],
                "上月持股": str(row.get(prev_col, "")).split()[1],
                "增加持股": int(row.get(inc_col, 0)),
                "減少持股": int(row.get(dec_col, 0)),
                "本月持股": str(row.get(curr_col, "")).split()[0]
            })
        return result
    except Exception as e:
        print(f"⚠️ 錯誤讀取 {filepath}: {e}")
        return []

def to_float_safe(val):
    try:
        return float(str(val).replace(",", "").split()[0])
    except:
        return 0.0

def fetch_insider_holdings(co_id: str, year: str, month: str, output_dir: str = "./output"):
    """爬取內部人持股異動申報表，若「自有股數(集中)」有異動則輸出 Excel。"""

    os.makedirs(output_dir, exist_ok=True)
    driver = webdriver.Chrome()
    filename = f"{year}年{month}月_{co_id}_內部人持股.xlsx"
    output_path = os.path.join(output_dir, filename)
    if os.path.exists(output_path):
        print(f"🟡 已存在：{output_path}，跳過爬取")
        return  # 直接跳過這次爬蟲
    driver.get("https://mopsov.twse.com.tw/mops/web/query6_1")

    try:
        wait = WebDriverWait(driver, 10)
        Select(wait.until(EC.presence_of_element_located((By.ID, "isnew")))).select_by_value("false")

        driver.find_element(By.ID, "co_id").send_keys(co_id)
        driver.find_element(By.ID, "year").send_keys(year)
        Select(driver.find_element(By.ID, "month")).select_by_value(month)
        driver.find_element(By.XPATH, "//input[@type='button' and @value=' 查詢 ']").click()

        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#table01 .hasBorder")))
        time.sleep(1)

        table_html = driver.find_element(By.CSS_SELECTOR, "#table01 .hasBorder").get_attribute("outerHTML")
        df_list = pd.read_html(table_html)
        if not df_list:
            print("⚠️ 沒有找到表格")
            return

        df = df_list[0]
        if df.empty:
            print("⚠️ 表格為空")
            return

        # 扁平化欄位
        df.columns = [' '.join(col).strip() if isinstance(col, tuple) else col for col in df.columns]

        # 找出本月增減欄位
        increase_col = next((col for col in df.columns if "本月增加" in col and "自有股數(集中)" in col), None)
        decrease_col = next((col for col in df.columns if "本月減少" in col and "自有股數(集中)" in col), None)

        if not increase_col and not decrease_col:
            print("⚠️ 找不到『自有股數(集中)』欄位")
            return

        # 分割欄位為數字欄位
        if increase_col:
            df["本月增加_自有股數(集中)"] = df[increase_col].apply(to_float_safe)

        if decrease_col:
            df["本月減少_自有股數(集中)"] = df[decrease_col].apply(to_float_safe)
        check_cols = []
        if "本月增加_自有股數(集中)" in df.columns:
            check_cols.append("本月增加_自有股數(集中)")
        if "本月減少_自有股數(集中)" in df.columns:
            check_cols.append("本月減少_自有股數(集中)")

        # 判斷有變動才輸出
        if check_cols and (df[check_cols] != 0).any(axis=None):
            output_path = os.path.join(output_dir, f"{year}年{month}月_{co_id}_內部人持股.xlsx")
            df.to_excel(output_path, index=False)
            print(f"✅ 有『自有股數(集中)』異動 → 已輸出：{output_path}")
            #print(df[["身份別", "姓 名"] + check_cols][(df[check_cols] != 0).any(axis=1)])
        else:
            print(f"🟡 {year}年{month}月：『自有股數(集中)』無異動")

    except Exception as e:
        print(f"❌ 發生錯誤：{e}")
    finally:
        driver.quit()

def load_filtered_data(year, month):
    folder = "./output"
    files = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.endswith(".xlsx") and f"{year}年{month}月" in f
    ]

    result = []
    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = [executor.submit(process_file, file, year, month) for file in files]
        for future in as_completed(futures):
            result.extend(future.result())

    return result

def show_data():
    year = year_entry.get()
    month = month_entry.get()
    if not year or not month:
        messagebox.showwarning("輸入錯誤", "請輸入年份與月份！")
        return

    loading = tk.Toplevel(root)
    loading.title("載入中")
    loading.geometry("200x80")
    loading_label = tk.Label(loading, text="查詢中...\n請稍候", font=("Arial", 12))
    loading_label.pack(expand=True)
    loading.update_idletasks()
    x = (loading.winfo_screenwidth() - loading.winfo_reqwidth()) // 2
    y = (loading.winfo_screenheight() - loading.winfo_reqheight()) // 2
    loading.geometry(f"+{x}+{y}")
    loading.protocol("WM_DELETE_WINDOW", lambda: None)

    def query():
        stock_df = pd.read_csv("stock_ids.csv", dtype=str)
        stock_list = stock_df.iloc[:, 0].dropna().tolist()
        max_threads = min(4, len(stock_list))  # 建議最多同時 5 執行緒
        with ThreadPoolExecutor(max_workers=max_threads) as executor:
            futures = {
                executor.submit(fetch_insider_holdings, stock_id, year, f"{int(month):02d}"): stock_id
                for stock_id in stock_list
            }
            for future in as_completed(futures):
                stock_id = futures[future]
                try:
                    future.result()
                except Exception as e:
                    print(f"⚠️ {stock_id} 查詢失敗：{e}")
        
        # 匯總顯示
        records = load_filtered_data(year, f"{int(month):02d}")
        for row in tree.get_children():
            tree.delete(row)
        for r in records:
            tree.insert("", "end", values=(r["公司檔名"], r["身份別"], r["姓名"], r["上月持股"], r["增加持股"], r["減少持股"], r["本月持股"]))
        if records:
            df_output = pd.DataFrame(records)
            output_file = os.path.join("./output", f"{year}年{month.zfill(2)}月_變化.xlsx")
            df_output.to_excel(output_file,index=False)
            print(f"已匯出資料到:{output_file}")
        loading.destroy()
    threading.Thread(target=query).start()

# main
root = tk.Tk()
root.title("內部人持股異動查詢 GUI")
root.geometry("900x600")
root.minsize(800, 400)
frame = tk.Frame(root)
frame.pack(pady=10)
tk.Label(frame, text="年份").grid(row=0, column=0)
year_entry = tk.Entry(frame, width=5)
year_entry.grid(row=0, column=1)
tk.Label(frame, text="月份").grid(row=0, column=2)
month_entry = tk.Entry(frame, width=5)
month_entry.grid(row=0, column=3)
tk.Button(frame, text="查詢", command=show_data).grid(row=0, column=4, padx=10)
columns = ("公司", "身份別", "姓名", "上月持股", "增加持股", "減少持股", "本月持股")
tree = ttk.Treeview(root, columns=columns, show="headings", height=15)
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor="center")
scroll_x = ttk.Scrollbar(root, orient="horizontal", command=tree.xview)
scroll_y = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
tree.configure(xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)
scroll_y.pack(side="right", fill="y")
scroll_x.pack(side="bottom", fill="x")
tree.pack(expand=True, fill="both")
root.mainloop()

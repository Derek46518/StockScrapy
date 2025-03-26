from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import os

def to_float_safe(val):
    try:
        return float(str(val).replace(",", "").split()[0])
    except:
        return 0.0


def fetch_insider_holdings(co_id: str, year: str, month: str, output_dir: str = "./output"):
    """爬取內部人持股異動申報表，若「自有股數(集中)」有異動則輸出 Excel。"""

    os.makedirs(output_dir, exist_ok=True)
    driver = webdriver.Chrome()
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
fetch_insider_holdings("2330", "113", "12")
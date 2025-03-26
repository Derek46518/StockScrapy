from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import time
import os

def fetch_insider_holdings(co_id: str, year: str, month: str, output_dir: str = "./output"):
    """根據公司代號、年度與月份，爬取公開資訊觀測站的內部人持股異動申報資料，並儲存為 Excel 檔案。"""

    # 建立輸出資料夾（若不存在）
    os.makedirs(output_dir, exist_ok=True)

    # 啟動 Selenium
    driver = webdriver.Chrome()
    driver.get("https://mopsov.twse.com.tw/mops/web/query6_1")

    try:
        wait = WebDriverWait(driver, 10)

        # 選擇「歷史資料」
        Select(wait.until(EC.presence_of_element_located((By.ID, "isnew")))).select_by_value("false")

        # 輸入參數
        driver.find_element(By.ID, "co_id").send_keys(co_id)
        driver.find_element(By.ID, "year").send_keys(year)
        Select(driver.find_element(By.ID, "month")).select_by_value(month)

        # 點擊查詢
        driver.find_element(By.XPATH, "//input[@type='button' and @value=' 查詢 ']").click()

        # 等待表格出現
        wait.until(EC.presence_of_element_located((By.ID, "table01")))
        time.sleep(1)  # 多等一下以確保 DOM 完全載入

        # 擷取表格 HTML 並轉換為 DataFrame
        table_html = driver.find_element(By.CSS_SELECTOR, "#table01 .hasBorder").get_attribute("outerHTML")
        df_list = pd.read_html(table_html)

        if df_list:
            
            df = df_list[0]
            
            # 展平 MultiIndex 欄位
            df.columns = [' '.join(map(str, col)).strip() for col in df.columns.values]
            output_path = os.path.join(output_dir, f"{co_id}_內部人持股_{year}年{month}月.xlsx")
            df.to_excel(output_path, index=False)
            print(f"✅ 資料已儲存至：{output_path}")
        else:
            print("⚠️ 沒有找到有效的表格")
    except Exception as e:
        print(f"❌ 發生錯誤：{e}")
    finally:
        driver.quit()
fetch_insider_holdings("2330", "113", "12")
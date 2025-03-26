from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import time

# 啟動 Selenium driver
driver = webdriver.Chrome()
driver.get("https://mopsov.twse.com.tw/mops/web/query6_1")

# 等待 select[id='isnew'] 出現（最多 10 秒）
wait = WebDriverWait(driver, 10)
select_element = wait.until(EC.presence_of_element_located((By.ID, "isnew")))

# 選擇歷史資料
Select(select_element).select_by_value("false")

# 填入公司代號、年度、月份
driver.find_element(By.ID, "co_id").send_keys("2330")
driver.find_element(By.ID, "year").send_keys("110")
Select(driver.find_element(By.ID, "month")).select_by_value("12")

# 點擊查詢按鈕
query_btn = driver.find_element(By.XPATH, "//input[@type='button' and @value=' 查詢 ']")
query_btn.click()

# 假設你已經打開 driver，並已經觸發查詢了
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "table01"))
)

time.sleep(1)  # 保險起見，再等一下 DOM 完全載入

# 抓取 table01 裡的 HTML
table_html = driver.find_element(By.CSS_SELECTOR, "#table01 .hasBorder").get_attribute("outerHTML")
# 解析成 pandas DataFrame
df_list = pd.read_html(table_html)
if df_list:
    df = df_list[0]
    print(df)
    output_path = "台積電_內部人持股_113年12月.xlsx"
    df.to_excel(output_path, index=True)
else:
    print("⚠️ 沒有找到有效的表格")

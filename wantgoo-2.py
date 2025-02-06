import sys

sys.path.append(r"C:\Users\user\Desktop\autoUpdateETF-master\venv\Lib\site-packages")
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


# 讀取 Google Sheets 中的 ETF 代碼
def get_etf_codes(sheet_id, range_name):
    credentials = Credentials.from_service_account_file(
        "spry-cat-442903-a3-91abd63953b6.json"
    )
    service = build("sheets", "v4", credentials=credentials)
    sheet = service.spreadsheets()

    result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
    values = result.get("values", [])

    return [row[0] for row in values]  # ETF 代碼在每行的第一列


# **使用 Selenium 爬取 近四季殖利率 & 年報酬率**
def fetch_all_etf_data(etf_codes):
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    dividend_yields = []
    annual_returns = []

    try:
        driver.get("http://127.0.0.1:8000/page.html")  # 你的 HTML 頁面
        for etf_code in etf_codes:
            try:
                # **找到 ETF 代碼**
                xpath_code = f"//a[normalize-space(text())='{etf_code}']"
                WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.XPATH, xpath_code))
                )

                # **抓取 近四季殖利率 (td[9]) & 年報酬率 (td[10])**
                xpath_dividend = f"{xpath_code}/../../td[9]"
                xpath_return = f"{xpath_code}/../../td[10]"

                dividend_element = driver.find_element(By.XPATH, xpath_dividend)
                annual_return_element = driver.find_element(By.XPATH, xpath_return)

                # **添加 "%" 符號**
                dividend_yield = dividend_element.text.strip()
                dividend_yield = (
                    f"{dividend_yield}%"
                    if dividend_yield and not dividend_yield.endswith("%")
                    else dividend_yield
                )

                dividend_yields.append(dividend_yield)
                annual_returns.append(annual_return_element.text.strip())

                print(
                    f"ETF: {etf_code}, 殖利率: {dividend_yield}, 年報酬率: {annual_return_element.text.strip()}"
                )
            except Exception as e:
                print(f"Error fetching data for {etf_code}: {e}")
                dividend_yields.append("N/A")
                annual_returns.append("N/A")
    finally:
        driver.quit()

    return dividend_yields, annual_returns


# **使用 BeautifulSoup 解析本地 HTML**
def fetch_all_etf_data_bs(local_html_path, etf_codes):
    with open(local_html_path, "r", encoding="utf-8") as file:
        soup = BeautifulSoup(file, "html.parser")

    dividend_yields = []
    annual_returns = []

    for etf_code in etf_codes:
        try:
            etf_element = soup.find("a", text=etf_code)
            if etf_element:
                row = etf_element.find_parent("tr")
                dividend_element = row.find_all("td")[8]  # **近四季殖利率**
                annual_return_element = row.find_all("td")[9]  # **年報酬率**

                # **添加 "%" 符號**
                dividend_yield = dividend_element.text.strip()
                dividend_yield = (
                    f"{dividend_yield}%"
                    if dividend_yield and not dividend_yield.endswith("%")
                    else dividend_yield
                )

                dividend_yields.append(dividend_yield)
                annual_returns.append(annual_return_element.text.strip())

                print(
                    f"ETF: {etf_code}, 殖利率: {dividend_yield}, 年報酬率: {annual_return_element.text.strip()}"
                )
            else:
                dividend_yields.append("N/A")
                annual_returns.append("N/A")
        except Exception as e:
            print(f"Error fetching data for {etf_code}: {e}")
            dividend_yields.append("N/A")
            annual_returns.append("N/A")

    return dividend_yields, annual_returns


# **更新 Google Sheets**
def update_google_sheet(
    sheet_id, range_dividend, range_return, dividend_values, return_values
):
    credentials = Credentials.from_service_account_file(
        "spry-cat-442903-a3-91abd63953b6.json"
    )
    service = build("sheets", "v4", credentials=credentials)
    sheet = service.spreadsheets()

    # **更新 殖利率**
    body_dividend = {"values": [[value] for value in dividend_values]}
    result_dividend = (
        sheet.values()
        .update(
            spreadsheetId=sheet_id,
            range=range_dividend,
            body=body_dividend,
            valueInputOption="RAW",
        )
        .execute()
    )
    print(f"Updated {result_dividend.get('updatedCells')} cells for Dividend Yield.")

    # **更新 年報酬率**
    body_return = {"values": [[value] for value in return_values]}
    result_return = (
        sheet.values()
        .update(
            spreadsheetId=sheet_id,
            range=range_return,
            body=body_return,
            valueInputOption="RAW",
        )
        .execute()
    )
    print(f"Updated {result_return.get('updatedCells')} cells for Annual Return.")


# **主函數**
def process_etf_data():
    sheet_id = "1ouX_BHS9g3HQBgyo73T2uc37SdAZ0diR1LgHn6YAWwM"
    read_range = "ETF分析!C5:C34"  # 讀取 ETF 代碼
    write_dividend_range = "ETF分析!I5:I34"  # **存放 近四季殖利率**
    write_return_range = "ETF分析!G5:G34"  # **存放 年報酬率**

    # **讀取 Google Sheets ETF 代碼**
    etf_codes = get_etf_codes(sheet_id, read_range)
    print(f"Found ETF codes: {etf_codes}")

    # **使用 Selenium 爬取數據**
    dividend_yields, annual_returns = fetch_all_etf_data(etf_codes)

    # 或 **使用 BeautifulSoup**
    # dividend_yields, annual_returns = fetch_all_etf_data_bs("/mnt/data/page.html", etf_codes)

    # **更新 Google Sheets**
    update_google_sheet(
        sheet_id,
        write_dividend_range,
        write_return_range,
        dividend_yields,
        annual_returns,
    )
    print("Google Sheets updated successfully!")


# **執行主程式**
if __name__ == "__main__":
    process_etf_data()

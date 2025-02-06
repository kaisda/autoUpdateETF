import os
import sys

# **確保 Python 讀取虛擬環境 (venv)**
sys.path.append(r"C:\Users\user\Desktop\autoUpdateETF-master\venv\Lib\site-packages")

import gspread
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials


# **Google Sheets 認證**
def get_google_sheet(sheet_id, range_code):
    creds = Credentials.from_service_account_file(
        "spry-cat-442903-a3-91abd63953b6.json",
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id).worksheet("ETF分析")

    # **讀取 ETF 代碼**
    etf_codes = sheet.get(range_code)
    etf_codes = [row[0].strip() for row in etf_codes if row]
    return sheet, etf_codes


# **解析本地 HTML**
def parse_html_for_etf_yield(html_file, etf_codes):
    with open(html_file, "r", encoding="utf-8") as file:
        soup = BeautifulSoup(file, "html.parser")

    dividend_yields = []

    for etf_code in etf_codes:
        try:
            etf_element = next(
                (a for a in soup.find_all("a") if a.text.strip() == etf_code), None
            )
            if etf_element:
                row = etf_element.find_parent("tr")

                # **抓取該行的所有 <td>**
                cells = row.find_all("td")
                cell_texts = [td.text.strip() for td in cells]
                print(f"ETF: {etf_code} -> Table Data: {cell_texts}")

                # **殖利率在 td[9]**
                if len(cells) > 9:
                    yield_value = cells[9].text.strip()
                else:
                    yield_value = "N/A"

                # **確保數據帶有 "%"**
                if (
                    yield_value
                    and yield_value != "N/A"
                    and not yield_value.endswith("%")
                ):
                    yield_value += "%"

                dividend_yields.append(yield_value)
                print(f"ETF: {etf_code}, 殖利率: {yield_value}")
            else:
                print(f"ETF {etf_code} not found in HTML.")
                dividend_yields.append("N/A")
        except Exception as e:
            print(f"Error fetching yield for {etf_code}: {e}")
            dividend_yields.append("N/A")

    return dividend_yields


# **更新 Google Sheets**
def update_google_sheet(sheet, range_dividend, dividend_values):
    sheet.update(range_dividend, [[value] for value in dividend_values])
    print("Google Sheets updated successfully!")


# **主函數**
def main():
    sheet_id = "1ouX_BHS9g3HQBgyo73T2uc37SdAZ0diR1LgHn6YAWwM"
    read_range = "C5:C34"  # **讀取 ETF 代碼**
    write_range = "H5:H34"  # **存放殖利率**

    # **使用相對路徑讀取 HTML 檔案**
    script_dir = os.path.dirname(
        os.path.abspath(__file__)
    )  # 取得當前 Python 檔案所在目錄
    html_file = os.path.join(script_dir, "moneydj.html")  # 假設 HTML 在相同目錄

    sheet, etf_codes = get_google_sheet(sheet_id, read_range)
    print(f"Found ETF codes: {etf_codes}")

    # **解析 HTML，獲取殖利率**
    dividend_yields = parse_html_for_etf_yield(html_file, etf_codes)

    # **更新 Google Sheets**
    update_google_sheet(sheet, write_range, dividend_yields)


# **執行主程式**
if __name__ == "__main__":
    main()

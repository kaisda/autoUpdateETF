import os
import sys

# **確保 Python 讀取虛擬環境**
sys.path.append(r"C:\Users\user\Desktop\autoUpdateETF-master\venv\Lib\site-packages")

import gspread
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials


# **Google Sheets 認證**
def get_google_sheet(sheet_id):
    creds = Credentials.from_service_account_file(
        "spry-cat-442903-a3-91abd63953b6.json",
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id).worksheet("ETF分析")

    # **自動偵測 C 欄位範圍**
    col_C = sheet.col_values(3)  # 取得整列 C
    etf_codes = col_C[4:]  # 從 C5 開始

    last_row = len(col_C)  # 取得最後一行
    read_range = f"C5:C{last_row}"
    write_dividend_range = f"I5:I{last_row}"  # I 欄寫入殖利率
    write_return_range = f"G5:G{last_row}"  # G 欄寫入年報酬率

    return sheet, etf_codes, read_range, write_dividend_range, write_return_range


# **解析本地 HTML**
def parse_html_for_etf_data(html_file, etf_codes):
    with open(html_file, "r", encoding="utf-8") as file:
        soup = BeautifulSoup(file, "html.parser")

    dividend_yields = []
    annual_returns = []

    for etf_code in etf_codes:
        try:
            etf_element = next(
                (a for a in soup.find_all("a") if a.text.strip() == etf_code), None
            )
            if etf_element:
                row = etf_element.find_parent("tr")
                cells = row.find_all("td")

                dividend_yield = cells[8].text.strip() if len(cells) > 8 else "N/A"
                annual_return = cells[9].text.strip() if len(cells) > 9 else "N/A"

                if (
                    dividend_yield
                    and dividend_yield != "N/A"
                    and not dividend_yield.endswith("%")
                ):
                    dividend_yield += "%"

                dividend_yields.append(dividend_yield)
                annual_returns.append(annual_return)

                print(
                    f"ETF: {etf_code}, 近四季殖利率: {dividend_yield}, 年報酬率: {annual_return}"
                )
            else:
                dividend_yields.append("N/A")
                annual_returns.append("N/A")
        except Exception as e:
            dividend_yields.append("N/A")
            annual_returns.append("N/A")

    return dividend_yields, annual_returns


# **更新 Google Sheets**
def update_google_sheet(
    sheet, write_dividend_range, write_return_range, dividend_values, return_values
):
    sheet.update(write_dividend_range, [[value] for value in dividend_values])
    sheet.update(write_return_range, [[value] for value in return_values])
    print("Google Sheets updated successfully!")


# **主函數**
def main():
    sheet_id = "1ouX_BHS9g3HQBgyo73T2uc37SdAZ0diR1LgHn6YAWwM"

    # **取得自動範圍**
    sheet, etf_codes, read_range, write_dividend_range, write_return_range = (
        get_google_sheet(sheet_id)
    )
    print(f"自動偵測範圍: {read_range} -> {write_dividend_range}, {write_return_range}")

    # **解析 HTML，獲取殖利率 & 年報酬率**
    script_dir = os.path.dirname(os.path.abspath(__file__))
    html_file = os.path.join(script_dir, "wantgoo.html")
    dividend_yields, annual_returns = parse_html_for_etf_data(html_file, etf_codes)

    # **更新 Google Sheets**
    update_google_sheet(
        sheet, write_dividend_range, write_return_range, dividend_yields, annual_returns
    )


# **執行主程式**
if __name__ == "__main__":
    main()

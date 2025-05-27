import os
import tempfile
import pandas as pd
from flask import Flask, request, render_template_string, send_file
from fpdf import FPDF
from FinMind.data import DataLoader
from datetime import datetime, timedelta

app = Flask(__name__)
FONT_PATH = "fonts/NotoSansTC.ttf"
TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJkYXRlIjoiMjAyNS0wNS0yOCAwMToyODoxMyIsInVzZXJfaWQiOiJqYW1lczkwMTAxNiIsImlwIjoiMTE4LjE1MC42My45OSJ9.QWxBrJYWM_GNDpTyvAyR2frCPwB4e7HP_Kj_KEX2tVs"

@app.route('/')
def index():
    with open("templates/index.html", encoding="utf-8") as f:
        return render_template_string(f.read())

@app.route('/analyze', methods=['POST'])
def analyze():
    # 自訂參數，若未填則使用預設值
    try:
        pe_ratio = float(request.form.get("pe_ratio", 15))
        pb_ratio = float(request.form.get("pb_ratio", 1.5))
        eps_growth = float(request.form.get("eps_growth", 10))
        roe = float(request.form.get("roe", 15))
        debt_ratio = float(request.form.get("debt_ratio", 50))
        price_3m = float(request.form.get("price_3m", 0.1))
        std_1y = float(request.form.get("std_1y", 0.2))
    except:
        return "❌ 請輸入正確的數字格式。"

    # 抓取資料（上市股票）
    api = DataLoader()
    api.login_by_token(api_token=TOKEN)

    stock_list = api.taiwan_stock_info()
    stock_list = stock_list[stock_list["type"] == "股票"]  # 避免 ETF

    result_list = []
    today = datetime.today()
    start = (today - timedelta(days=365)).strftime("%Y-%m-%d")
    end = today.strftime("%Y-%m-%d")

    for _, row in stock_list.iterrows():
        stock_id = row["stock_id"]
        stock_name = row["stock_name"]

        try:
            price_df = api.taiwan_stock_daily(stock_id=stock_id, start_date=start, end_date=end)
            if price_df.empty or len(price_df) < 60:
                continue

            price_df["return"] = price_df["close"].pct_change()
            std = price_df["return"].std() * (252 ** 0.5)
            pct_3m = (price_df.iloc[-1]["close"] / price_df.iloc[-63]["close"]) - 1

            fin_df = api.taiwan_stock_financial_statement(
                stock_id=stock_id,
                start_date="2023-01-01",
                end_date="2024-12-31"
            )
            if fin_df.empty:
                continue

            # 最近期數據
            latest = fin_df.iloc[-1]

            # 條件篩選
            if (
                latest["PBR"] < pb_ratio and
                latest["PER"] < pe_ratio and
                latest["EPS"] > 0 and
                latest["EPS"] > eps_growth / 100 * latest["EPS"] and
                latest["ROE"] > roe and
                latest["DebtRatio"] < debt_ratio and
                pct_3m > price_3m and
                std < std_1y
            ):
                result_list.append({
                    "symbol": stock_id,
                    "name": stock_name,
                    "pe_ratio": latest["PER"],
                    "pb_ratio": latest["PBR"],
                    "eps": latest["EPS"],
                    "roe": latest["ROE"],
                    "debt_ratio": latest["DebtRatio"],
                    "price_3m": round(pct_3m, 2),
                    "std_1y": round(std, 4)
                })

        except Exception as e:
            continue

    df = pd.DataFrame(result_list)

    # 產出 PDF 報告
    pdf = FPDF()
    pdf.add_page()
    try:
        pdf.add_font('NotoSansTC', '', FONT_PATH, uni=True)
        pdf.set_font('NotoSansTC', '', 14)
    except RuntimeError:
        print("⚠️ 找不到中文字型，改用內建字型（PDF 僅英文顯示）")
        pdf.set_font('Arial', '', 14)


    pdf.cell(0, 10, "多因子選股分析結果", ln=True)

    if df.empty:
        pdf.cell(0, 10, "沒有符合條件的個股", ln=True)
    else:
        for i, row in df.iterrows():
            line = f"{row['symbol']} - {row['name']} | PE: {row['pe_ratio']} PB: {row['pb_ratio']} EPS: {row['eps']} ROE: {row['roe']} 負債比: {row['debt_ratio']} 3月漲: {row['price_3m']} 波動: {row['std_1y']}"
            pdf.multi_cell(0, 10, line)

    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(temp.name)

    return send_file(temp.name, as_attachment=True, download_name="選股分析報告.pdf")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)


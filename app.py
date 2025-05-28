import os
import tempfile
import pandas as pd
from flask import Flask, request, render_template_string, send_file
from datetime import datetime, timedelta
from FinMind.data import DataLoader
from docx import Document

app = Flask(__name__)
TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJkYXRlIjoiMjAyNS0wNS0yOCAwMToyODoxMyIsInVzZXJfaWQiOiJqYW1lczkwMTAxNiIsImlwIjoiMTE4LjE1MC42My45OSJ9.QWxBrJYWM_GNDpTyvAyR2frCPwB4e7HP_Kj_KEX2tVs"

@app.route("/")
def index():
    with open("templates/index.html", encoding="utf-8") as f:
        return render_template_string(f.read())

@app.route("/analyze", methods=["POST"])
def analyze():
    # 使用者參數，預設值
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

    # FinMind 抓資料
    api = DataLoader()
    api.login_by_token(api_token=TOKEN)
    stock_list = api.taiwan_stock_info()
    stock_list = stock_list[stock_list["type"] == "股票"]  # 排除 ETF

    result = []
    today = datetime.today()
    start = (today - timedelta(days=365)).strftime("%Y-%m-%d")
    end = today.strftime("%Y-%m-%d")

    for _, row in stock_list.iterrows():
        stock_id = row["stock_id"]
        stock_name = row["stock_name"]
        try:
            # 價格資料
            price_df = api.taiwan_stock_daily(stock_id=stock_id, start_date=start, end_date=end)
            if price_df.empty or len(price_df) < 65:
                continue

            price_df["return"] = price_df["close"].pct_change()
            std = price_df["return"].std() * (252 ** 0.5)
            pct_3m = (price_df.iloc[-1]["close"] / price_df.iloc[-63]["close"]) - 1

            # 財報資料
            fin_df = api.taiwan_stock_financial_statement(
                stock_id=stock_id,
                start_date="2023-01-01",
                end_date="2024-12-31"
            )
            if fin_df.empty:
                continue

            latest = fin_df.iloc[-1]
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
                result.append({
                    "symbol": stock_id,
                    "name": stock_name,
                    "pe_ratio": latest["PER"],
                    "pb_ratio": latest["PBR"],
                    "eps": latest["EPS"],
                    "roe": latest["ROE"],
                    "debt_ratio": latest["DebtRatio"],
                    "price_3m": round(pct_3m, 3),
                    "std_1y": round(std, 4)
                })
        except:
            continue

    df = pd.DataFrame(result)

    # Word 報告
    doc = Document()
    doc.add_heading("📊 多因子選股分析報告", 0)
    if df.empty:
        doc.add_paragraph("⚠️ 沒有符合條件的個股")
    else:
        for _, r in df.iterrows():
            doc.add_paragraph(
                f"{r['symbol']} - {r['name']}｜PE: {r['pe_ratio']:.1f} PB: {r['pb_ratio']:.1f} EPS: {r['eps']:.2f} ROE: {r['roe']}% 負債比: {r['debt_ratio']}% 三月漲幅: {r['price_3m']*100:.1f}% 波動: {r['std_1y']:.2f}"
            )

    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp.name)
    return send_file(temp.name, as_attachment=True, download_name="選股分析報告.docx")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)


from flask import Flask, render_template, request, send_file
import pandas as pd
from FinMind.data import DataLoader
from fpdf import FPDF
import tempfile
import os

app = Flask(__name__)
token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJkYXRlIjoiMjAyNS0wNS0yOCAwMToyODoxMyIsInVzZXJfaWQiOiJqYW1lczkwMTAxNiIsImlwIjoiMTE4LjE1MC42My45OSJ9.QWxBrJYWM_GNDpTyvAyR2frCPwB4e7HP_Kj_KEX2tVs"

def fetch_stock_data():
    dl = DataLoader()
    dl.login_by_token(api_token=token)
    info = dl.taiwan_stock_info()
    info = info[(info['type'].isin(['sii', 'otc'])) & (~info['industry_category'].str.contains("ETF|ETN|受益|指數", na=False))]
    return info[['stock_id', 'stock_name']].drop_duplicates()

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/analyze", methods=["POST"])
def analyze():
    # 取得前端參數
    pe_ratio = float(request.form.get("pe_ratio", 15))
    pb_ratio = float(request.form.get("pb_ratio", 1.5))
    eps_growth = float(request.form.get("eps_growth", 10))
    roe = float(request.form.get("roe", 15))
    debt_ratio = float(request.form.get("debt_ratio", 50))
    price_3m = float(request.form.get("price_3m", 0.1))
    std_1y = float(request.form.get("std_1y", 0.2))

    stock_list = fetch_stock_data()
    selected = []

    for _, row in stock_list.iterrows():
        stock_id = row["stock_id"]
        try:
            df = DataLoader().taiwan_stock_financial_statement(stock_id=stock_id, start_date="2023-01-01", end_date="2023-12-31")
            if df.empty: continue
            df = df.sort_values("date").drop_duplicates("type", keep="last")

            stock_pe = df["PE"].astype(float).values[-1]
            stock_pb = df["PB"].astype(float).values[-1]
            stock_eps = df["EPS"].astype(float).pct_change().values[-1] * 100
            stock_roe = df["ROE"].astype(float).values[-1]
            stock_debt = df["負債比率"].astype(float).values[-1]

            if stock_pe < pe_ratio and stock_pb < pb_ratio and stock_eps > eps_growth and stock_roe > roe and stock_debt < debt_ratio:
                selected.append([stock_id, row["stock_name"], stock_pe, stock_pb, stock_eps, stock_roe, stock_debt])
        except:
            continue

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="多因子選股結果", ln=True, align="C")
    pdf.ln(10)

    for row in selected:
        pdf.cell(0, 10, txt=f"{row[0]} {row[1]} | PE: {row[2]:.2f} PB: {row[3]:.2f} EPS成長: {row[4]:.2f}% ROE: {row[5]:.2f}% 負債比: {row[6]:.2f}%", ln=True)

    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(temp.name)

    return send_file(temp.name, as_attachment=True, download_name="selected_stocks.pdf")

if __name__ == "__main__":
    app.run(debug=True)

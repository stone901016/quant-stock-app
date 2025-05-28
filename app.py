import os
import tempfile
import pandas as pd
from flask import Flask, request, render_template_string, send_file
from docx import Document
from datetime import datetime

app = Flask(__name__)

@app.route('/')
def index():
    with open("templates/index.html", encoding="utf-8") as f:
        return render_template_string(f.read())

@app.route('/analyze', methods=['POST'])
def analyze():
    # 讀取使用者上傳的 CSV
    if 'file' not in request.files:
        return "❌ 沒有檔案"
    file = request.files['file']
    try:
        df = pd.read_csv(file)
    except Exception as e:
        return f"❌ CSV 檔案錯誤：{e}"

    # 取得自訂參數（若沒輸入則用預設值）
    try:
        pe_ratio = float(request.form.get("pe_ratio", 15))
        pb_ratio = float(request.form.get("pb_ratio", 1.5))
        eps_growth = float(request.form.get("eps_growth", 10))
        roe = float(request.form.get("roe", 15))
        debt_ratio = float(request.form.get("debt_ratio", 50))
        price_3m = float(request.form.get("price_3m", 0.1))
        std_1y = float(request.form.get("std_1y", 0.2))
    except ValueError:
        return "❌ 請輸入有效的參數數字"

    # 欄位檢查
    required_cols = ["symbol", "name", "pe_ratio", "pb_ratio", "eps", "roe", "debt_ratio", "price_3m", "std_1y"]
    for col in required_cols:
        if col not in df.columns:
            return f"❌ 缺少欄位：{col}"

    # 條件篩選
    filtered = df[
        (df["pe_ratio"] < pe_ratio) &
        (df["pb_ratio"] < pb_ratio) &
        (df["eps"] > 0) &
        (df["eps"] > eps_growth / 100 * df["eps"]) &
        (df["roe"] > roe) &
        (df["debt_ratio"] < debt_ratio) &
        (df["price_3m"] > price_3m) &
        (df["std_1y"] < std_1y)
    ]

    # 產出 Word 檔
    doc = Document()
    doc.add_heading("📊 多因子選股分析結果", level=1)
    doc.add_paragraph(f"產出時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    if filtered.empty:
        doc.add_paragraph("沒有符合條件的個股。")
    else:
        for _, row in filtered.iterrows():
            doc.add_paragraph(
                f"{row['symbol']} - {row['name']}\n"
                f"PE: {row['pe_ratio']}  PB: {row['pb_ratio']}  EPS: {row['eps']}\n"
                f"ROE: {row['roe']}  負債比: {row['debt_ratio']}  三月漲幅: {row['price_3m']}  波動: {row['std_1y']}"
            )

    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp.name)
    return send_file(temp.name, as_attachment=True, download_name="選股分析報告.docx")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)

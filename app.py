from flask import Flask, render_template, request
from openpyxl import load_workbook
from datetime import datetime

app = Flask(__name__)

EXCEL_FILE = "diem_thi_namdinh.xlsx"
COLS = {
    "van": 2,   # Cột B
    "toan": 3,  # Cột C
    "anh": 4,   # Cột D
    "tong": 5   # Cột E
}
SBD_COL = 1    # Cột A

@app.route("/", methods=["GET", "POST"])
def index():
    ranks = {}
    scores = {}
    sbd = ""

    if request.method == "POST":
        sbd = request.form["sbd"].strip()
        print(f"[{datetime.now()}] Tra cứu SBD: {sbd}")
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active

        for key, col in COLS.items():
            all_scores = []
            for row in ws.iter_rows(min_row=2):
                sbd_cell = str(row[SBD_COL - 1].value).strip()
                raw_score = row[col - 1].value
                try:
                    score = float(str(raw_score).replace(",", "."))
                    all_scores.append((sbd_cell, score))
                except:
                    continue
            all_scores.sort(key=lambda x: x[1], reverse=True)
            for idx, (curr_sbd, score) in enumerate(all_scores, start=1):
                if curr_sbd == sbd:
                    scores[key] = score
                    ranks[key] = idx
                    break

    return render_template("index.html", ranks=ranks, scores=scores, sbd=sbd)

if __name__ == "__main__":
    app.run(debug=True)
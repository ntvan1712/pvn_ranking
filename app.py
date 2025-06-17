from flask import Flask, render_template, request
from openpyxl import load_workbook
from datetime import datetime

app = Flask(__name__)

EXCEL_FILE = "diem_thi_namdinh.xlsx"
TOTAL_SCORE_COL = 5  # Cột E
SBD_COL = 1  # Cột A

@app.route("/", methods=["GET", "POST"])
def index():
    rank = None
    your_score = None
    total_scores = []

    if request.method == "POST":
        sbd = request.form["sbd"].strip()
        print(f"[{datetime.now()}] Tra cứu SBD: {sbd}")
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active

        for row in ws.iter_rows(min_row=2):
            sbd_cell = str(row[SBD_COL - 1].value).strip()
            score_raw = row[TOTAL_SCORE_COL - 1].value

            try:
                score = float(str(score_raw).replace(",", "."))
                total_scores.append((sbd_cell, score))
            except:
                continue

        # Sắp xếp theo điểm giảm dần
        total_scores.sort(key=lambda x: x[1], reverse=True)

        # Tìm điểm và thứ hạng của SBD
        for idx, (curr_sbd, score) in enumerate(total_scores, start=1):
            if curr_sbd == sbd:
                your_score = score
                rank = idx
                break

    return render_template("index.html", rank=rank, score=your_score)

if __name__ == "__main__":
    app.run(debug=True)

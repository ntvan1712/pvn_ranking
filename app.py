from flask import Flask, render_template, request
from openpyxl import load_workbook
from datetime import datetime

app = Flask(__name__)

COLS = {
    "van": 2,   # Cột B
    "toan": 3,  # Cột C
    "anh": 4,   # Cột D
    "tong": 5   # Cột E
}

RANK_COLS = {
    "van": 6,   # Cột F
    "toan": 7,  # Cột G
    "anh": 8,   # Cột H
    "tong": 9   # Cột I
}

SBD_COL = 1    # Cột A

SUPPORTED_SCHOOLS = {
    "20": "Nguyễn Đức Thuận",
    "21": "Hoàng Văn Thụ",  
    "22": "Lương Thế Vinh",
    "24": "Tống Văn Trân",  
    "26": "Phạm Văn Nghị",
    "27": "Đại An"
}

SCHOOL_CANDIDATES = {
    "20": 442,
    "21": 425,
    "22": 425,
    "24": 604,
    "26": 494,
    "27": 317,
}


@app.route("/", methods=["GET", "POST"])
def index():
    ranks = {}
    scores = {}
    sbd = ""
    error = ""
    school_name = ""
    total_candidates = 0

    if request.method == "POST":
        sbd = request.form["sbd"].strip()
        print(f"[{datetime.now()}] Tra cứu SBD: {sbd}")

        if len(sbd) < 2:
            error = "Mã số báo danh không hợp lệ."
        else:
            school_code = sbd[:2]
            if school_code not in SUPPORTED_SCHOOLS:
                error = "Chỉ hỗ trợ các trường Đại An, Phạm Văn Nghị, Lương Thế Vinh, Nguyễn Đức Thuận, Hoàng Văn Thụ, Tống Văn Trân."
            else:
                school_name = SUPPORTED_SCHOOLS[school_code]
                excel_file = f"{school_code}.xlsx"
                total_candidates = SCHOOL_CANDIDATES[school_code]

                try:
                    wb = load_workbook(excel_file)
                    ws = wb.active

                    found = False
                    for row in ws.iter_rows(min_row=2):
                        row_sbd = str(row[SBD_COL - 1].value).strip()
                        if row_sbd == sbd:
                            found = True
                            for key in COLS:
                                raw_score = row[COLS[key] - 1].value
                                rank_value = row[RANK_COLS[key] - 1].value
                                try:
                                    scores[key] = float(str(raw_score).replace(",", "."))
                                except:
                                    scores[key] = raw_score
                                ranks[key] = rank_value
                            break

                    if not found:
                        error = "Không tìm thấy mã số báo danh này."
                except FileNotFoundError:
                    error = f"Không tìm thấy dữ liệu cho trường {school_name}."

    return render_template(
        "index.html",
        ranks=ranks,
        scores=scores,
        sbd=sbd,
        error=error,
        school_name=school_name,
        total_candidates=total_candidates
    )


if __name__ == "__main__":
    app.run(debug=True)

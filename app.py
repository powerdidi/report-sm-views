from flask import Flask, render_template

import pandas as pd

from cleaner import TITLE_COL, VISIT_DIFF_COL


app = Flask(__name__)

MONTHLY_REPORT_FILE = "monthly_report.xlsx"


def to_persian_number(value):
    english = "0123456789"
    persian = "۰۱۲۳۴۵۶۷۸۹"

    return str(value).translate(
        str.maketrans(english, persian)
    )

def to_persian_rank(value):
    english = "0123456789"
    persian = "۰۱۲۳۴۵۶۷۸۹"

    return str(value).translate(
        str.maketrans(english, persian)
    )

@app.route("/")
def index():

    df = pd.read_excel(
        MONTHLY_REPORT_FILE,
        sheet_name="Top 10"
    )

    df = df[[
        TITLE_COL,
        VISIT_DIFF_COL
    ]]

    df[VISIT_DIFF_COL] = (
        df[VISIT_DIFF_COL]
        .fillna(0)
        .astype(int)
        .apply(to_persian_number)
    )

    rows = df.values.tolist()

    return render_template(
    "index.html",
    rows=rows,
    to_persian_rank=to_persian_rank
)


if __name__ == "__main__":
    app.run(debug=True)

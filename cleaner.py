import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import jdatetime





IMPORT_FILE = "data/import.xlsx"
HISTORY_FILE = "data/history.xlsx"
MONTHLY_REPORT_FILE = "monthly_report.xlsx"
HISTORY_New = "data/history_new.xlsx"


VISIT_DIFF_COL = "تعداد بازدید"
GROWTH_COL = "درصد رشد"
TITLE_COL = "عنوان"
TYPE_COL = "نوع خدمت"
VISIT_COL = "تعداد بازدید"
ELECTRONIC_SERVICE = "خدمات الکترونیکی"
CATEGORY_COL = "طبقه بندي"

EXCLUDED_CATEGORIES = [
    "مجوز هاي كسب و كار ( نمايشي )",
    "خدمات پركاربردمشهد",
    "ساير (تجميع در يك صفحه)",
    "تست",]



pd_history = pd.read_excel(HISTORY_FILE)
pd_import = pd.read_excel(IMPORT_FILE)




#نرمال سازی
def normalize(text):
    if pd.isna(text):
        return ""
    text = str(text)
    text = text.replace("ي", "ی")
    text = text.replace("ك", "ک")
    text = text.replace("ة", "ه")
    text = text.replace("\u200c", " ")  # نیم‌فاصله
    text = " ".join(text.split())       # حذف فاصله‌های اضافی
    return text.strip()



#جدا کردن خدمات الکترونیکی
def filter_service(df):
    df = df.copy()

    df[TYPE_COL] = df[TYPE_COL].apply(normalize)
    df[TITLE_COL] = df[TITLE_COL].apply(normalize)
    df[CATEGORY_COL] = df[CATEGORY_COL].apply(normalize)
    # فقط خدمات الکترونیکی
    result = df[df[TYPE_COL] == ELECTRONIC_SERVICE].copy()
    # حذف بر اساس طبقه بندی
    excluded_categories = [normalize(item) for item in EXCLUDED_CATEGORIES]
    result = result[~result[CATEGORY_COL].isin(excluded_categories)]
    # فقط ستون‌های مورد نیاز
    result = result[[TITLE_COL, VISIT_COL]]
     # تجمیع خدمات تکراری
    result = result.groupby(TITLE_COL, as_index=False)[VISIT_COL].sum()
    return result


#بروزرسانی فایل history
def update_history(pd_history, report_pd):
    today = jdatetime.date.today().strftime("%Y-%m")
    pd_history[TITLE_COL] = pd_history[TITLE_COL].apply(normalize)
    report_pd[TITLE_COL]  = report_pd[TITLE_COL].apply(normalize)

    if today in pd_history.columns:
        pd_history.drop(columns = [today], inplace=True)
    report_pd = report_pd.rename(columns={"تعداد بازدید": today})

    pd_history = pd_history.merge(report_pd, on=TITLE_COL, how="outer")
    return pd_history


# محاسبه اختلاف و درصد رشد
def calculate_growth(pd_history):
    pd_growth = pd_history.copy()

    previous_month = pd_growth.columns[-2]
    current_month = pd_growth.columns[-1]

    pd_growth[VISIT_DIFF_COL] = (
        pd_growth[current_month] - pd_growth[previous_month]
    )

    growth = (
        pd_growth[VISIT_DIFF_COL]
        / pd_growth[previous_month].replace(0, float("nan"))
        * 100
    )

    pd_growth[GROWTH_COL] = growth.round().astype("Int64")

    pd_growth = pd_growth.sort_values(
        by=VISIT_DIFF_COL,
        ascending=False
    )

    result = pd_growth[[
        TITLE_COL,
        previous_month,
        current_month,
        VISIT_DIFF_COL,
        GROWTH_COL
    ]]

    return result



#قالب‌بندی فایل اکسل
def format_excel(file_name):
    wb = load_workbook(file_name)
    ws = wb.active
    ws.sheet_view.rightToLeft = True
    alignment = Alignment(horizontal="right", vertical="center")
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = alignment
    wb.save(file_name)
    wb.close()


df_electronic = filter_service(pd_import)

pd_history = update_history(pd_history, df_electronic)

pd_history.to_excel(HISTORY_New, index=False)
format_excel(HISTORY_New)

pd_growth = calculate_growth(pd_history)

pd_top10 = pd_growth.head(10)
with pd.ExcelWriter(MONTHLY_REPORT_FILE) as writer:
    pd_growth.to_excel(
        writer,
        sheet_name="گزارش کامل",
        index=False
    )

    pd_top10.to_excel(
        writer,
        sheet_name="Top 10",
        index=False
    )
format_excel(MONTHLY_REPORT_FILE)

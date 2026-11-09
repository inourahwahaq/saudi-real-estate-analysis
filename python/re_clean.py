"""Clean six quarters of Saudi real estate deals.

usage: python3 re_clean.py <raw_dir> <output_dir>
"""
import glob
import os
import sys

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font

raw_dir, out = sys.argv[1:3]

REGIONS = {
    "منطقة مكة المكرمه": "Makkah", "منطقة الرياض": "Riyadh", "منطقة الشرقية": "Eastern", "منطقة القصيم": "Qassim", "منطقة المدينة المنوره": "Madinah",
    "منطقة حائل": "Hail", "منطقة عسير": "Asir", "منطقة جازان": "Jazan", "منطقة تبوك": "Tabuk", "منطقة الجوف": "Al Jawf", "منطقة نجران": "Najran",
    "منطقة الحدود الشمالية": "Northern Borders", "منطقة الباحة": "Al Bahah",
}
CITIES = {
    "جده": "Jeddah", "الرياض": "Riyadh", "المدينة المنورة": "Madinah", "مكة المكرمة": "Makkah", "بريده": "Buraydah", "الدمام": "Dammam", "حائل": "Hail",
    "الخبر": "Khobar", "الطائف": "Taif", "الهفوف": "Hofuf", "خميس مشيط": "Khamis Mushait", "ابها": "Abha", "ينبع": "Yanbu", "الجبيل": "Jubail", "تبوك": "Tabuk",
    "عنيزة": "Unaizah", "عنيزه": "Unaizah", "الدرعيه": "Diriyah", "الدرعية": "Diriyah", "الخرج": "Al Kharj", "الظهران": "Dhahran", "نجران": "Najran", "القطيف": "Qatif",
    "سكاكا": "Sakaka", "جازان": "Jazan", "جيزان": "Jazan", "المبرز": "Mubarraz", "رابغ": "Rabigh", "عرعر": "Arar", "الباحة": "Al Bahah", "الباحه": "Al Bahah",
    "حفر الباطن": "Hafar Al Batin", "الاحساء": "Al Ahsa", "الأحساء": "Al Ahsa", "الزلفي": "Az Zulfi", "الرس": "Ar Rass", "صبيا": "Sabya", "بيشة": "Bishah", "بيشه": "Bishah",
}
CLASSES = {"سكني": "Residential", "تجاري": "Commercial", "زراعي": "Agricultural", "سكني تجاري": "Mixed"}
RAW_COLUMNS = ["region_ar", "city_ar", "neighbourhood_ar", "deal_ref", "date", "date_hijri", "class_ar", "properties", "price", "area"]
BANDS, BAND_LABELS = [0, 250e3, 500e3, 1e6, 2e6, np.inf], ["Under 250K", "250K-500K", "500K-1M", "1M-2M", "Over 2M"]
COLUMNS = ["deal_ref", "date", "month_start", "quarter", "year", "region", "city", "neighbourhood", "class", "properties", "single_property", "price", "area", "price_per_m2", "price_per_m2_valid", "price_band", "region_ar", "city_ar", "neighbourhood_ar", "class_ar", "date_hijri", "source_quarter"]


def write_sheet(workbook, name, df):
    ws = workbook.create_sheet(name)
    ws.freeze_panes = "A2"
    header = [WriteOnlyCell(ws, value=str(c)) for c in df.columns]
    for cell in header:
        cell.font = Font(bold=True)
    ws.append(header)
    for row in df.itertuples(index=False):
        ws.append([v.date() if isinstance(v, pd.Timestamp) else v.item() if hasattr(v, "item") else v for v in row])


# Load the quarterly files
frames = []
for path in sorted(glob.glob(f"{raw_dir}/*.csv")):
    frame = pd.read_csv(path, encoding="utf-8-sig", dtype=str)
    frame.columns = RAW_COLUMNS
    frames.append(frame.assign(source_quarter=os.path.basename(path)[:-4]))
raw = pd.concat(frames, ignore_index=True)

# Clean rows, numbers and dates
d = raw.dropna(subset=["deal_ref"]).copy()
assert not d.deal_ref.duplicated().any()
for col in ["price", "area", "properties"]:
    d[col] = pd.to_numeric(d[col].str.replace(",", "", regex=False).str.strip())
d["date"] = pd.to_datetime(d.date)

# Add English labels
d["region"] = d.region_ar.str.strip().map(REGIONS)
d["class"] = d.class_ar.str.strip().map(CLASSES)
assert d[["region", "class"]].notna().all().all()
d["city"] = d.city_ar.str.strip().map(CITIES).fillna(d.city_ar.str.strip())
d["neighbourhood"] = d.neighbourhood_ar.str.split("/", n=1).str[1].str.strip()

# Add date and price fields
d["quarter"] = d.date.dt.to_period("Q").astype(str).str.replace("Q", " Q")
d["year"] = d.date.dt.year
d["month_start"] = d.date.dt.to_period("M").dt.to_timestamp()
d["price_per_m2"] = d.price / d.area
d["single_property"] = d.properties == 1
d["price_band"] = pd.cut(d.price, bins=BANDS, labels=BAND_LABELS, right=False).astype(str)

# One price can cover several plots, so price per m2 is only valid on single-property deals.
# The most extreme 1% at each end is flagged by class. Those rows still count in deals and value.
d["price_per_m2_valid"] = False
bounds = {}
for name, group in d[d.single_property].groupby("class"):
    low, high = group.price_per_m2.quantile([.01, .99])
    bounds[name] = f"{low:.2f} to {high:.2f}"
    d.loc[group.index, "price_per_m2_valid"] = group.price_per_m2.between(low, high)

clean = d[COLUMNS].sort_values(["date", "deal_ref"]).reset_index(drop=True)
valid_residential = clean[clean.price_per_m2_valid & (clean["class"] == "Residential")]
quarterly = clean.groupby("quarter").agg(deals=("deal_ref", "size"), value_sar=("price", "sum")).join(valid_residential.groupby("quarter").price_per_m2.median().rename("median_residential_price_per_m2")).reset_index()

# Cleaning log and summary for the Excel file
log = pd.DataFrame([
    ("all", "Removed fully empty rows", len(raw) - len(d)),
    ("deal_ref", "Checked for duplicate deals, none found", 0),
    ("price, area, properties, date", "Converted text to numbers and dates", len(d)),
    ("region, class, city, neighbourhood", "Added English labels and split the neighbourhood from the city", len(d)),
    ("price_per_m2_valid", "Flagged the most extreme 1% at each end by class, for price per m2 only", int((~d.price_per_m2_valid).sum())),
], columns=["field", "action", "rows_affected"])
summary = pd.DataFrame([
    ("Source rows", len(raw)), ("Clean rows", len(clean)),
    ("Period", f"{clean.date.min().date()} to {clean.date.max().date()}"),
    ("Regions", clean.region.nunique()), ("Cities", clean.city.nunique()), ("Neighbourhoods", clean.neighbourhood_ar.nunique()),
    ("Total value (SAR)", int(clean.price.sum())),
    ("Price per m2 bounds by class", "; ".join(f"{k}: {v}" for k, v in bounds.items())),
    ("Note", "The source describes these as the main approved sale deals, so totals may not cover every sale in the Kingdom."),
], columns=["item", "value"])

# Save results
clean.to_csv(f"{out}/real_estate_clean.csv", index=False, date_format="%Y-%m-%d")
workbook = Workbook(write_only=True)
sheets = {"Summary": summary, "Cleaning_Log": log, "Quarterly": quarterly, "Region_Map": pd.DataFrame(sorted(REGIONS.items(), key=lambda x: x[1]), columns=["region_ar", "region"]), "Clean": clean}
for name, df in sheets.items():
    write_sheet(workbook, name, df)
workbook.save(f"{out}/real_estate_cleaning.xlsx")

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

from utils.state import init_state

init_state()

st.title("Dashboard")

if not st.session_state.get("cleaning_done"):
    st.info("Run Cleaning first.")
    st.stop()

results = st.session_state.get("results", {})
if not results:
    st.warning("No results found.")
    st.stop()

CURRENCY = "TND"

# -----------------------------
# Helpers: build "Global Vehicle List" from brand FINAL workbooks in session_state
# -----------------------------
def _normalize_str(x) -> str:
    return ("" if x is None else str(x)).replace("\u00A0", " ").strip()

def _find_vehicle_list_sheet(xls: pd.ExcelFile, brand: str) -> str | None:
    # The exporter names it like: "{brand} Vehicle List"
    preferred = f"{brand} Vehicle List"
    if preferred in xls.sheet_names:
        return preferred
    # fallback: contains "vehicle list"
    for s in xls.sheet_names:
        if "vehicle list" in s.lower():
            return s
    return None

def load_global_vehicle_list_from_session(results_dict: dict) -> pd.DataFrame:
    rows = []
    for brand, res in results_dict.items():
        b = res.get("final_xlsx")
        if not b:
            continue

        xls = pd.ExcelFile(BytesIO(b))
        sheet = _find_vehicle_list_sheet(xls, brand)
        if not sheet:
            continue

        df = pd.read_excel(xls, sheet_name=sheet)
        df = df.rename(columns={
            "Véhicule ID": "VehicleID",
            "Vehicule ID": "VehicleID",
            "Total HTVA": "TotalHTVA",
            "Total": "TotalHTVA",
        })

        if "VehicleID" not in df.columns or "TotalHTVA" not in df.columns:
            continue

        df["VehicleID"] = df["VehicleID"].apply(_normalize_str)
        df["TotalHTVA"] = pd.to_numeric(df["TotalHTVA"], errors="coerce")

        # drop footer rows like "TOTAL of VEHICLES", blanks, NaNs
        df = df.dropna(subset=["TotalHTVA"]).copy()
        df = df[df["VehicleID"].ne("")].copy()
        df = df[~df["VehicleID"].str.contains("TOTAL", case=False, na=False)].copy()

        # add brand label for later aggregation
        df["Brand"] = brand
        rows.append(df[["VehicleID", "TotalHTVA", "Brand"]])

    if not rows:
        return pd.DataFrame(columns=["VehicleID", "TotalHTVA", "Brands"])

    all_df = pd.concat(rows, ignore_index=True)

    # Aggregate per vehicle across brands
    agg = (
        all_df.groupby("VehicleID", as_index=False)
              .agg(
                  TotalHTVA=("TotalHTVA", "sum"),
                  Brands=("Brand", lambda s: ", ".join(sorted(set(map(str, s)))))
              )
    )

    # round like your excel outputs
    agg["TotalHTVA"] = pd.to_numeric(agg["TotalHTVA"], errors="coerce").fillna(0.0).round(3)
    agg["Brands"] = agg["Brands"].astype(str).replace({"": "Unknown"}).fillna("Unknown")

    return agg


dfv = load_global_vehicle_list_from_session(results)

if dfv.empty:
    st.error("Could not build Global Vehicle List from brand final workbooks. Check that each brand produced a FINAL file with a '* Vehicle List' sheet.")
    st.stop()

# -----------------------------
# Shared computations
# -----------------------------
def split_brands(s: str) -> list[str]:
    parts = [p.strip() for p in str(s).split(",")]
    parts = [p for p in parts if p]
    return parts if parts else ["Unknown"]

total_cost = float(dfv["TotalHTVA"].sum()) if len(dfv) else 0.0
n_vehicles = int(dfv["VehicleID"].nunique()) if len(dfv) else 0
avg_cost = float(dfv["TotalHTVA"].mean()) if len(dfv) else 0.0
median_cost = float(dfv["TotalHTVA"].median()) if len(dfv) else 0.0
min_cost = float(dfv["TotalHTVA"].min()) if len(dfv) else 0.0
max_cost = float(dfv["TotalHTVA"].max()) if len(dfv) else 0.0
cost_range = max_cost - min_cost

dfv_sorted = dfv.sort_values("TotalHTVA", ascending=False).reset_index(drop=True)
top10 = dfv_sorted.head(10).copy()
top10_cost = float(top10["TotalHTVA"].sum()) if len(top10) else 0.0
top10_pct = (top10_cost / total_cost * 100) if total_cost else 0.0

# Brand allocated cost (split multi-brand equally)
brand_rows = dfv.copy()
brand_rows["BrandList"] = brand_rows["Brands"].apply(split_brands)
brand_rows["nBrands"] = brand_rows["BrandList"].apply(len)
brand_rows = brand_rows.explode("BrandList")
brand_rows["Brand"] = brand_rows["BrandList"].astype(str).str.strip()
brand_rows["AllocatedCost"] = brand_rows["TotalHTVA"] / brand_rows["nBrands"]

brand_agg = (
    brand_rows.groupby("Brand", as_index=False)
              .agg(TotalHTVA=("AllocatedCost", "sum"))
              .sort_values("TotalHTVA", ascending=False)
)
brand_agg["Percent"] = ((brand_agg["TotalHTVA"] / total_cost) * 100) if total_cost else 0.0
brand_agg["Percent"] = brand_agg["Percent"].round(1)
brand_agg["TotalHTVA"] = brand_agg["TotalHTVA"].round(3)

# Distribution bins (same as your script)
BINS = [0, 500, 1000, 2500, 5000, np.inf]
BIN_LABELS = ["0–500", "500–1000", "1000–2500", "2500–5000", "5000+"]

dfv["CostBucket"] = pd.cut(
    dfv["TotalHTVA"],
    bins=BINS,
    labels=BIN_LABELS,
    include_lowest=True,
    right=False
)

dist_agg = (
    dfv.groupby("CostBucket", as_index=False, observed=False)
       .agg(Vehicles=("VehicleID", "nunique"))
)

# Concentration / risk
top5 = dfv_sorted.head(5).copy()
top5_cost = float(top5["TotalHTVA"].sum()) if len(top5) else 0.0
pct_top5 = (top5_cost / total_cost * 100) if total_cost else 0.0
pct_top10 = (top10_cost / total_cost * 100) if total_cost else 0.0

top_n = int(np.ceil(0.2 * len(dfv_sorted))) if len(dfv_sorted) else 0
top20_sum = float(dfv_sorted.head(top_n)["TotalHTVA"].sum()) if top_n else 0.0
pareto_ratio = (top20_sum / total_cost * 100) if total_cost else 0.0

THRESHOLD_2000 = 2000
THRESHOLD_5000 = 5000
count_above_2000 = int(dfv.loc[dfv["TotalHTVA"] > THRESHOLD_2000, "VehicleID"].nunique())
count_above_5000 = int(dfv.loc[dfv["TotalHTVA"] > THRESHOLD_5000, "VehicleID"].nunique())
pct_above_5000 = (count_above_5000 / n_vehicles * 100) if n_vehicles else 0.0

# High-cost ratio by brand (vehicle counts)
risk_rows = dfv.copy()
risk_rows["BrandList"] = risk_rows["Brands"].apply(split_brands)
risk_rows = risk_rows.explode("BrandList")
risk_rows["Brand"] = risk_rows["BrandList"].astype(str).str.strip()

brand_total_vehicles = risk_rows.groupby("Brand", as_index=False).agg(Vehicles=("VehicleID", "nunique"))
brand_high_vehicles = (
    risk_rows[risk_rows["TotalHTVA"] > THRESHOLD_5000]
    .groupby("Brand", as_index=False)
    .agg(HighCostVehicles=("VehicleID", "nunique"))
)

brand_risk = brand_total_vehicles.merge(brand_high_vehicles, on="Brand", how="left")
brand_risk["HighCostVehicles"] = brand_risk["HighCostVehicles"].fillna(0).astype(int)
brand_risk["HighCostRatioPct"] = (brand_risk["HighCostVehicles"] / brand_risk["Vehicles"] * 100).replace([np.inf, -np.inf], 0.0)
brand_risk["HighCostRatioPct"] = brand_risk["HighCostRatioPct"].round(1)
brand_risk = brand_risk.sort_values("HighCostRatioPct", ascending=False).reset_index(drop=True)

# -----------------------------
# UI
# -----------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "Vehicle Maintenance (full)",
    "Core Cost KPIs",
    "Concentration & Risk",
    "1-Page Dashboard",
])

def kpi_row(items):
    cols = st.columns(len(items))
    for c, (label, val) in zip(cols, items):
        c.metric(label, val)

with tab1:
    st.subheader("Vehicle Maintenance Dashboard")

    kpi_row([
        ("Total Repair Cost", f"{total_cost:,.3f} {CURRENCY}"),
        ("Average Cost", f"{avg_cost:,.2f} {CURRENCY}"),
        ("Median Cost", f"{median_cost:,.2f} {CURRENCY}"),
        ("% Cost from Top 10", f"{top10_pct:.1f}%"),
    ])

    c1, c2 = st.columns(2)
    with c1:
        fig_brand = px.bar(brand_agg, x="Brand", y="TotalHTVA", title=f"Cost by Brand ({CURRENCY})")
        fig_brand.update_layout(height=360, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig_brand, use_container_width=True)

    with c2:
        fig_dist = px.bar(dist_agg, x="CostBucket", y="Vehicles", title="Cost Distribution (Vehicle Count)")
        fig_dist.update_layout(height=360, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig_dist, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        fig_top10 = px.bar(
            top10.sort_values("TotalHTVA", ascending=True),
            x="TotalHTVA", y="VehicleID", orientation="h",
            title=f"Top 10 Most Expensive Vehicles ({CURRENCY})"
        )
        fig_top10.update_layout(height=460, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig_top10, use_container_width=True)

    with c4:
        st.markdown("**Top 10 Most Expensive Vehicles**")
        st.dataframe(top10[["VehicleID", "TotalHTVA", "Brands"]], use_container_width=True, height=460)

    st.markdown("**Cost by Brand (with %)**")
    st.dataframe(brand_agg[["Brand", "TotalHTVA", "Percent"]], use_container_width=True)
    st.caption("Note: If a vehicle has multiple brands, its cost is split equally across those brands.")

with tab2:
    st.subheader("Core Cost KPIs")

    kpi_row([
        ("Total Repair Cost", f"{total_cost:,.3f} {CURRENCY}"),
        ("Total Vehicles", f"{n_vehicles:,}"),
        ("Average per Vehicle", f"{avg_cost:,.3f} {CURRENCY}"),
        ("Median per Vehicle", f"{median_cost:,.3f} {CURRENCY}"),
    ])
    kpi_row([
        ("Min Cost", f"{min_cost:,.3f} {CURRENCY}"),
        ("Max Cost", f"{max_cost:,.3f} {CURRENCY}"),
        ("Range (Max-Min)", f"{cost_range:,.3f} {CURRENCY}"),
        ("Top 10 %", f"{top10_pct:.1f}%"),
    ])

    st.markdown("**Top 10 Vehicles**")
    st.dataframe(top10[["VehicleID", "TotalHTVA", "Brands"]], use_container_width=True)

with tab3:
    st.subheader("Concentration & Risk KPIs")

    kpi_row([
        ("% of Total Cost from Top 5", f"{pct_top5:.1f}%"),
        ("% of Total Cost from Top 10", f"{pct_top10:.1f}%"),
        ("Top 20% Cost Contribution (Pareto)", f"{pareto_ratio:.1f}%"),
    ])
    kpi_row([
        (f"Vehicles Above {THRESHOLD_2000:,}", f"{count_above_2000:,}"),
        (f"Vehicles Above {THRESHOLD_5000:,}", f"{count_above_5000:,}"),
        (f"% Vehicles Above {THRESHOLD_5000:,}", f"{pct_above_5000:.1f}%"),
    ])

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Top 5 Most Expensive Vehicles**")
        st.dataframe(top5[["VehicleID", "TotalHTVA", "Brands"]], use_container_width=True)
        st.markdown("**Top 10 Most Expensive Vehicles**")
        st.dataframe(top10[["VehicleID", "TotalHTVA", "Brands"]], use_container_width=True)

    with c2:
        fig_brand_risk = px.bar(
            brand_risk.sort_values("HighCostRatioPct", ascending=True),
            x="HighCostRatioPct", y="Brand", orientation="h",
            title=f"High-Cost Vehicle Ratio by Brand (> {THRESHOLD_5000:,})"
        )
        fig_brand_risk.update_layout(height=520, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig_brand_risk, use_container_width=True)

    st.markdown("**High-Cost Vehicle Ratio by Brand**")
    st.dataframe(brand_risk[["Brand", "Vehicles", "HighCostVehicles", "HighCostRatioPct"]], use_container_width=True)
    st.caption("Note: A vehicle with multiple brands counts toward each listed brand for the ratio.")

with tab4:
    st.subheader("1-Page Dashboard")

    kpi_row([
        ("Total Repair Cost", f"{total_cost:,.3f} {CURRENCY}"),
        ("Total Vehicles", f"{n_vehicles:,}"),
        ("Avg Cost / Vehicle", f"{avg_cost:,.3f} {CURRENCY}"),
        ("% Cost from Top 10", f"{top10_pct:.1f}%"),
    ])

    # Avg cost per brand (allocated)
    brand_kpis = (
        brand_rows.groupby("Brand", as_index=False)
                  .agg(TotalCost=("AllocatedCost", "sum"),
                       AvgCost=("AllocatedCost", "mean"))
                  .sort_values("TotalCost", ascending=False)
    )
    brand_kpis["PctOfTotal"] = (brand_kpis["TotalCost"] / total_cost * 100) if total_cost else 0.0
    brand_kpis["PctOfTotal"] = brand_kpis["PctOfTotal"].round(1)
    brand_kpis["TotalCost"] = brand_kpis["TotalCost"].round(3)
    brand_kpis["AvgCost"] = brand_kpis["AvgCost"].round(3)

    c1, c2, c3 = st.columns(3)
    with c1:
        fig_brand_total = px.bar(brand_kpis, x="Brand", y="TotalCost", title=f"Cost by Brand ({CURRENCY})")
        fig_brand_total.update_layout(height=330, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig_brand_total, use_container_width=True)
    with c2:
        fig_brand_avg = px.bar(
            brand_kpis.sort_values("AvgCost", ascending=False),
            x="Brand", y="AvgCost",
            title=f"Avg Cost per Brand ({CURRENCY})"
        )
        fig_brand_avg.update_layout(height=330, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig_brand_avg, use_container_width=True)
    with c3:
        fig_dist2 = px.bar(dist_agg, x="CostBucket", y="Vehicles", title="Cost Distribution (Vehicle Count)")
        fig_dist2.update_layout(height=330, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig_dist2, use_container_width=True)

    c4, c5 = st.columns(2)
    with c4:
        st.markdown("**Cost by Brand (% included)**")
        st.dataframe(brand_kpis[["Brand", "TotalCost", "PctOfTotal", "AvgCost"]], use_container_width=True, height=420)
        st.caption("Note: multi-brand vehicle cost is split equally across listed brands.")
    with c5:
        st.markdown("**Top 10 Most Expensive Vehicles**")
        st.dataframe(top10[["VehicleID", "TotalHTVA", "Brands"]], use_container_width=True, height=420)
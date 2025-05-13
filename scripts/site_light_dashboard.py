import pandas as pd, plotly.graph_objects as go
from plotly.subplots import make_subplots
import datetime as dt, pathlib

PRIMARY   = "#00C4B3"     # 메인 민트
ACCENT    = "#0079FF"     # 포인트 블루
CARD_BG   = "#FFFFFF"
PAGE_BG   = "#F4F7FB"
FONT_FAM  = "Inter"

# ① 현장 좌표 (대략값)
SITE_COORD = {
    "AGI": {"lat": 24.999, "lon": 54.790, "name": "Al Ghallan Island"},
    "DAS": {"lat": 25.142, "lon": 52.900, "name": "Das Island"},
    "MIR": {"lat": 24.110, "lon": 53.510, "name": "Mirfa"},
    "SHU": {"lat": 24.187, "lon": 52.620, "name": "Shuweihat"},
}

def load_status(path):
    df = pd.read_excel(path, sheet_name="STATUS")
    df.columns = [c.strip().replace("\n", " ") for c in df.columns]
    return df

def kpi_cards(fig, orders, qty, vendor, row=1):
    kpis = [("ORDERS", orders), ("QUANTITY", qty), ("VENDORS", vendor)]
    for i,(title,val) in enumerate(kpis,1):
        fig.add_trace(go.Indicator(
            mode="number", value=val,
            title={"text": f"<span style='font-size:12px'>{title}</span>"},
            number={"font":{"size":28, "color":ACCENT}},
        ), row=row, col=i)

def find_col(df, candidates):
    for c in candidates:
        for col in df.columns:
            if c.lower().replace(' ', '') == col.lower().replace(' ', '').replace('\n',''):
                return col
    raise KeyError(f"No matching column for: {candidates}")

def build_map():
    lats  = [v["lat"] for v in SITE_COORD.values()]
    lons  = [v["lon"] for v in SITE_COORD.values()]
    names = [v["name"] for v in SITE_COORD.values()]
    sites = list(SITE_COORD.keys())
    return go.Scattergeo(
        lat=lats, lon=lons, text=[f"{s} – {n}" for s,n in zip(sites,names)],
        mode="markers+text", marker=dict(size=12, color=PRIMARY),
        textposition="bottom center", hoverinfo="text")

site_cols = ["SHU.1", "MIR.1", "DAS.1", "AGI.1"]
def extract_site(row):
    for col in site_cols:
        if col in row and pd.notnull(row[col]) and row[col]:
            return col.split('.')[0]
    return "UNK"

def build_dashboard(status_path, out_html):
    df = load_status(status_path)
    # 자동 컬럼명 매칭
    col_pkg   = find_col(df, ["PACKAGE NO", "PKG", "PKG NO", "패키지번호"])
    col_ship  = find_col(df, ["SCT SHIP NO.", "SHIP NO", "SHIPNO"])
    col_vendor= find_col(df, ["VENDOR", "벤더"])
    col_ata   = find_col(df, ["ATA", "입항일"])
    # SITE 추출
    df["SITE"] = df.apply(extract_site, axis=1)
    col_site = "SITE"

    orders  = df[col_ship].nunique()
    qty     = df[col_pkg].count()
    vendors = df[col_vendor].nunique()

    fig = make_subplots(
        rows=3, cols=3,
        specs=[[{"type":"indicator"},{"type":"indicator"},{"type":"indicator"}],
               [ {"colspan":3,"type":"scattergeo"}, None, None],
               [ {"type":"table"},{"type":"heatmap"}, None]],
        row_heights=[0.20,0.45,0.35],
        vertical_spacing=0.05, horizontal_spacing=0.05)

    kpi_cards(fig, orders, qty, vendors)
    fig.add_trace(build_map(), row=2, col=1)
    fig.update_geos(
        row=2, col=1, showcountries=False, showcoastlines=False,
        scope="asia", projection_type="mercator",
        fitbounds="locations", bgcolor=PAGE_BG)

    tbl = (df.groupby(col_site)[col_pkg]
             .count().reset_index()
             .rename(columns={col_pkg:"Packages", col_site:"SITE"}))
    fig.add_trace(go.Table(
        header=dict(values=list(tbl.columns),
                    fill_color=CARD_BG, font=dict(color=ACCENT)),
        cells=dict(values=[tbl[c] for c in tbl.columns],
                   fill_color=PAGE_BG)), row=3, col=1)

    df["month"] = pd.to_datetime(df[col_ata]).dt.month
    hm = (df.groupby([col_site,"month"])[col_ship]
            .count().unstack(fill_value=0)
            .reindex(index=SITE_COORD.keys()).sort_index())
    fig.add_trace(go.Heatmap(
        z=hm.values, x=hm.columns, y=hm.index,
        colorscale=[[0, "#E0F9F7"], [1, PRIMARY]],
        showscale=False), row=3, col=2)

    fig.update_layout(
        template="plotly_white",
        paper_bgcolor=PAGE_BG, plot_bgcolor=CARD_BG,
        font=dict(family=FONT_FAM, color="#002B5B"),
        margin=dict(t=80, l=20, r=20, b=20),
        title=f"<b>HVDC Supply-Chain Dashboard</b> - {dt.datetime.now():%Y-%m-%d}",
        height=780)

    pathlib.Path(out_html).parent.mkdir(exist_ok=True)
    fig.write_html(out_html, include_plotlyjs="cdn")
    print("✅ Saved:", out_html)

if __name__ == "__main__":
    build_dashboard(
        status_path="data/HVDC-STATUS(20250513).xlsx",
        out_html="output/site_light_dashboard.html") 
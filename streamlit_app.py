import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(page_title="품목군별 발주/입고 대시보드", layout="wide")

st.title("품목군별 발주 / 입고율 대시보드")
st.caption("구분(품목군) 기준으로 발주현황, 입고율, 미입고 현황을 확인할 수 있습니다.")


def standardize_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    rename_map = {
        "품목 향": "구분",
        "정입고수량": "정입고수량",
        "발주단가": "단가",
        "발주금액": "금액",
    }

    existing_map = {k: v for k, v in rename_map.items() if k in df.columns and v not in df.columns}
    if existing_map:
        df = df.rename(columns=existing_map)

    return df


def pick_best_header_dataframe(uploaded_file, sheet_name=0):
    candidates = []
    for header_row in [0, 1]:
        try:
            tmp = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row)
            tmp = standardize_columns(tmp)
            score = sum(
                col in tmp.columns
                for col in ["발주일자", "거래처명", "품목코드", "품목명", "발주수량", "가입고수량", "미입고수량", "구분"]
            )
            candidates.append((score, tmp))
        except Exception:
            continue

    if not candidates:
        return pd.read_excel(uploaded_file, sheet_name=sheet_name)

    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]

def load_data(uploaded_file):
    if uploaded_file.name.endswith(".csv"):
        return standardize_columns(pd.read_csv(uploaded_file))

    xls = pd.ExcelFile(uploaded_file)

    if "3월출고_세일즈팩" in xls.sheet_names:
        return pick_best_header_dataframe(uploaded_file, sheet_name="3월출고_세일즈팩")

    return pick_best_header_dataframe(uploaded_file, sheet_name=0)

def preprocess(df):
    df = standardize_columns(df)

    if "구분" not in df.columns and "품목 향" in df.columns:
        df["구분"] = df["품목 향"]

    df = df.copy()

    required_cols = [
        "발주일자", "거래처명", "품목코드", "품목명",
        "발주수량", "가입고수량", "미입고수량", "구분"
    ]

    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"필수 컬럼이 없습니다: {', '.join(missing)}")
        st.stop()

    if "납품가능일자" not in df.columns:
        df["납품가능일자"] = pd.NaT
    if "단가" not in df.columns:
        df["단가"] = 0
    if "금액" not in df.columns:
        df["금액"] = 0
    if "발주번호" not in df.columns:
        df["발주번호"] = ""
    if "진행구분" not in df.columns:
        df["진행구분"] = ""
    if "생산완료수량" not in df.columns:
        df["생산완료수량"] = 0

    text_cols = ["진행구분", "발주번호", "거래처명", "품목코드", "품목명", "구분"]
    for c in text_cols:
        df[c] = df[c].fillna("").astype(str).str.strip()

    date_cols = ["발주일자", "납품가능일자"]
    for c in date_cols:
        df[c] = pd.to_datetime(df[c], errors="coerce")

    num_cols = ["발주수량", "가입고수량", "미입고수량", "생산완료수량", "단가", "금액"]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    if "정입고수량" in df.columns:
        df["정입고수량"] = pd.to_numeric(df["정입고수량"], errors="coerce").fillna(0)

    if (df["금액"] == 0).all():
        df["금액"] = df["발주수량"] * df["단가"]

    if "정입고수량" in df.columns and df["정입고수량"].sum() > 0:
        df["입고수량"] = df["정입고수량"]
    else:
        df["입고수량"] = (df["발주수량"] - df["미입고수량"]).clip(lower=0)

    df["정입고수량"] = df["입고수량"]
    df["입고율"] = (df["입고수량"] / df["발주수량"].replace(0, pd.NA) * 100).fillna(0)
    df["생산완료율"] = (df["생산완료수량"] / df["발주수량"].replace(0, pd.NA) * 100).fillna(0)

    def status(row):
        if row["입고수량"] <= 0 and row["가입고수량"] <= 0:
            return "미입고"
        if row["미입고수량"] > 0:
            return "부분입고"
        return "입고완료"

    df["입고상태"] = df.apply(status, axis=1)

    today = pd.Timestamp.today().normalize()
    df["지연여부"] = "정상"
    df.loc[
        df["납품가능일자"].notna() &
        (df["납품가능일자"] < today) &
        (df["입고상태"] != "입고완료"),
        "지연여부"
    ] = "지연"

    return df

def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="조회결과")
    output.seek(0)
    return output

with st.sidebar:
    st.header("파일 업로드")
    uploaded_file = st.file_uploader("엑셀 업로드", type=["xlsx", "xls", "csv"])
    st.caption("시트명: 3월출고_세일즈팩")

if uploaded_file is None:
    st.info("왼쪽에서 엑셀 파일을 업로드하세요.")
    st.stop()

raw_df = load_data(uploaded_file)
df = preprocess(raw_df)

st.subheader("조회 조건")

c1, c2, c3, c4 = st.columns(4)

with c1:
    group_options = sorted([x for x in df["구분"].dropna().unique() if x != ""])
    selected_groups = st.multiselect(
        "구분(품목군)",
        group_options,
        default=group_options
    )

with c2:
    vendor_options = sorted([x for x in df["거래처명"].dropna().unique() if x != ""])
    selected_vendors = st.multiselect(
        "거래처",
        vendor_options,
        default=vendor_options
    )

with c3:
    status_options = sorted(df["입고상태"].dropna().unique())
    selected_status = st.multiselect(
        "입고상태",
        status_options,
        default=status_options
    )

with c4:
    min_date = df["발주일자"].dropna().min()
    max_date = df["발주일자"].dropna().max()

    if pd.isna(min_date) or pd.isna(max_date):
        st.error("발주일자 데이터가 없습니다.")
        st.stop()

    selected_dates = st.date_input(
        "발주일자 기간",
        value=(min_date.date(), max_date.date()),
        min_value=min_date.date(),
        max_value=max_date.date()
    )

filtered = df[
    df["구분"].isin(selected_groups) &
    df["거래처명"].isin(selected_vendors) &
    df["입고상태"].isin(selected_status)
].copy()

if isinstance(selected_dates, tuple) and len(selected_dates) == 2:
    start_date, end_date = selected_dates
    filtered = filtered[
        (filtered["발주일자"].dt.date >= start_date) &
        (filtered["발주일자"].dt.date <= end_date)
    ]

if filtered.empty:
    st.warning("선택한 조건에 해당하는 데이터가 없습니다.")
    st.stop()

st.subheader("발주 KPI")

k1, k2, k3, k4 = st.columns(4)
k1.metric("발주건수", len(filtered))
k2.metric("총 발주수량", f"{int(filtered['발주수량'].sum()):,}")
k3.metric("총 입고수량", f"{int(filtered['입고수량'].sum()):,}")
k4.metric("평균 입고율", f"{filtered['입고율'].mean():.1f}%")

st.subheader("품목군별 입고 현황")

summary = filtered.groupby("구분").agg(
    발주품목수=("품목코드", "nunique"),
    발주건수=("품목코드", "size"),
    발주수량=("발주수량", "sum"),
    입고수량=("입고수량", "sum"),
    미입고수량=("미입고수량", "sum"),
    생산완료수량=("생산완료수량", "sum")
).reset_index()

summary["입고율"] = (summary["입고수량"] / summary["발주수량"].replace(0, pd.NA) * 100).fillna(0).round(1)
summary["생산완료율"] = (summary["생산완료수량"] / summary["발주수량"].replace(0, pd.NA) * 100).fillna(0).round(1)

total_row = {
    "구분": "TOTAL",
    "발주품목수": int(summary["발주품목수"].sum()),
    "발주건수": int(summary["발주건수"].sum()),
    "발주수량": float(summary["발주수량"].sum()),
    "입고수량": float(summary["입고수량"].sum()),
    "미입고수량": float(summary["미입고수량"].sum()),
    "생산완료수량": float(summary["생산완료수량"].sum()),
}
total_row["입고율"] = round((total_row["입고수량"] / total_row["발주수량"] * 100), 1) if total_row["발주수량"] else 0
total_row["생산완료율"] = round((total_row["생산완료수량"] / total_row["발주수량"] * 100), 1) if total_row["발주수량"] else 0

summary_display = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)

st.markdown("### 1. 세일즈팩 입고현황")
st.markdown("**1) 입고현황**")

summary_view = summary_display.copy()
summary_view["발주수량"] = summary_view["발주수량"].round(0).astype(int)
summary_view["입고수량"] = summary_view["입고수량"].round(0).astype(int)
summary_view["미입고수량"] = summary_view["미입고수량"].round(0).astype(int)
summary_view["생산완료수량"] = summary_view["생산완료수량"].round(0).astype(int)

summary_view["생산완료수량"] = summary_view["생산완료수량"].apply(lambda x: "-" if x == 0 else f"{x:,}")
summary_view["발주수량"] = summary_view["발주수량"].apply(lambda x: f"{x:,}")
summary_view["입고수량"] = summary_view["입고수량"].apply(lambda x: f"{x:,}")
summary_view["미입고수량"] = summary_view["미입고수량"].apply(lambda x: f"{x:,}")
summary_view["입고율"] = summary_view["입고율"].apply(lambda x: f"{x:.0f}%")
summary_view["생산완료율"] = summary_view["생산완료율"].apply(lambda x: "-" if x == 0 else f"{x:.0f}%")

summary_view = summary_view[
    ["구분", "발주수량", "입고수량", "미입고수량", "생산완료수량", "입고율", "생산완료율"]
]

st.table(summary_view)

st.subheader("품목군별 지표 시각화")

fig_qty = px.bar(
    summary,
    x="구분",
    y=["발주수량", "입고수량", "미입고수량", "생산완료수량"],
    barmode="group",
    text_auto=True
)
fig_qty.update_layout(
    legend_title_text="수량지표",
    yaxis_title="수량",
    xaxis_title="품목 구분"
)
st.plotly_chart(fig_qty, use_container_width=True)

fig_mix = go.Figure()
fig_mix.add_trace(go.Bar(name="발주수량", x=summary["구분"], y=summary["발주수량"]))
fig_mix.add_trace(go.Bar(name="입고수량", x=summary["구분"], y=summary["입고수량"]))
fig_mix.add_trace(go.Bar(name="미입고수량", x=summary["구분"], y=summary["미입고수량"]))
fig_mix.add_trace(go.Bar(name="생산완료수량", x=summary["구분"], y=summary["생산완료수량"]))
fig_mix.add_trace(
    go.Scatter(
        name="입고율(%)",
        x=summary["구분"],
        y=summary["입고율"],
        mode="lines+markers+text",
        text=[f"{v:.1f}%" for v in summary["입고율"]],
        textposition="top center",
        yaxis="y2"
    )
)
fig_mix.add_trace(
    go.Scatter(
        name="생산완료율(%)",
        x=summary["구분"],
        y=summary["생산완료율"],
        mode="lines+markers+text",
        text=[f"{v:.1f}%" for v in summary["생산완료율"]],
        textposition="top center",
        yaxis="y2"
    )
)
fig_mix.update_layout(
    barmode="group",
    xaxis_title="품목 구분",
    yaxis=dict(title="수량"),
    yaxis2=dict(title="입고율(%)", overlaying="y", side="right", range=[0, 110]),
    legend=dict(orientation="h", y=1.08, x=0),
)
st.plotly_chart(fig_mix, use_container_width=True)

st.subheader("상세 데이터")

detail_cols = [
    c for c in [
        "구분", "진행구분", "발주일자", "발주번호", "거래처명",
        "품목코드", "품목명", "발주수량", "가입고수량",
        "입고수량", "미입고수량", "생산완료수량", "금액", "입고율", "생산완료율", "입고상태", "지연여부"
    ] if c in filtered.columns
]

show_df = filtered[detail_cols].sort_values(["발주일자", "발주번호"], ascending=[False, True]).copy()
st.dataframe(show_df, use_container_width=True)

st.download_button(
    "조회결과 엑셀 다운로드",
    data=to_excel_bytes(show_df),
    file_name="발주입고현황.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

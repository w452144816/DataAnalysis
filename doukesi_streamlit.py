import streamlit as st
import pandas as pd


# import matplotlib.pyplot as plt
# import seaborn as sns
# from matplotlib import rcParams
#
# # 方法1：直接指定常用中文字体
# rcParams['font.sans-serif'] = ['SimHei']   # 黑体
# rcParams['axes.unicode_minus'] = False     # 解决负号显示问题

import re

def normalize_store_name(name: str) -> str:
    if not name:
        return ""

    name = name.lower()
    name = re.sub(r"[（）()]", "", name)   # 去括号
    name = re.sub(r"店|门店", "", name)    # 去“店”
    name = re.sub(r"\s+", "", name)        # 去空格
    return name

from rapidfuzz import process, fuzz

def build_store_mapping(df_mt, df_coupon, score_threshold=85):
    mappings = []

    coupon_names = df_coupon[["poi_id", "poi_name", "norm_name"]].drop_duplicates()

    for _, mt_row in df_mt[["poi_id", "poi_name", "norm_name"]].drop_duplicates().iterrows():
        match = process.extractOne(
            mt_row["norm_name"],
            coupon_names["norm_name"],
            scorer=fuzz.token_sort_ratio
        )

        if match and match[1] >= score_threshold:
            matched_row = coupon_names.loc[match[2]]    # label index ✅


            mappings.append({
                "mt_poi_id": mt_row["poi_id"],
                "mt_poi_name": mt_row["poi_name"],
                "coupon_poi_id": matched_row["poi_id"],
                "coupon_poi_name": matched_row["poi_name"],
                "score": match[1]
            })

    return pd.DataFrame(mappings)


from dousike_data_class import G_stats, G_orders, G_redeems
from dousike_data_meituan_class import G_meituan, MeituanOrder

def stats_to_dataframe(stats_dict):
    """
    stats_dict: dict[(store_id, stat_date), StoreDailyStat]
    """
    data = []
    for stat in stats_dict.values():
        data.append({
            "store_id": stat.store_id,
            "store_name": stat.store_name,
            "date": stat.stat_date,
            "order_cnt": stat.order_cnt,
            "order_with_lday_cnt": max(0, stat.order_cnt - stat.redeem_from_prev_days),
            "redeem_cnt": stat.redeem_cnt,
            "redeem_from_prev_days": stat.redeem_from_prev_days
        })
    df = pd.DataFrame(data)
    df["date"] = pd.to_datetime(df["date"])
    return df

# 假设 stats_dict 已经生成
# stats_dict = build_store_daily_stats(orders, redeems)
df = stats_to_dataframe(G_stats)


import pandas as pd

def meituan_orders_to_df(orders: list[MeituanOrder]) -> pd.DataFrame:
    df = pd.DataFrame([{
        "date": o.pay_time.date(),
        "poi_id": o.poi_id,
        "poi_name": o.poi_name,
        "order_cnt": 1,
        "gmv": o.total_fee / 100,
        "actual_gmv": o.actual_receipt_fee / 100,
        "has_refund": o.has_refund
    } for o in orders])

    df["date"] = pd.to_datetime(df["date"])
    return df
df_mt_daily_store = meituan_orders_to_df(G_meituan)


def store_stats_to_df(stats_dict) -> pd.DataFrame:
    rows = []

    for (_, _), stat in stats_dict.items():
        rows.append({
            "store_id": stat.store_id,
            "store_name": stat.store_name,
            "date": pd.to_datetime(stat.stat_date),
            "coupon_order_cnt": stat.order_cnt,
            "redeem_cnt": stat.redeem_cnt,
            "redeem_from_prev_days": stat.redeem_from_prev_days
        })

    return pd.DataFrame(rows)

df_coupon_daily_store = store_stats_to_df(G_stats)

df_coupon_daily_store = df_coupon_daily_store.rename(columns={
    "store_id": "poi_id",
    "store_name": "poi_name"
})

df_mt_daily_store["poi_id"] = df_mt_daily_store["poi_id"].astype(str)
df_coupon_daily_store["poi_id"] = df_coupon_daily_store["poi_id"].astype(str)
df_mt_daily_store = df_mt_daily_store.rename(columns={"order_cnt": "meituan_order_cnt"})
df_coupon_daily_store = df_coupon_daily_store.rename(columns={"order_cnt": "coupon_order_cnt"})

df_mt_daily_store["norm_name"] = (
    df_mt_daily_store["poi_name"].apply(normalize_store_name)
)

df_coupon_daily_store["norm_name"] = (
    df_coupon_daily_store["poi_name"].apply(normalize_store_name)
)

df_store_mapping = build_store_mapping(
    df_mt_daily_store,
    df_coupon_daily_store,
    score_threshold=85
)

poi_id_map = (
    df_store_mapping
    .set_index("coupon_poi_id")["mt_poi_id"]
    .to_dict()
)
df_coupon_daily_store["poi_id"] = (
    df_coupon_daily_store["poi_id"].map(poi_id_map)
)
df_coupon_daily_store = (
    df_coupon_daily_store
    .dropna(subset=["poi_id"])
)
df_unmatched = df_coupon_daily_store[df_coupon_daily_store["poi_id"].isna()]
mt_name_map = (
    df_store_mapping
    .set_index("mt_poi_id")["mt_poi_name"]
    .to_dict()
)

df_coupon_daily_store["poi_name"] = (
    df_coupon_daily_store["poi_id"].map(mt_name_map)
)

df_mt_daily_store = (
    df_mt_daily_store
    .groupby(["poi_id", "poi_name", "date"], as_index=False)
    .agg({
        "meituan_order_cnt": "sum"   # ⚠️ 关键
    })
)



st.title("店铺每日优惠券统计")

# -------------------------------
# 1️⃣ 选择店铺
# -------------------------------
stores = df["store_name"].unique()
selected_store = st.selectbox("选择店铺", stores)

df_store = df[df["store_name"] == selected_store].sort_values("date")

# -------------------------------
# 2️⃣ 显示表格
# -------------------------------
st.subheader("每日明细表")
st.dataframe(df_store)

# -------------------------------
# 3️⃣ 折线图：下单 & 核销趋势
# -------------------------------
st.subheader(f"{selected_store} 下单 & 核销趋势")

# 把日期设置为索引，列为折线图多条线
chart_line_data = df_store.set_index("date")[["order_cnt", "order_with_lday_cnt", "redeem_cnt", "redeem_from_prev_days"]]

st.line_chart(chart_line_data)

# -------------------------------
# 4️⃣ 柱状图：每日核销分布
# -------------------------------
st.subheader(f"{selected_store} 每日核销柱状图")

chart_bar_data = df_store.set_index("date")[["redeem_cnt", "redeem_from_prev_days"]]

st.bar_chart(chart_bar_data)

df_order_daily = (
    pd.DataFrame([{
        "date": o.order_date,
        "online_order_cnt": o.order_cnt
    } for o in G_orders])
)

df_order_daily["date"] = pd.to_datetime(df_order_daily["date"])

df_order_daily = (
    df_order_daily
    .groupby("date", as_index=False)
    .agg({"online_order_cnt": "sum"})
)
df_redeem_daily = (
    df
    .groupby("date", as_index=False)
    .agg({
        "redeem_cnt": "sum",
        "redeem_from_prev_days": "sum"
    })
)

df_redeem_daily["date"] = pd.to_datetime(df_redeem_daily["date"])

df_all = (
    df_order_daily
    .merge(df_redeem_daily, on="date", how="left")
    .fillna(0)
    .sort_values("date")
)

df_all["same_day_redeem_cnt"] = (
    df_all["redeem_cnt"] - df_all["redeem_from_prev_days"]
)

df_all["same_day_redeem_rate"] = (
    df_all["same_day_redeem_cnt"] / df_all["online_order_cnt"]
).fillna(0)

st.header("📊 全店铺每日汇总（线上下单口径）")

st.dataframe(
    df_all[[
        "date",
        "online_order_cnt",
        "redeem_cnt",
        "redeem_from_prev_days",
        "same_day_redeem_cnt",
        "same_day_redeem_rate"
    ]]
)

st.subheader("全店铺每日趋势")

st.line_chart(
    df_all
    .set_index("date")[[
        "online_order_cnt",
        "redeem_cnt",
        "redeem_from_prev_days"
    ]]
)

st.subheader("当日线上下单 vs 当日核销转化率")
st.line_chart(
    df_all
    .set_index("date")[["same_day_redeem_rate"]]
)
st.caption(
    "转化率 = （当日核销数 - 跨天核销数） / 当日线上下单数"
)


df_week = df_all.copy()

# 周一作为一周开始（最常用）
df_week["week"] = df_week["date"].dt.to_period("W-MON").apply(lambda x: x.start_time)

df_weekly = (
    df_week
    .groupby("week", as_index=False)
    .agg({
        "online_order_cnt": "sum",
        "redeem_cnt": "sum",
        "redeem_from_prev_days": "sum",
        "same_day_redeem_cnt": "sum"
    })
    .sort_values("week")
)

df_weekly["weekly_redeem_rate"] = (
    df_weekly["same_day_redeem_cnt"] / df_weekly["online_order_cnt"]
).fillna(0)

st.header("📅 周结算汇总（全店铺）")

st.subheader("周汇总表")

st.dataframe(
    df_weekly[[
        "week",
        "online_order_cnt",
        "redeem_cnt",
        "redeem_from_prev_days",
        "same_day_redeem_cnt",
        "weekly_redeem_rate"
    ]]
)
st.subheader("周趋势（下单 / 核销 / 跨天核销）")

st.line_chart(
    df_weekly
    .set_index("week")[[
        "online_order_cnt",
        "redeem_cnt",
        "redeem_from_prev_days"
    ]]
)

st.subheader("周转化率趋势（当周下单 → 当周核销）")

st.line_chart(
    df_weekly
    .set_index("week")[["weekly_redeem_rate"]]
)

st.caption(
    "周转化率 = 当周非跨天核销数 ÷ 当周线上下单数"
)


df_week = df_all.copy()

# 确保是 datetime
df_week["date"] = pd.to_datetime(df_week["date"])

# 计算当月 1 号
month_start = df_week["date"].values.astype("datetime64[M]")

# 计算第几天（从 0 开始）
day_offset = (df_week["date"] - month_start).dt.days

# 计算业务周起始日
df_week["week"] = month_start + pd.to_timedelta((day_offset // 7) * 7, unit="D")


df_weekly = (
    df_week
    .groupby("week", as_index=False)
    .agg({
        "online_order_cnt": "sum",
        "redeem_cnt": "sum",
        "redeem_from_prev_days": "sum",
        "same_day_redeem_cnt": "sum"
    })
    .sort_values("week")
)

df_weekly["weekly_redeem_rate"] = (
    df_weekly["same_day_redeem_cnt"] / df_weekly["online_order_cnt"]
).fillna(0)

st.header("📅 周结算汇总（按月 1 号起算）")

st.dataframe(
    df_weekly[[
        "week",
        "online_order_cnt",
        "redeem_cnt",
        "redeem_from_prev_days",
        "same_day_redeem_cnt",
        "weekly_redeem_rate"
    ]]
)

st.subheader("业务周趋势")

st.line_chart(
    df_weekly
    .set_index("week")[[
        "online_order_cnt",
        "redeem_cnt",
        "redeem_from_prev_days"
    ]]
)

st.subheader("业务周转化率")

st.line_chart(
    df_weekly
    .set_index("week")[["weekly_redeem_rate"]]
)

st.caption("周转化率 = 当周非跨天核销数 ÷ 当周线上下单数（以每月1号为起点）")


df_orders = pd.DataFrame([{
    "coupon_id": o.coupon_id,
    "order_date": pd.to_datetime(o.order_date)
} for o in G_orders])

df_redeems = pd.DataFrame([{
    "coupon_id": r.coupon_id,
    "redeem_date": pd.to_datetime(r.redeem_date)
} for r in G_redeems])

# 每张券只保留首次核销
df_first_redeem = (
    df_redeems
    .sort_values("redeem_date")
    .drop_duplicates("coupon_id", keep="first")
)

df_base = (
    df_orders
    .merge(df_first_redeem, on="coupon_id", how="left")
)

# 计算下单 → 核销的天数
df_base["days_to_redeem"] = (
    df_base["redeem_date"] - df_base["order_date"]
).dt.days

df_base["order_day"] = df_base["order_date"].dt.date
WINDOWS = [0, 1, 3, 7, 14, 30]
day_cohort = []

for day, g in df_base.groupby("order_day"):
    total = len(g)

    row = {
        "order_day": day,
        "order_cnt": total
    }

    for w in WINDOWS:
        row[f"redeem_{w}d"] = (
            (g["days_to_redeem"] <= w).sum()
        )
        row[f"rate_{w}d"] = (
            row[f"redeem_{w}d"] / total if total > 0 else 0
        )

    day_cohort.append(row)

df_day_cohort = pd.DataFrame(day_cohort).sort_values("order_day")

month_start = df_base["order_date"].values.astype("datetime64[M]")
day_offset = (df_base["order_date"] - month_start).dt.days
df_base["order_week"] = (
    month_start + pd.to_timedelta((day_offset // 7) * 7, unit="D")
)

week_cohort = []

for week, g in df_base.groupby("order_week"):
    total = len(g)

    row = {
        "order_week": week,
        "order_cnt": total
    }

    for w in WINDOWS:
        row[f"redeem_{w}d"] = (
            (g["days_to_redeem"] <= w).sum()
        )
        row[f"rate_{w}d"] = (
            row[f"redeem_{w}d"] / total if total > 0 else 0
        )

    week_cohort.append(row)

df_week_cohort = pd.DataFrame(week_cohort).sort_values("order_week")
st.header("📊 日下单 Cohort 转化率")

st.dataframe(
    df_day_cohort[
        ["order_day", "order_cnt"] +
        [f"rate_{w}d" for w in WINDOWS]
    ]
)
st.header("📅 周下单 Cohort 转化率（按月1号起算）")

st.dataframe(
    df_week_cohort[
        ["order_week", "order_cnt"] +
        [f"rate_{w}d" for w in WINDOWS]
    ]
)

# =========================
# 页面配置
# =========================
st.set_page_config(
    page_title="私域转化分析",
    layout="wide"
)

st.title("📊 美团 → 优惠券 私域转化分析")

# =========================
# 数据准备
# =========================
# df_mt_daily_store:
#   poi_id | poi_name | date | meituan_order_cnt
#
# df_coupon_daily_store:
#   poi_id | poi_name | date | coupon_order_cnt

df = (
    df_mt_daily_store
    .merge(
        df_coupon_daily_store[
            ["poi_id", "poi_name", "date", "coupon_order_cnt"]
        ],
        on=["poi_id", "poi_name", "date"],
        how="left"
    )
    .fillna(0)
)
print(df.columns)

df["date"] = pd.to_datetime(df["date"])

df["private_domain_rate"] = (
    df["coupon_order_cnt"] / df["meituan_order_cnt"]
)

df["private_domain_rate"] = (
    df["private_domain_rate"]
    .replace([float("inf")], 0)
    .fillna(0)
)

# =========================
# 筛选条件
# =========================
store_list = sorted(df["poi_name"].unique())

selected_store = st.selectbox(
    "选择店铺",
    store_list
)

min_date = df["date"].min().date()
max_date = df["date"].max().date()

date_range = st.date_input(
    "选择日期范围",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)

df_store = (
    df[
        (df["poi_name"] == selected_store) &
        (df["date"].dt.date >= date_range[0]) &
        (df["date"].dt.date <= date_range[1])
    ]
    .sort_values("date")
)

# =========================
# 核心指标
# =========================
col1, col2, col3 = st.columns(3)

with col1:
    st.metric(
        "美团下单数",
        int(df_store["meituan_order_cnt"].sum())
    )

with col2:
    st.metric(
        "优惠券下单数",
        int(df_store["coupon_order_cnt"].sum())
    )

with col3:
    total_mt = df_store["meituan_order_cnt"].sum()
    total_coupon = df_store["coupon_order_cnt"].sum()
    rate = total_coupon / total_mt if total_mt > 0 else 0

    st.metric(
        "私域转化率",
        f"{rate:.2%}"
    )

# =========================
# 趋势图
# =========================
st.subheader("📈 到店 vs 私域 下单趋势")

st.line_chart(
    df_store
    .set_index("date")[[
        "meituan_order_cnt",
        "coupon_order_cnt"
    ]]
)

st.subheader("📉 私域转化率趋势")

st.line_chart(
    df_store
    .set_index("date")[["private_domain_rate"]]
)

# =========================
# 明细表
# =========================
st.subheader("📋 每日明细数据")

st.dataframe(
    df_store[['date', 'meituan_order_cnt', 'coupon_order_cnt', 'private_domain_rate']],
    width="stretch"   # 原 use_container_width=True 的效果
)

# =========================
# 全店汇总
# =========================
st.divider()
st.header("🏬 全部店铺汇总（每日）")

df_all = (
    df
    .groupby("date", as_index=False)
    .agg({
        "meituan_order_cnt": "sum",
        "coupon_order_cnt": "sum"
    })
)

df_all["private_domain_rate"] = (
    df_all["coupon_order_cnt"] / df_all["meituan_order_cnt"]
).fillna(0)

st.subheader("📈 全店下单趋势")

st.line_chart(
    df_all
    .set_index("date")[[
        "meituan_order_cnt",
        "coupon_order_cnt"
    ]]
)

st.subheader("📉 全店私域转化率趋势")

st.line_chart(
    df_all
    .set_index("date")[["private_domain_rate"]]
)

import streamlit as st
import pandas as pd


# import matplotlib.pyplot as plt
# import seaborn as sns
# from matplotlib import rcParams
#
# # 方法1：直接指定常用中文字体
# rcParams['font.sans-serif'] = ['SimHei']   # 黑体
# rcParams['axes.unicode_minus'] = False     # 解决负号显示问题

from dousike_data_class import G_stats, G_orders, G_redeems

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

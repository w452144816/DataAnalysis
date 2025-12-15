import streamlit as st
import pandas as pd


# import matplotlib.pyplot as plt
# import seaborn as sns
# from matplotlib import rcParams
#
# # 方法1：直接指定常用中文字体
# rcParams['font.sans-serif'] = ['SimHei']   # 黑体
# rcParams['axes.unicode_minus'] = False     # 解决负号显示问题

from dousike_data_class import G_stats

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
chart_line_data = df_store.set_index("date")[["下单数", "核销数", "跨天下单核销"]]

st.line_chart(chart_line_data)

# -------------------------------
# 4️⃣ 柱状图：每日核销分布
# -------------------------------
st.subheader(f"{selected_store} 每日核销柱状图")

chart_bar_data = df_store.set_index("date")[["核销数", "跨天下单核销"]]

st.bar_chart(chart_bar_data)

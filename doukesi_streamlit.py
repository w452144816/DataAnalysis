import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns

import pandas as pd
from matplotlib import rcParams

# 方法1：直接指定常用中文字体
rcParams['font.sans-serif'] = ['SimHei']   # 黑体
rcParams['axes.unicode_minus'] = False     # 解决负号显示问题

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

# 1️⃣ 表格展示
st.subheader("每日明细表")
st.dataframe(df.sort_values(["store_id", "date"]))

# 2️⃣ 可选店铺选择
stores = df["store_name"].unique()
selected_store = st.selectbox("选择店铺", stores)

df_store = df[df["store_name"] == selected_store]

# 3️⃣ 折线图：下单与核销趋势
st.subheader(f"{selected_store} 下单 & 核销趋势")
plt.figure(figsize=(10,4))
sns.lineplot(df_store, x="date", y="order_cnt", label="下单数", marker="o")
sns.lineplot(df_store, x="date", y="redeem_cnt", label="核销数", marker="o")
sns.lineplot(df_store, x="date", y="redeem_from_prev_days", label="跨天下单核销", marker="o")
plt.xticks(rotation=45)
plt.ylabel("数量")
plt.xlabel("日期")
plt.legend()
st.pyplot(plt.gcf())

# 4️⃣ 柱状图：每天核销占比
st.subheader(f"{selected_store} 核销分布")
plt.figure(figsize=(10,4))
sns.barplot(df_store, x="date", y="redeem_cnt", color="skyblue", label="核销数")
sns.barplot(df_store, x="date", y="redeem_from_prev_days", color="orange", label="跨天下单核销")
plt.xticks(rotation=45)
plt.ylabel("数量")
plt.xlabel("日期")
plt.legend()
st.pyplot(plt.gcf())

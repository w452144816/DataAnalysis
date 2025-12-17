from dataclasses import dataclass
from datetime import date
from typing import Optional
from datetime import datetime
today_str = datetime.now().strftime("%Y-%m-%d")
print(today_str)

@dataclass
class Order:
    order_id: str
    coupon_id: str
    order_date: date
    order_cnt: int = 1

@dataclass
class Redeem:
    coupon_id: str
    store_id: str
    store_name: str
    redeem_date: date
    redeem_cnt: int = 1

@dataclass
class StoreDailyStat:
    store_id: str
    store_name: str
    stat_date: date

    order_cnt: int = 0                 # 当日下单
    redeem_cnt: int = 0                # 当日核销
    redeem_from_prev_days: int = 0     # 跨天下单的核销

from collections import defaultdict

def build_order_index(orders: list[Order]):
    idx = {}
    for o in orders:
        idx[o.coupon_id] = o
    return idx


import requests
from datetime import datetime, timedelta

def date_range(start_date: str, end_date: str):
    start = datetime.strptime(start_date, "%Y-%m-%d").date()
    end = datetime.strptime(end_date, "%Y-%m-%d").date()

    cur = start
    while cur <= end:
        yield cur.strftime("%Y-%m-%d")
        cur += timedelta(days=1)


def fetch_orders(start_date: str, end_date: str) -> list[Order]:
    orders: list[Order] = []

    for dt in date_range(start_date, end_date):
        try:
            resp = requests.get(
                "https://api.dousike.cyou/api/coupon/statistics",
                params={"date": dt},
                timeout=10,
                verify = False
            )
            resp.raise_for_status()

            data = resp.json().get("data", [])

            for d in data['createList']:
                orders.append(
                    Order(
                        order_id=str(d["user_id"]),
                        coupon_id=str(datetime.strptime(d["create_time"], "%Y-%m-%d %H:%M:%S").timestamp()),
                        order_date=datetime.strptime(d["create_time"].split(' ')[0], "%Y-%m-%d").date()
                    )
                )

        except Exception as e:
            # 生产环境建议换成日志
            print(f"[WARN] fetch_orders failed, date={dt}, err={e}")

    return orders

def fetch_redeems(start_date: str, end_date: str) -> list[Redeem]:
    redeems: list[Redeem] = []

    for dt in date_range(start_date, end_date):
        try:
            resp = requests.get(
                "https://api.dousike.cyou/api/coupon/statistics",  # 假设核销接口
                params={"date": dt},
                timeout=10,
                verify=False
            )
            resp.raise_for_status()
            data = resp.json().get("data", [])

            for d in data['useCount']:
                for r in d['child']:
                    redeems.append(
                        Redeem(
                            coupon_id=str(str(datetime.strptime(r["create_time"], "%Y-%m-%d %H:%M:%S").timestamp())),
                            store_id=str(d["redeem_store_id"]),
                            store_name=str(d["redeem_store_name"]),
                            redeem_date=datetime.strptime(r["use_time"].split(' ')[0], "%Y-%m-%d").date(),
                            redeem_cnt=r.get("count", 1)
                        )
                    )

        except Exception as e:
            print(f"[WARN] fetch_redeems failed, date={dt}, err={e}")

    return redeems

G_orders = fetch_orders("2025-12-01", today_str)
G_redeems = fetch_redeems("2025-12-01", today_str)



def build_store_daily_stats(
    orders: list[Order],
    redeems: list[Redeem]
) -> dict[tuple, StoreDailyStat]:

    stats = {}
    order_idx = build_order_index(orders)

    for r in redeems:
        # 1. 当日核销
        key = (r.store_id, r.redeem_date)
        stat = stats.setdefault(
            key,
            StoreDailyStat(store_id=r.store_id, store_name=r.store_name, stat_date=r.redeem_date)
        )
        stat.redeem_cnt += 1

        # 2. 找到对应的下单
        order = order_idx.get(r.coupon_id)
        if not order:
            continue

        # 3. 下单归因到店铺
        order_key = (r.store_id, order.order_date)
        order_stat = stats.setdefault(
            order_key,
            StoreDailyStat(store_id=r.store_id, store_name=r.store_name, stat_date=order.order_date)
        )
        order_stat.order_cnt += 1

        # 4. 判断是否跨天
        if order.order_date != r.redeem_date:
            stat.redeem_from_prev_days += 1

    return stats

G_stats = build_store_daily_stats(G_orders, G_redeems)

a = G_stats

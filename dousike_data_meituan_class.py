from dataclasses import dataclass
from datetime import datetime, timedelta
import requests
today_str = datetime.now().strftime("%Y-%m-%d")
print(today_str)

my_cookie = '_lxsdk_cuid=19b211721bcc8-0d879b08ae6f4a8-26061a51-384000-19b211721bcc8; uuid=a90fa4db2b584dcb9794.1765786906.1.0.0; mtcdn=K; _lx_utm=utm_source%3DBaidu%26utm_medium%3Dorganic; _lxsdk=19b211721bcc8-0d879b08ae6f4a8-26061a51-384000-19b211721bcc8; WEBDFPID=91522v461xv55209z8y3638vzv8125z980y70u5223297958yww7598v-1766021846163-1765935445702KWGWQWCfd79fef3d01d5e9aadc18ccd4d0c95071967; utm_source_rg=AM%2543PlOlO%25258; e_b_id_352126=9ff359676b20c9e05d616b884d72e624; loginToken=LQomDaCDUmapFeMyQLpgCHPSX2LvTcIC5Kxwty-jHsJ8m1iKQB2FsI0bEM1W4_RYWzXJRXwJ3PlqXdTcmTeVQA; merchantNo=88018712; logan_session_token=5o7n264keif6bxlfnwv0; home-payment-modal=[%2288018712%22]; _lxsdk_s=19b2a26ec84-a09-ce7-063%7C%7C53'
my_poiIds = [
        602116774,
        602242330,
        602247533,
        602240810,
        602446304,
        602208839,
        602363042,
        602269882,
        602412381,
        602392364,
        602373215,
        602549413,
        602275914,
        602535489,
        602401769,
        602693060,
        602640049,
        602676057,
        602610104,
        602661126
    ]

def parse_date_range(start_date: str, end_date: str):
    """
    输入:
        '2025-12-01', '2025-12-15'
    输出:
        datetime(2025,12,1,0,0,0),
        datetime(2025,12,15,23,59,59)
    """
    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    end_dt = datetime.strptime(end_date, "%Y-%m-%d") + timedelta(days=1) - timedelta(seconds=1)
    return start_dt, end_dt

def to_ms(dt: datetime) -> int:
    return int(dt.timestamp() * 1000)

@dataclass
class MeituanOrder:
    trade_no: str
    pay_time: datetime
    poi_id: int
    poi_name: str

    total_fee: int              # 单位：分
    actual_receipt_fee: int     # 实收金额
    has_refund: bool

    business_type: int
    business_type_name: str
    scene_name: str

class MeituanClient:

    def __init__(self, cookie: str):
        self.url = (
            "https://pos.meituan.com/web/api/v2/reports/pay-bill"
            "?yodaReady=h5&csecplatform=4&csecversion=4.1.1"
        )
        self.headers = {
            "Content-Type": "application/json",
            "appcode": "49",
            "model": "chrome",
            "cookie": cookie
        }

    def fetch_page(self, payload: dict) -> dict:
        resp = requests.post(self.url, json=payload, headers=self.headers, timeout=20)
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") != 0:
            raise RuntimeError(f"Meituan API error: {data}")
        return data["data"]

    def fetch_orders(
            self,
            start_date: str,
            end_date: str,
            poi_ids: list[int],
            page_size: int = 100
    ) -> list[MeituanOrder]:

        start_dt, end_dt = parse_date_range(start_date, end_date)

        page_no = 1
        orders = []

        while True:
            payload = {
                "startDate": int(start_dt.timestamp() * 1000),
                "endDate": int(end_dt.timestamp() * 1000),
                "beginDate": int(start_dt.timestamp() * 1000),
                "onlyNo": False,
                "pageNo": page_no,
                "pageSize": page_size,
                "poiIds": poi_ids,
                "aggrChannelType": 1,
                "permissionCode": 593
            }

            data = self.fetch_page(payload)

            for item in data.get("items", []):
                orders.append(MeituanOrder(
                    trade_no=str(item["tradeno"]),
                    pay_time=datetime.strptime(item["payTime"], "%Y-%m-%d %H:%M:%S"),
                    poi_id=item["poiId"],
                    poi_name=item["poiName"],
                    total_fee=item["totalFee"],
                    actual_receipt_fee=item["actualReceiptFee"],
                    has_refund=item["hasRefund"],
                    business_type=item["businessType"],
                    business_type_name=item["businessTypeName"],
                    scene_name=item["sceneName"]
                ))

            page = data["page"]
            if page_no >= page["totalPageSize"]:
                break

            page_no += 1

        return orders

G_meituan_client = MeituanClient(my_cookie)
G_meituan = G_meituan_client.fetch_orders("2025-12-01", today_str, my_poiIds)
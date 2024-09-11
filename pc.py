# -*- coding:utf-8*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import time
import requests
time1=time.time()
import pandas as pd
import  json


################定义数据结构列表存储数据####################
iname=[]
icard=[]

################循环发送请求解析数据#######################
for i in range(1,101):
    print('正在抓取第'+str(i)+"页.................................")
    url="https://sp0.baidu.com/8aQDcjqpAAV3otqbppnN2DJv/api.php?resource_id=6899&query=%E8%80%81%E8%B5%96&pn="+str(i*10)+"&ie=utf-8&oe=utf-8&format=json"
    head={
    "Host": "sp0.baidu.com",
    "Connection": "keep-alive",
    "User-Agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.101 Safari/537.36",
    "Accept": "*/*",
    "Referer": "https://www.baidu.com/s?ie=utf-8&f=3&rsv_bp=1&tn=95943715_hao_pg&wd=%E8%80%81%E8%B5%96&oq=%25E8%2580%2581%25E8%25B5%2596&rsv_pq=ec5e631d0003d8eb&rsv_t=b295wWZB5DEWWt%2FICZvMsf2TZJVPmof2YpTR0MpCszb28dLtEQmdjyBEidZohtPIr%2FBmMrB3&rqlang=cn&rsv_enter=0&prefixsug=%25E8%2580%2581%25E8%25B5%2596&rsp=0&rsv_sug9=es_0_1&rsv_sug=9",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "zh-CN,zh;q=0.8"
    }

    html=requests.get(url,headers=head).content
    html_json=json.loads(html)
    html_data=html_json['data']
    for each in html_data:
        k=each['result']
        for each in k:
            print(each['iname'],each['cardNum'])
            iname.append(each['iname'])
            icard.append(each['cardNum'])


#####################将数据组织成数据框###########################
data=pd.DataFrame({"name":iname,"IDCard":icard})

#################数据框去重####################################
data1=data.drop_duplicates()


# #########################写出数据到excel#########################################
# pd.DataFrame.to_excel(data1,"F:\\iname_icard.xlsx",header=True,encoding='gbk',index=False)
# time2=time.time()
# print u'ok,爬虫结束!'
# print u'总共耗时：'+str(time2-time1)+'s'
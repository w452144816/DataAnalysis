import datetime
import sys, os
import time
from dataclasses import dataclass, field
import json, time
import openpyxl
from thefuzz import fuzz

row_tag = [['转出日期', '日期'],'时间1','银行','转出账号','转出金额',['转入', '转入日期'],['转入时间', '时间', '时间2'],'付款账号','收到账号','转入金额','备注']
bill_id = 0
@dataclass(repr=False)
class personage:
    name: str = field(default=None)
    account: list = field(default=None)
    def __init__(self , name, account=None):
        self.name = name

@dataclass(repr=False)
class bill_item:
    id: int = field(default=None)
    date: datetime = field(default=None)
    payer_info: personage = field(default=None)
    payee_info: personage = field(default=None)
    type: int = field(default=None)
    figure: float = field(default=None)
    def __init__(self, date=None, payer_info=None, payee_info=None, type=None, figure=None):
        global bill_id
        self.id = bill_id
        bill_id +=1
        self.date = date
        self.payer_info = payer_info
        self.payee_info = payee_info
        self.type = type
        self.figure = figure

@dataclass(repr=False)
class personage_info(personage):
    bill: list[bill_item] = field(default=None)
    total_count: int = field(default=0)
    # name_another: str = field(default=None)

    def __init__(self, _name):
        super().__init__(_name)
        self.bill = []
    # def __post_init__(self):
    #     self.total_price = self.price * self.count
    #
    # def take(self):
    #     self.total_count = self.total_count - self.count
    # def addcount(self, _count):
    #     self.total_count = self.total_count + 1 * _count

def softID(_it, _itl, _score):
    res_score = {}
    for it in _itl:
        ratio_it = fuzz.partial_ratio(it, _it)
        res_score[it] = ratio_it

    max_key_value = max(res_score.items(), key=lambda item: item[1])
    # print(sorted_dict)
    return max_key_value[0], max_key_value[1]

def loadxlsx(_sheet, type, ac_info):
    if _sheet.max_row < 2:
        return None

    tag_map = {}
    tags = list(_sheet[2])
    index = 0
    for it in row_tag:
        for cell in tags[index:]:  # 获取第 2 行的所有单元格
            if cell.value and cell.value in it:
                if isinstance(it, list):
                    tag_map[it[-1]] = cell.column
                else:
                    tag_map[it] = cell.column
                index += 1
                break
    print(tag_map)
    for i in range(_sheet.max_row + 1):
        if i < 3:
            continue
        date1 = _sheet.cell(row=i, column=tag_map['日期']).value if '日期' in tag_map.keys() else None
        time1 = _sheet.cell(row=i, column=tag_map['时间1']).value if '时间1' in tag_map.keys() else None
        send1 = _sheet.cell(row=i, column=tag_map['银行']).value if '银行' in tag_map.keys() else None
        recv1 = _sheet.cell(row=i, column=tag_map['转出账号']).value if '转出账号' in tag_map.keys() else None
        p1 = _sheet.cell(row=i, column=tag_map['转出金额']).value if '转出金额' in tag_map.keys() else None
        date2 = _sheet.cell(row=i, column=tag_map['转入日期']).value if '转入日期' in tag_map.keys() else None
        time2 = _sheet.cell(row=i, column=tag_map['时间2']).value if '时间2' in tag_map.keys() else None
        send2 = _sheet.cell(row=i, column=tag_map['付款账号']).value if '付款账号' in tag_map.keys() else None
        recv2 = _sheet.cell(row=i, column=tag_map['收到账号']).value if '收到账号' in tag_map.keys() else None
        p2 = _sheet.cell(row=i, column=tag_map['转入金额']).value if '转入金额' in tag_map.keys() else None

        if p1 is None and p2 is None:
            continue
        try:
            p1=int(p1) if p1 is not None else None
            p2=int(p2) if p2 is not None else None
        except:
            print('{}用户的金额 的单元格 异常 : L {}'.format(ac_info.name, i))
        if not isinstance(p1, (int, float)) and not isinstance(p2, (int, float)):
            print('{}用户的金额 的单元格 异常 : L {}'.format(ac_info.name, i))
            continue

        if not isinstance(date1, datetime.datetime) and date1:
            if '.' in date1:
                date1 = datetime.datetime.strptime(date1, '%Y.%m.%d')
            elif '-' in date1:
                date1 = datetime.datetime.strptime(date1, '%Y-%m-%d')
            else:
                date1 = datetime.datetime.strptime(date1.replace("\t", "").replace('号','日'), '%Y年%m月%d日')
        if not isinstance(date2, datetime.datetime) and date2:
            if '.' in date2:
                date2 = datetime.datetime.strptime(date2, '%Y.%m.%d')
            elif '-' in date2:
                date2 = datetime.datetime.strptime(date2, '%Y-%m-%d')
            else:
                date2 = datetime.datetime.strptime(date2.replace("\t", "").replace('号','日'), '%Y年%m月%d日')

        if not isinstance(time1, datetime.time) and time1:
            if ':' in time1:
                time1 = datetime.datetime.strptime(time1, '%H:%M:%S')
            else:
                time1 = datetime.time(0,0)
        if not isinstance(time2, datetime.time) and time2:
            if ':' in time2:
                time2 = datetime.datetime.strptime(time2, '%H:%M:%S')
            else:
                time2 = datetime.time(0,0)

        if isinstance(p1, int):
            time1 = datetime.time(0,0) if time1 is not isinstance(time1, datetime.datetime) else time1
            date1 = datetime.datetime(0,0, 0) if date1 is None else date1
            date1 = date1.replace(hour=time1.hour, minute=time1.minute, second=time1.second)
            bill_A = bill_item(payer_info=personage(send1), payee_info=personage(recv1),
                               figure=p1, date=date1, type=0)
            ac_info.bill.append(bill_A)

        if isinstance(p2, int) :
            time1 = datetime.time(0,0) if time1 is not isinstance(time1, datetime.datetime) else time1
            time2 = time1 if time2 is None else time2
            date2 = date1 if date2 is None else date2
            date2 = date2.replace(hour=time2.hour, minute=time2.minute, second=time2.second)
            bill_A = bill_item(payer_info=personage(send2), payee_info=personage(recv2),
                               figure=p2, date=date2, type=1)
            ac_info.bill.append(bill_A)

    return
    for i in range(_sheet.max_row + 1):
        if i < 2:
            continue
        data_id = _sheet.cell(row=i, column=1).value
        ttype = type
        data_type = 1 if ttype else 0
        data_name = _sheet.cell(row=i, column=2).value
        data_price = _sheet.cell(row=i, column=6).value
        # data_tcount = sheet.cell(row=i, column=9).value
        data_tcount = 0

        st = _sheet.cell(row=i, column=7).value
        yg = _sheet.cell(row=i, column=8).value

        if data_name is None:
            continue
        data_count = 1
        # if '金龙鱼' in data_name or '维达' in data_name:
        #     data_count = 2

        t_commodity = commodity_info(id=data_id, count=data_count, total_count=data_tcount, type=data_type, name=data_name, price=data_price, geishitang=st, geiyuangong=yg)
        commodity_data[data_name + str(data_type)] = t_commodity

commodity_data={}

def get_path_formR(_data_p):
    for root, dirs, _ in os.walk(_data_p):
        if len(dirs) < 1:
            print('目录未找到人员文件夹')
            return None, None
        for ac_dir in dirs:
            for ac_p, _, files in os.walk(os.path.join(root, ac_dir)):
                xlsxDs = []
                for it_f in files:
                    if it_f.endswith('xlsx'):
                        xlsxDs.append(it_f)
                if len(xlsxDs) < 1:
                    print('{} 人员文件夹 没有xlsx 数据'.format(ac_p))
                for it_f in xlsxDs:
                    t_file = os.path.join(ac_p, it_f)
                    print('load file : {}'.format(t_file))
                    yield ac_dir, t_file


def load_data(_data_p):
    Collated_D = {}
    for ac_name, file in get_path_formR(_data_p):
        if ac_name is None:
            continue
        if ac_name not in Collated_D.keys():
            print('create ac : {}'.format(ac_name))
            Collated_D[ac_name] = personage_info(ac_name)

        # if ac_name == '行长':
        #     continue
        workbook = openpyxl.load_workbook(file, data_only=True)
        sheet = workbook['data']
        loadxlsx(sheet, 0, Collated_D[ac_name])

    print('done')


def foo(data_root_path, output_path):
    load_data(data_root_path)

    return None

    workbook = openpyxl.load_workbook(commodity_data_path, data_only=True)
    sheet = workbook['仓发1']
    loaddata(sheet, 0)
    sheet = workbook['代发1']
    loaddata(sheet, 1)



    workbook = openpyxl.load_workbook(src_data_path, data_only=True)
    sheet = workbook['Sheet1']

    wb = openpyxl.Workbook()
    # 选择默认的工作表
    # wsheet = wb.active
    ws1 = wb.create_sheet('代发')
    ws2 = wb.create_sheet('直发')
    ws3 = wb.create_sheet('other')

    # 给工作表重命名
    # wsheet.title = 'data'

    # 最大行 sheet.max_row # 最大列 sheet.max_column
    backdata = None

    totalAcountprice = {}
    # titleNameList = ['订单号', '姓名', '电话', '商家SKU编码', '单价', '数量', '应付', '实付', '收货人', '收货手机号', '省', '市', '区', '地址', '买家留言'] # back
    titleNameList = ['商品编号', '商品价格', '商品数量', '买家应付货款', '收货人名称', '联系电话', '联系手机', '省', '市', '区', '街道', '地址', '发票抬头', '买家留言', '收货人应付邮费'] # 0311
    for i in range(sheet.max_row + 1):
        if i < 2:
            if i == 1:
                ws1.append(titleNameList)
                ws2.append(titleNameList)
                # ws3.append(['商品编号', '商品名称', '销售总数', '下单价', '给食堂', '给员工', '总金额'])
                ws3.append(['商品编号', '商品名称', '销售总数', '下单价', '下单价总金额', '员工价', '员工价总金额', '食堂价', '食堂价总金额'])
            continue
        data_id = sheet.cell(row=i, column=1).value
        data_name = sheet.cell(row=i, column=2).value
        data_ipone = sheet.cell(row=i, column=3).value
        data_d1 = sheet.cell(row=i, column=4).value
        data_d2 = sheet.cell(row=i, column=5).value
        data_sname = sheet.cell(row=i, column=6).value
        data_sipone = sheet.cell(row=i, column=7).value
        data_saddr = sheet.cell(row=i, column=8).value
        data_saddr_all = sheet.cell(row=i, column=9).value
        data_total_amount = sheet.cell(row=i, column=10).value
        if data_id is None:
            continue

        data_sku_z_list = []
        data_count_z_list = []
        data_price_z_list = []
        data_sku_c_list = []
        data_count_c_list = []
        data_price_c_list = []
        for tdata in [data_d1, data_d2]:
            for it_di in tdata.split('┋'):
                if it_di == '(空)':
                    continue
                tcount = int(it_di.split('〖')[-1][0])
                resId, ratio = softID(it_di.split('〖')[0], commodity_data.keys(), 97)
                if resId:
                    it = resId
                    print(' {} : {} -- {}'.format(ratio, it, it_di))
                    if commodity_data[it].type == 1:
                        data_sku_z_list.append(commodity_data[it].id)
                        data_count_z_list.append(tcount)
                        data_price_z_list.append(commodity_data[it].price)
                    else:
                        data_sku_c_list.append(commodity_data[it].id)
                        data_count_c_list.append(tcount)
                        data_price_c_list.append(commodity_data[it].price)
                    # commodity_data[it].take()
                    commodity_data[it].addcount(tcount)
                else:
                    print('!!!!!!!!!!!!!!!!!! : {}', it_di)
                    if '金龙鱼' in it_di:
                        tkey1 = '金龙鱼乳玉皇妃稻香贡米5kg0'
                        tkey2 = '金龙鱼葵花籽油5L0'
                        if commodity_data[tkey1].type == 1:
                            data_sku_z_list.append(commodity_data[tkey1].id)
                            data_count_z_list.append(commodity_data[tkey1].count)
                            data_price_z_list.append(commodity_data[tkey1].price)
                            data_sku_z_list.append(commodity_data[tkey2].id)
                            data_count_z_list.append(commodity_data[tkey2].count)
                            data_price_z_list.append(commodity_data[tkey2].price)
                        else:
                            data_sku_c_list.append(commodity_data[tkey1].id)
                            data_count_c_list.append(commodity_data[tkey1].count)
                            data_price_c_list.append(commodity_data[tkey1].price)
                            data_sku_c_list.append(commodity_data[tkey2].id)
                            data_count_c_list.append(commodity_data[tkey2].count)
                            data_price_c_list.append(commodity_data[tkey2].price)

        type = 0
        yanzhengall = 0
        for t_sku_list, t_count_list, t_price_list in zip([data_sku_z_list, data_sku_c_list], [data_count_z_list, data_count_c_list], [data_price_z_list, data_price_c_list]):
            if len(t_sku_list) == 0:
                type += 1
                continue
            data_price_all = 0
            for tprice, tcount in zip(t_price_list, t_count_list):
                data_price_all = data_price_all + tcount * tprice
            data_sku = ' | '.join(t_sku_list)
            data_price = ' | '.join(list(map(str, t_price_list)))
            data_count = ' | '.join(list(map(str, t_count_list)))
            if '跳过' in data_saddr:
                data_sheng, data_shi, data_qv = '', '', ''
            else:
                data_sheng, data_shi, data_qv = data_saddr.split('-')

            # writeData = [data_id, data_name, data_ipone, data_sku, data_price, data_count, data_price_all, data_price_all,
            #              data_sname, data_sipone, data_sheng, data_shi, data_qv, data_saddr_all]

            # writeData = [data_sku, data_price, data_count, data_price_all, data_price_all,
            #              data_sname, '', data_sipone, data_sheng, data_shi, data_qv, '', data_saddr_all, '', '', '']

            writeData = [data_sku, data_price, data_count, data_price_all,
                         data_sname, '', data_sipone, data_sheng, data_shi, data_qv, '', data_saddr_all, '', '', '']

            print(writeData)

            yanzhengall = yanzhengall + data_price_all
            if type:
                ws1.append(
                    writeData)
            else:
                ws2.append(
                    writeData)
            type += 1

        totalAcountprice[data_sname] = yanzhengall
        print('all price : {}'.format(yanzhengall))

    print(totalAcountprice)
    print('total acount all price : {}'.format(sum(totalAcountprice.values())))

    for i in commodity_data.keys():
        if commodity_data[i].total_count > 0:
            writeData = [commodity_data[i].id, commodity_data[i].name, commodity_data[i].total_count,
                         commodity_data[i].price, commodity_data[i].total_count*commodity_data[i].price*commodity_data[i].count,
                         commodity_data[i].geiyuangong, commodity_data[i].total_count*commodity_data[i].geiyuangong*commodity_data[i].count,
                         commodity_data[i].geishitang, commodity_data[i].total_count*commodity_data[i].geishitang*commodity_data[i].count]
            print(writeData)
            ws3.append(
                writeData)
    # 保存 Excel 文件
    wb.save(resultname + str(time.time()) + '.xlsx')


if __name__ == '__main__':
    root_path = 'data/wjj'
    pj_path = 'pj1'
    output_path = 'output'

    foo(os.path.join(root_path, pj_path), os.path.join(root_path, output_path))
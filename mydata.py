import openpyxl
import mydata_data
import mydata_datab

metadata = {
    "SequenceNumber": 0,
    "QuestionType": 0,
    "Question": "",
    "KnowledgePoint": "",
    "Option": "",
    "Options": [],
    "StandardAnswer": "",
    "AnswerAnalysis": "",
    "tanswername": []
}

all_data = []
finddata = {}
for it in mydata_data.da['data']:
    tanswer = []
    tanswernames = []
    if len(it["Options"]) == 0:
        if it["StandardAnswer"] == 'B':
            tanswer.append('这道题错啦')
        elif it["StandardAnswer"] == 'A':
            tanswer.append('对着的')
        else:
            tanswer.append('填空题 答案是: \n \t{}'.format(it["StandardAnswer"]))
    else:
        for ii in it["Options"]:
            tanswer.append(ii['Key'])
            tanswer.append(ii['Value'])
            if ii['Key'] in it["StandardAnswer"]:
                tanswer.append(" True")
                tanswernames.append(ii['Value'])
            tanswer.append("    \n\t")

    # for ii in it["Option"]:
    #     tanswer.append(ii['Key'])
    #     tanswer.append(ii['Value'])
    #     # if ii['Key'] in it["StandardAnswer"]:
    #     #     tanswer.append(" True")
    #     tanswer.append("    \n\t")

    metadata["SequenceNumber"] = it["SequenceNumber"]
    metadata["QuestionType"] = it["QuestionType"]
    metadata["Question"] = it["Question"] + '\n'
    metadata["Option"] = ' '.join(tanswer)
    metadata["Options"] = it["Options"]
    metadata["StandardAnswer"] = it["StandardAnswer"]
    metadata["AnswerAnalysis"] = it["AnswerAnalysis"]
    metadata["KnowledgePoint"] = it["KnowledgePoint"]
    metadata["tanswername"] = tanswernames
    all_data.append(metadata.copy())
    finddata[metadata["Question"]] = metadata.copy()

metadatab = {
    "SequenceNumber": 0,
    "QuestionType": 0,
    "Question": "",
    "KnowledgePoint": "",
    "Option": "",
    "Options": [],
    "StandardAnswer": "",
    "AnswerAnalysis": "",
}
shijuan = []
result = []
for it in mydata_datab.db['data']['Questions']:
    sanswer = []
    tanswer = []
    for ii in it["Option"]:
        sanswer.append(ii['Key'])
        sanswer.append(ii['Value'])
        # if ii['Key'] in it["StandardAnswer"]:
        #     tanswer.append(" True")
        sanswer.append("    \n\t")

    metadatab["SequenceNumber"] = it["SequenceNumber"]
    metadatab["Question"] = it["Question"] + '\n'
    metadatab["Option"] = ' '.join(tanswer)
    if it["Question"]+'\n' in finddata.keys():
        tm = finddata[it["Question"]+'\n']["tanswername"]
        for ii in it["Option"]:
            if ii['Value'] in tm:
                tanswer.append(ii['Key'])
            # tanswer.append(ii['Value'])
            # if ii['Key'] in it["StandardAnswer"]:
            #     tanswer.append(" True")
            # tanswer.append("    \n\t")

        if it["QuestionType"] == 3:
            result.append({"id": metadatab["SequenceNumber"], "q": it["Question"] + '\n', "answer": finddata[it["Question"]+'\n']["StandardAnswer"]})
        elif it["QuestionType"] == 4:
            if finddata[it["Question"] + '\n']["StandardAnswer"] == "B":
                result.append({"id": metadatab["SequenceNumber"], "q": it["Question"] + '\n',
                               "answer": "错误的"})
            elif finddata[it["Question"] + '\n']["StandardAnswer"] == "A":
                result.append({"id": metadatab["SequenceNumber"], "q": it["Question"] + '\n',
                               "answer": "正确的"})
        else:
            result.append({"id": metadatab["SequenceNumber"], "q": it["Question"]+'\n', "answer": ' '.join(tanswer)})

    else:
        result.append({"id": metadatab["SequenceNumber"], "q": it["Question"] + '\n', "answer": "未知"})


print(result)


# 新建工作簿
wb = openpyxl.Workbook()
# 选择默认的工作表
sheet = wb.active
# 给工作表重命名
sheet.title = 'data'

# 写入多行数据
for row in all_data:
  sheet.append([row["SequenceNumber"],row["Question"],row["Option"],row["StandardAnswer"],row["AnswerAnalysis"],row["KnowledgePoint"]])

# count = 0
# for row in result:
#
#     if count % 5 == 0:
#         sheet.append(["", "", ""])
#     sheet.append([row["id"], row["answer"], row["q"]])
#     count += 1

# 往某个单元格子写入数据
# sheet['A1'] = 'superman'

# 保存 Excel 文件
wb.save('a.xlsx')


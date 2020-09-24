import pymongo

from util import _to_chinese4, init_excel, init_word, pretty_text

myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["wenshu"]
zbwenshu = mydb["zbwenshu"]
zbwenshu_relationship = mydb["zbwenshu_relationship"]

## 第7、10、14、30、37、38、39、40、41、42、46、47、48、50条
# (1)第7、14、30、38、39、40、46、50条
demand1 = [7, 14, 30, 38, 39, 40, 46, 50]
demand1_cn = ["第" + _to_chinese4(i) + "条" for i in demand1]
# (2)第10、48条
demand2 = [10, 48]
demand2_cn = ["第" + _to_chinese4(i) + "条" for i in demand2]
# (3)第41、42、47条
demand3 = [41, 42, 47]
demand3_cn = ["第" + _to_chinese4(i) + "条" for i in demand3]


# 打印文书到word
def print_word(row, index, doc):
    content = ""
    content += "===========================【案件" + str(index) + "】==========================\n"
    content += "【案件id】" + str(row["_id"]) + "\n"
    content += "【案号】" + row["case_info"]["案号"] + "\n"
    content += "【适用法规】" + "\n"
    for legal in row["legal_base"]:
        content += legal["法规名称"] + "\n"
        for i in legal["Items"]:
            content += i["法条名称"] + "\n"

    content += "【案件原文】" + "\n"

    doc.add_paragraph(content)

    pretty_text(doc, row["doc"])
    doc.add_paragraph("===========================【案件" + str(index) + "】==========================")

# 打印excel
def print_excel(row, index, worksheet):
    list = []
    list.append(index)
    # 案件id
    list.append(row["_id"])
    # 案件号
    list.append(row["case_info"]["案号"])
    # 目标法条
    list.append(row["legal_base"])
    # 裁判日期
    list.append(row["relate_info"]["裁判日期"])
    # 案由
    list.append(row["relate_info"]["案由"])
    # 审理程序
    list.append(row["relate_info"]["审理程序"])
    # 省份
    list.append(row["case_info"]["法院省份"])
    # 法院名字
    list.append(row["case_info"]["法院名称"])

    for ii, item in enumerate(list):
        worksheet.write(index, ii, label=str(item))

# 查询数据
def query_data(item, limit):
    # 查询
    table = zbwenshu.find({
        "legal_base":
            {
                "$elemMatch": {
                    "法规名称": "《中华人民共和国劳动合同法》",
                    "Items.法条名称": {"$regex": "^" + item}
                }
            }})

    # 写word
    doc = init_word()

    # 写excel
    workbook, worksheet = init_excel()

    index = 1
    for row in table:
        print("处理第" + str(index) + "个")

        try:
            print_word(row, index, doc)
            print_excel(row, index, worksheet)
        except Exception as e:
            print(e)
            continue

        index += 1

        if index > limit:
            break

    # 保存
    doc.save("result/附件2：" + item + "（" + str(limit) + "份）.docx")
    workbook.save("result/附件1：" + item + "（" + str(limit) + "份）.xls")

if __name__ == '__main__':
    for i in demand1_cn:
        query_data(i, 2000)

    for i in demand2_cn:
        query_data(i, 2500)

    for i in demand3_cn:
        query_data(i, 1000)


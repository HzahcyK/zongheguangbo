import pymysql
import xlsxwriter

start_time = input("请输入查询起始时间：")
end_time = input("请输入查询截止时间：")
conn = pymysql.connect(host="10.10.10.240", port=5432, user="root", password="tF!e5UN?iGMRkB7Z80Ln#O@uCsP^mS", db="dj_analytics", charset="utf8")
cursor = conn.cursor()
cursor.execute("select pub_date, title, case is_original when 0 then '转载' when 1 then '原创' else 'null' end original, category, editor, read_count, department from views_article_department_channel where department='综合广播' and (pub_date between '%s' and '%s');" % (start_time, end_time))
result = cursor.fetchall()
result = list(result)
datas = []
for item in result:
    item = {'发布时间': item[0], '标题': item[1], '原创/转载': item[2], '栏目': item[3], '编辑': item[4], '阅读量': item[5], '部门': item[6]}
    datas.append(item)
print(datas)

workbook = xlsxwriter.Workbook(zhgb.xlsx)
worksheet = workbook.add_worksheet()
for i in range(1, len(datas)):
    for data in datas:
        row = 'A' + str(i)
        worksheet.write_row(row, data)
        break
workbook.close()


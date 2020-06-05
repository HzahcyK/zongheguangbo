import pymysql
import xlsxwriter

start_time = input("请输入查询起始时间：")
end_time = input("请输入查询截止时间：")
file_name = input("文件名：")
conn = pymysql.connect(host="10.10.10.240", port=5432, user="root", password="tF!e5UN?iGMRkB7Z80Ln#O@uCsP^mS", db="dj_analytics", charset="utf8")
cursor = conn.cursor()
cursor.execute("select pub_date, title, case is_original when 0 then '转载' when 1 then '原创' else 'null' end original, category, editor, read_count, department from views_article_department_channel where department='综合广播' and (pub_date between '%s' and '%s');" % (start_time, end_time))
result = cursor.fetchall()
result = list(result)
datas = []
for item in result:
    item = [item[0], item[1],  item[2], item[3],  item[4], item[5], item[6]]
    datas.append(item)
print(datas)

workbook = xlsxwriter.Workbook(file_name + '.xlsx')
worksheet = workbook.add_worksheet()
date_format = workbook.add_format({'num_format':'yyyy/mm/dd hh:mm'})
worksheet.set_column('A:A', date_format)
for i in range(1, len(datas)):
        row = 'A' + str(i)
        worksheet.write_row(row, data[i-1])
        break
workbook.close()


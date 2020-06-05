import pymysql
import xlsxwriter

year = input("请输入查询年份：")
start_n = input("请输入查询起始月份：")
end_n = input("请输入查询截止月份：")
start_time = '2020-' + start_n + '-01 00:00:00'
end_time = '2020-' + end_n + '-01 00:00:00'
file_name = year + '-' + start_n + '~' + end_n
conn = pymysql.connect(host="10.10.10.240", port=5432, user="root", password="tF!e5UN?iGMRkB7Z80Ln#O@uCsP^mS", db="dj_analytics", charset="utf8")
cursor = conn.cursor()
cursor.execute("select (@i:=@i+1) id, pub_date, title, case is_original when 0 then '转载' when 1 then '原创' else 'null' end original, category, editor, read_count, department from views_article_department_channel, (select @i:=0) r where department='综合广播' and (pub_date between '%s' and '%s');" % (start_time, end_time))
result = cursor.fetchall()
result = list(result)
datas = []
for item in result:
    item = [item[0], str(item[1]),  item[2], item[3],  item[4], item[5], item[6], item[7]]
    datas.append(item)
print(datas)

workbook = xlsxwriter.Workbook(file_name + '.xlsx')
worksheet = workbook.add_worksheet()
title_style = workbook.add_format({
    'bold': 1,             #字体加粗
    'fg_color': 'yellow',   #单元格背景颜色
    'align': 'center',     #对齐方式
    'valign': 'vcenter',   #字体对齐方式
 })
id_style = workbook.add_format({
    'bold': 1,
    'align': 'center',
})
content_style = workbook.add_format({
    'align': 'center',
})
worksheet.set_column("A:A", None, id_style)
worksheet.set_column("B:B", 21)
worksheet.set_column("C:C", 70)
title = [u'id', u'发布时间', u'标题', u'是否原创', u'栏目', u'编辑', u'阅读量', u'部门']
worksheet.write_row('A1', title, title_style)
for i in range(2, len(datas)):
        row = 'A' + str(i)
        worksheet.write_row(row, datas[i-2], content_style)
workbook.close()


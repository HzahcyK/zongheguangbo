import pymysql
import xlsxwriter

year = input("请输入查询年份：")
start_n = input("请输入查询起始月份：")
end_n = input("请输入查询截止月份：")
xj_name = input("请输入县级名称")
start_time = year + '-' + start_n + '-01 00:00:00'
end_time = year + '-' + end_n + '-01 00:00:00'
file_name = year + '-' + start_n + '~' + end_n
conn = pymysql.connect(host="10.10.10.240", port=5432, user="root", password="tF!e5UN?iGMRkB7Z80Ln#O@uCsP^mS", db="dj_analytics", charset="utf8")
cursor = conn.cursor()
cursor.execute("select (@i:=@i+1) id, pub_date, title, case is_original when 0 then '转载' when 1 then '原创' else 'null' end original, read_count, id, source from views_article_region, (select @i:=0) r where region='%s' and (pub_date between '%s' and '%s');" % (xj_name, start_time, end_time))
result = cursor.fetchall()
result = list(result)
datas = []
for item in result:
    item = [item[0], str(item[1]),  item[2], item[3],  item[4], 'https://movement.gzstv.com/news/detail/' + str(item[5]), item[6]]
    datas.append(item)
print(datas)

workbook = xlsxwriter.Workbook(file_name + '.xlsx')
worksheet = workbook.add_worksheet()
colname_style = workbook.add_format({
    'bold': 1,             #字体加粗
    'fg_color': 'yellow',   #单元格背景颜色
    'align': 'center',     #对齐方式
    'valign': 'vcenter',   #字体对齐方式
 })
id_style = workbook.add_format({
    'bold': 1,
    'align': 'center',
})
# content_style = workbook.add_format({
#     'align': 'center',
# })

worksheet.set_column("A:A", None, id_style)
worksheet.set_column("B:B", 21)
worksheet.set_column("C:C", 70)
title = [u'序号', u'发布时间', u'标题', u'原创', u'阅读量', u'链接', u'来源']
worksheet.write_row('A1', title, colname_style)
for i in range(2, len(datas)):
        row = 'A' + str(i)
        worksheet.write_row(row, datas[i-2])
workbook.close()


import xlrd


class ReadExcel:
    def readexcel(self, url):
        data = xlrd.open_workbook(url)  # 打开xls文件
        table = data.sheets()[0]  # 打开第一张表
        nrows = table.nrows  # 获取表的行数

        htmlhead = '''<!DOCTYPE html>
                        <html lang="en">
                            <head>
                                <meta charset="UTF-8">
                                <title>Title</title>
                            </head>
                            <body>'''

        htmltable = '<table border="1">'

        htmltable += '<tr>'

        for row in range(nrows):
            htmltable += '<tr>'

            for e in table.row_values(row):
                htmltable += '<td>' + str(e) + '</td>'

            htmltable += '</tr>'

        htmltable += '</table>'
        htmltail = '</body></html>'

        html = htmlhead + htmltable + htmltail

        print(html)

        return html

# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import xlrd
import pymssql
import datetime
import os

# 建立本地连接
conn = pymssql.connect(server='127.0.0.1', user='sa', password='123', database='yblr_xz')  # 地址，用户名，密码，数据库
path = "C:/Users/Administrator/Downloads/献血数据 (1)/献血数据/"
cursor = conn.cursor()


def readfile(file):
    bk = xlrd.open_workbook(file)
    xlrd.Book.encoding = "utf-8"
    sh = bk.sheets()[0]
    start_time = datetime.datetime.now()
    sql3 = ''
    for i in range(2, sh.nrows - 3, 2):
        a = []
        sql = '('
        for j in range(1, sh.ncols):
            if j == 1:  # 原文件第一列是空行
                sql += "'" + str(sh.cell(i + 1, j).value) + "'" + ','
            elif j == sh.ncols - 3:  # 数字为空处理
                if str(sh.cell(i, j).value) == '':
                    sql += "'" + '0' + "'" + ','
                else:
                    sql += "'" + str(sh.cell(i, j).value) + "'" + ','
            elif j == sh.ncols - 5:  # 时间处理
                if str(sh.cell(i, j).value) == '':
                    sql += "''" + ','
                else:
                    # date_arr = str(sh.cell(i, j).value).split('.')
                    date = xlrd.xldate.xldate_as_datetime(sh.cell(i, j).value, 0)
                    sql += "'" + str(date.date()) + "'" + ','
            elif j == sh.ncols - 8:  # 时间处理
                if str(sh.cell(i, j).value) == '':
                    sql += "''" + ','
                else:
                    # date_arr = str(sh.cell(i, j).value).split('.')
                    date = xlrd.xldate.xldate_as_datetime(sh.cell(i, j).value, 0)
                    sql += "'" + str(date.date()) + "'" + ','
            else:  # 其他情况处理，单引号字符处理
                sql += "'" + str(sh.cell(i, j).value).replace("'", "''") + "'" + ','
        sql2 = sql.strip(',')
        sql3 += sql2.strip() + '),'
        if i % 1000 == 0:
            sql3 = sql3.strip(',')
            sql1 = 'insert into [dbo].[Table_1]([献血码],[献血者姓名],[血型],[性别],[年龄],[献血类型],[证件号码],[地址],[所属单位],[体检日期],[体检结果],[初检结果],[采血日期],[采血形式],[采血量],[复检结果],[献血证号]) VALUES %s' % sql3
            print(file)
            print(sql1)
            # 去掉sql语句中的特殊字符
            cursor.execute(sql1.replace('\u0000', '').replace('\x00', ''))
            sql = ""
            sql3 = ""
    sql3 = sql3.strip(',')
    sql1 = 'insert into [dbo].[Table_1]([献血码],[献血者姓名],[血型],[性别],[年龄],[献血类型],[证件号码],[地址],[所属单位],[体检日期],[体检结果],[初检结果],[采血日期],[采血形式],[采血量],[复检结果],[献血证号]) VALUES  %s' % sql3
    print(file)
    print(sql1)
    cursor.execute(sql1.replace('\u0000', '').replace('\x00', ''))
    # conn.commit()
    conn.close
    end_time = datetime.datetime.now()
    speed = end_time - start_time
    return speed


pathDir = os.listdir(path)
for allDir in pathDir:
    child = os.path.join('%s%s' % (path, allDir))
    readfile(child)
    # print(child)
conn.commit()

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/

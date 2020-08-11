# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import xlrd
import pymssql
import datetime
import os

# 建立本地连接
conn = pymssql.connect(server='127.0.0.1', user='sa', password='1234', database='zmzb2')  # 地址，用户名，密码，数据库
path = 'd:'
cursor = conn.cursor()


def readFile(file):
    bk = xlrd.open_workbook(file)
    sh = bk.sheets()[0]
    start_time = datetime.datetime.now()
    sql3 = ''
    for i in range(1, sh.nrows):
        a = []
        sql = '('
        for j in range(sh.ncols):
            sql += "'" + str(sh.cell(i, j).value) + "'" + ','
        sql2 = sql.strip(',')
        sql3 += sql2.strip() + '),'
        if i % 1000 == 0:
            sql3 = sql3.strip(',')
            sql1 = 'insert into '
            cursor.execute(sql1)
            sql = ""
            sql3 = ""
    sql3 = sql3.strip(',')
    sql1 = 'insert into '
    cursor.execute(sql1)
    conn.commit()
    end_time = datetime.datetime.now()
    speed = end_time - start_time
    return speed


pathDir = os.listdir(path)
for allDir in pathDir:
    child = os.path.join('%s%s' % (path, allDir))
    print(child)

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/

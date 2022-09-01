import xlrd
from pymongo import MongoClient


data = xlrd.open_workbook('aac13.xls')
table = data.sheets()[0]

def names(table):
    ans = []
    namelist = table.col_values(4)
    print(namelist[0])
    for i in namelist[2:]:
        if i.isalnum():
            ans.append(i)
    return ans


def count(table,i):
    count = 0
    check = table.col_values(8)
    # for i in range(len(check)):
    #     if len(check[i])==7:
    #         count+=1
    # print(count)
    if len(check[i])==7:
        return True

#count(table)
print(len(table.col_values(4)))
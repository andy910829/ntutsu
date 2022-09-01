import xlrd
import pymongo
from pymongo import MongoClient


data = xlrd.open_workbook('aac13.xls')


table = data.sheets()[0]



def turnintodict(table,i):
    results = {}
    #print(names(table)[i],students_id(table)[i],pay(table)[i])
    results["id"] = i+1
    results["name"] = table.row_values(i)[4]
    results["student_id"] = table.row_values(i)[6]
    results["payment"] = table.row_values(i)[9]
    #print(results)
    return results

def check(table,i):
    check = table.col_values(8)
    if len(check[i])==7:
        return True


def main(table):
    cluster = MongoClient("mongodb://localhost:27017")
    db = cluster["ntutsu"]
    collection = db["freshman"]
    count = 0
    for i in range(len(table.col_values(4))):
        if check(table,i):
            print(turnintodict(table,i))
            post = turnintodict(table,i)
            collection.insert_one(post)
    #         count+=1
    # print(count)

main(table)


# len(table.col_values(4))
#-----------------------------------------------------------------------------------------------------------------------------------------------

# def names(table):
#     ans = []
#     namelist = table.col_values(4)
#     print(namelist[0])
#     for i in namelist[2:]:
#         if i.isalnum():
#             ans.append(i)
#     return ans

# def students_id(table):
#     ans = []
#     student_id = table.col_values(6)
#     for i in student_id[2:]:
#         if i.isalnum():
#             ans.append(i)
#     return ans
# def pay(table):
#     ans = []
#     payment = table.col_values(9)
#     for i in payment:
#         if isinstance(i,float):
#             ans.append(i)
#     return ans


# def turnintodict(table,i):
#     results = {}
#     #print(names(table)[i],students_id(table)[i],pay(table)[i])
#     results["id"] = i
#     results["name"] = names(table)[i]
#     results["student_id"] = students_id(table)[i]
#     results["payment"] = pay(table)[i]
#     #print(results)
#     return results

# def main(table):
#     cluster = MongoClient("mongodb://localhost:27017")
#     db = cluster["ntutsu"]
#     collection = db["freshman"]
#     for i in range(len(names(table))):
#         print(i)
#         post = turnintodict(table,i)
#         collection.insert_one(post)



# for i in range(len(names(table))):
#         post = turnintodict(table,i)
#         print(post)
# main(table)
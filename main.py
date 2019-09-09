import xlwt
import xlrd

fileName = 'xinfo.xlsx'
sheetName = '信息表'
data = {
    "1": ["张三", 150, 120, 100],
    "2": ["李四", 90, 99, 95],
    "3": ["王五", 60, 66, 68]
}


# 新建excel表
def main():
    work_book, sheet = newFile(fileName, sheetName)
    write(work_book, sheet)
    read(fileName, sheetName)


def newFile(fileName, sheetName):
    work_book = xlwt.Workbook(encoding='utf-8')
    sheet = work_book.add_sheet(sheetName)
    work_book.save(fileName)
    return work_book, sheet


def read(fileName, sheetName):
    work_book = xlrd.open_workbook(fileName)
    sheet = work_book.sheet_by_name(sheetName)
    rows = sheet.nrows
    cols = sheet.ncols

    for i in range(rows):
        for j in range(cols):
            cell = sheet.cell_value(i, j)  # 某一单元格数据
            print(cell)


def write(work_book, sheet):
    ldata = []
    num = [a for a in data]
    # for循环指定取出key值存入num中
    num.sort()
    # 字典数据取出后无需，需要先排序
    for x in num:  # for循环将data字典中的键和值分批的保存在ldata中
        t = [int(x)]
        for a in data[x]:
            t.append(a)
        ldata.append(t)
    for i, p in enumerate(ldata):
        # 将数据写入文件,i是enumerate()函数返回的序号数
        for j, q in enumerate(p):
            sheet.write(i, j, q)
    work_book.save(fileName)


if __name__ == '__main__':
    main()

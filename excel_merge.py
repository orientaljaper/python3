import xlwt
import xlrd
import xlsxwriter

"""
从文件1中的"新","旧" sheet捞取每行数据，交叉写入新文件。
"""


if __name__ == "__main__":
    # 打开Excel原文件
    origin_path = "文件1.xls"
    origin_book = xlrd.open_workbook(origin_path)
    write_book = xlsxwriter.Workbook("文件2.xlsx")
    write_sheet = write_book.add_worksheet("merge")

    origin_sheet_new = origin_book.sheet_by_name("新")
    origin_sheet_old = origin_book.sheet_by_name("老")

    for new_row in range(origin_sheet_new.nrows):
        # write_sheet.write_row(2 * new_row, 0, str(origin_sheet_new.row(new_row)))
        write_sheet.write_row("A" + str(new_row * 2), origin_sheet_new.row_values(new_row))

    for old_row in range(origin_sheet_old.nrows):
        write_sheet.write_row("A" + str(old_row * 2 + 1), origin_sheet_old.row_values(old_row))

        #write_sheet.write_row(2 * old_row + 1,0, str(origin_sheet_new.row(new_row)))

    write_book.close()


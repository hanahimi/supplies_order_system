#-*-coding:UTF-8-*-
'''
Created on 2018年2月10日-下午3:04:23
author: Gary-W
'''

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

import os

class OrderItem:
    def __init__(self):
        self.item_id = None         # 料号
        self.order_id = None        # 订单号
        self.income_num = 0        # 到货数量
    
    def __str__(self):
        return "ID:%s ORD:%s N:%d" % (self.item_id, self.order_id, self.income_num)
    
    
class OrderDataSheet:
    """ 读取到货情况XLSX表格，获取其中数据
    """
    def __init__(self,xlsx_path, sheetname="325"):  #zhijia
        """ 读入已有的xls表，并制定其中的sheet
        """
        self.xlsx_path = xlsx_path
        self.workbook = load_workbook(xlsx_path)
        self.main_sheet = self.workbook.get_sheet_by_name(sheetname)

        # 获取表格参数
        self.header_rows_offset = 1    # 标题所占行数
        self.sample_num = self._get_sample_num()
        self.item_list = []
        self.get_items()
        
    def _get_sample_num(self):
        """ 统计数表的有效行数
        """
        rows = len(self.main_sheet.rows)
        for i in range(rows,0,-1):
            if self.main_sheet["B"+str(i)].value != None:
                sample_num = i - 1
                break
        return sample_num

    def get_items(self):
        """ 获得表中所有的数据
        """
        for row_id in range(self.sample_num):
            orditem = OrderItem()
            orditem.item_id = str(self.main_sheet["E"+str(row_id+2)].value)
            orditem.order_id = str(self.main_sheet["I"+str(row_id+2)].value)
            orditem.income_num = int(self.main_sheet["G"+str(row_id+2)].value)
            self.item_list.append(orditem)


class TrackItem:
    def __init__(self):
        self.item_id = None         # 料号
        self.order_id = None        # 订单号
        self.target_num = 0         # 目标总数量
        self.income_num = 0         # 入库总数量
        self.rest_num = 0           # 未交数量
        self.mat_row = 0            # 该货物所在的行

    def __str__(self):
        return "%s %s %d %d" % (self.item_id, self.order_id, self.target_num, self.rest_num)
    
class TrackDataSheet:
    """ 写入到货情况XLSX表格，更新其中数据
    """
    def __init__(self,xlsx_path, sheetname="Sheet1"):
        """ 读入已有的xls表，并制定其中的sheet
        """
        self.xlsx_path = xlsx_path
        self.workbook = load_workbook(xlsx_path)
        self.main_sheet = self.workbook.get_sheet_by_name(sheetname)

        # 获取表格参数
        self.header_rows_offset = 2    # 标题所占行数
        self.sample_num = self._get_sample_num()
        self.item_id_col = 1
        self.ord_id_col = 2

        self.item_table = {}
        self._get_item_table()

    def _get_sample_num(self):
        """ 统计数表的有效行数
        """
        rows = len(self.main_sheet.rows)
        for i in range(rows,0,-1):
            if self.main_sheet["B"+str(i)].value != None:
                sample_num = i - 2
                break
        return sample_num

    def _get_item_table(self):
        latest_item_id = None
        for row_id in range(self.sample_num):
            item_id = self.main_sheet["A"+str(row_id+3)].value
            order_id = self.main_sheet["B"+str(row_id+3)].value
            
            if item_id is not None:
                self.item_table[item_id] = {}
                latest_item_id = item_id

            latest_item_id = str(latest_item_id)
            order_id = str(order_id)
            
            # 填写 料号，订单号，数量，未交数量； 对应的表行id
            track_item = TrackItem()
            track_item.item_id = latest_item_id
            track_item.order_id = order_id
            track_item.mat_row = row_id+3
            
            track_item.target_num = int(self.main_sheet["F"+str(row_id+3)].value)
            track_item.rest_num = int(self.main_sheet["G"+str(row_id+3)].value)
            track_item.income_num = track_item.target_num - track_item.rest_num
            self.item_table[latest_item_id][order_id] = track_item
    
    
    def load(self, insert_col, income_items):
        order_col = insert_col
        income_col = insert_col + 1
        rest_col = 7    # 未交数量更新
        err_msg_list = []
        align = Alignment(horizontal='right', vertical='center')
        for i, income_item in enumerate(income_items):
            if income_item.item_id in self.item_table:
                item_id = income_item.item_id
                order_id = income_item.order_id
                income_num = income_item.income_num
                if order_id not in self.item_table[item_id]:
                    err_msg = "ItemID: %s ORD: %s not found: %s" % (item_id, order_id, income_item)
                    err_msg_list.append(err_msg)
                else:
                    # 完整匹配所有信息
                    insert_row = self.item_table[item_id][order_id].mat_row
                    
                    # 将订单号和入库数量写入对应的单元格
                    order_cell = self.main_sheet.cell(row = insert_row, column = order_col)
                    income_cell = self.main_sheet.cell(row = insert_row, column = income_col)
                    rest_cell = self.main_sheet.cell(row = insert_row, column = rest_col)
                    
                    # 更新未交数量
                    rest_cell.value = rest_cell.value - income_num
                    # 更新当前已交数量
                    if income_cell.value == None:
                        income_cell.value = income_num
                    else:
                        income_cell.value += income_num
                    # 更新订单号
                    order_cell.value = int(order_id)
                    
                    order_cell.alignment = align
                    income_cell.alignment = align
            else:
                err_msg = "ItemID: %s not found: %s" % (income_item.item_id, income_item)
                err_msg_list.append(err_msg)
        print "update:",self.xlsx_path
        self.workbook.save(self.xlsx_path)

        err_log_path = self.xlsx_path[:-5] + "_err.txt"
        with open(err_log_path, "w") as f:
            for err_msg in err_msg_list:
                f.write(err_msg+"\n")



def main():
    
#     input_xlsx_path = r"D:\到料明细.xlsx"
#     update_xlsx_path = r"D:\物料订单执行跟踪表.xlsx"
    input_xlsx_path = raw_input("料明细 路径: ")
    sheetname = raw_input("到料明细表名: ")
    update_xlsx_path = raw_input("物料订单执行跟踪表 路径: ")
    insert_col = raw_input("输入跟踪表更新的列: ")

    if not os.path.exists(input_xlsx_path):
        print "找不到",input_xlsx_path
        raw_input("按回车退出")
        return
    
    if not os.path.exists(update_xlsx_path):
        print "找不到",update_xlsx_path
        raw_input("按回车退出")
        return

    insert_col = int(insert_col)
    if insert_col < 10:
        print "无法对应列", insert_col
        raw_input("按回车退出")
        return

    ds_input = OrderDataSheet(input_xlsx_path, sheetname)

    ds_update = TrackDataSheet(update_xlsx_path)
    ds_update.load(insert_col, ds_input.item_list)
    raw_input("更新成功 按回车退出")
 



if __name__=="__main__":
    pass
    main()

import os
import xlrd

file_path = os.path.join(os.path.dirname(__file__),'../data/test_case.xlsx')
sheet_name = 'Sheet1'

class excelUtil():
    def __init__(self,file_path,sheet_name):
        self.file_path = file_path
        self.sheet_name =sheet_name
        self.wb = xlrd.open_workbook(self.file_path)
        self.sheet = self.ws()


    def ws(self):
        ws = self.wb.sheet_by_name(self.sheet_name)
        return ws

    def get_rows_num(self):
        return self.sheet.nrows

    def get_cols_num(self):
        return self.sheet.ncols

    def __get_cell_value(self, row_index, col_index):
        return self.sheet.cell_value(row_index, col_index)

    def __get_merge_cells(self):
        """
        获取合并单元格
        :return:
        """
        return self.sheet.merged_cells

    def get_all_type_cell_value(self, row_index, col_index):
        """
        获取合并单元格/非合并单元格的内容
        :param row_index:
        :param col_index:
        :return:
        """
        for (min_row, max_row, min_col, max_col) in self.__get_merge_cells():
            if min_row <= row_index < max_row and min_col <= col_index < max_col:
                cell_value = self.__get_cell_value(min_row, min_col)
                break
            else:
                cell_value = self.__get_cell_value(row_index, col_index)
        return cell_value

    def get_value_by_dict(self):
        data_list = []
        for r in range(1, self.get_rows_num()):  # 控制行
            dict_data = {}
            for l in range(0, self.get_cols_num()):  # 控制列
                dict_data[self.get_all_type_cell_value(0, l)] = self.get_all_type_cell_value(r, l)
            data_list.append(dict_data)
        return data_list



if __name__ == '__main__':
    for i in range(9):
        print(excelUtil(file_path, sheet_name).get_all_type_cell_value(i,0))
    print(excelUtil(file_path, sheet_name).get_value_by_dict())


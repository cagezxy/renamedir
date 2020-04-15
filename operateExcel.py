# coding=utf-8
import os
import xlrd


class ExcelUtil():
	def __init__(self, excelpath, sheetname="Sheet1"):
		self.data = xlrd.open_workbook(excelpath)
		self.table = self.data.sheet_by_name(sheetname)
		# 获取第二行作为key值
		self.keys = self.table.row_values(1)
		# 获取总行数
		self.rowNum = self.table.nrows
		# 获取总列数
		self.colNum = self.table.ncols

	# def dict_data(self):
	# 	if self.rowNum <= 3:
	# 		print("总行数小于3")
	# 	else:
	# 		r = []
	# 		j = 3
	# 		for i in list(range(self.rowNum - 3)):
	# 			s = {}
	# 			# 从第三行取对应values值
	# 			s['rowNum'] = i + 3
	# 			values = self.table.row_values(j)
	# 			for x in list(range(self.colNum)):
	# 				s[self.keys[x]] = values[x]
	# 			r.append(s)
	# 			j += 1
	# 		return r

	def dict_data(self):
		if self.rowNum <= 3:
			print("总行数小于3")
		else:
			r = []
			s = {}
			j = 3
			for i in list(range(self.rowNum - 3)):
				values = self.table.row_values(j)
				r.append(values)
				# 由于excel列固定，找到第4行的编号，第5行的名称，放入字典中
				file_id = int(values[4])
				file_name = values[5]
				s[file_name] = file_id
				j += 1
			return s


if __name__ == "__main__":
	print(os.getcwd())
	filepath = "D:\\示例\\大区汇总表（20年）.xlsx"
	sheetName = "Sheet1"
	data = ExcelUtil(filepath, sheetName)
	print(data.dict_data())
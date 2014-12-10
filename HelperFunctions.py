import xlrd
from xlrd.sheet import ctype_text 
import re

#Global Variables ----> TODO : Take as user input later ?? 

IFC_FILE_PATH = "slab_segmentation_V4_3floors_SEEBIM.ifc"
EXCEL_FILE_PATH = "loadTables.xlsx"
# EXCEL_SHEET_1 = "slab_segmentation_V4_3floors_SE"

################################################
################ HELPER CLASS ##############
################################################


class Helping:


################################################
################ EXCEL CLASS METHODS ##############
################################################


	def getExcelCellAtRowCol(self, sheetName, rows, cols):
		book = xlrd.open_workbook(EXCEL_FILE_PATH)
		for name in book.sheet_names():
			if name == sheetName:
				sheet = book.sheet_by_name(name)
				if rows<= sheet.nrows and cols<= sheet.ncols:
					return sheet.cell(rows, cols).value
				else:
					return ""


	



	# def openExcelFile(self):
	# 	book = xlrd.open_workbook(EXCEL_FILE_PATH)
	# 	for name in book.sheet_names():
	# 		print name
	# 		if name == EXCEL_SHEET_1:
	# 			sheet = book.sheet_by_name(name)

	# 			# Attempt to find a matching row (search the first column for 'john')
	# 			rowIndex = -1
				
	# 			row=0

	# 			for row in range(0, sheet.nrows):
	# 				for col in range(0, sheet.ncols):
	# 					if sheet.cell(row, col).value=="24 x 24 - PC":
	# 						print "Hello :: ",sheet.cell(row, col+1).valuec
	# 						break

	# 			book.unload_sheet(name) 


################################################
################ OTHER HELPER METHODS ##############
################################################


	def openFile(self, filePath):
		return open(filePath, 'r')

	def findStringBetweeenStrings(self, originalString, startStr, endStr):
		try:
			startIndex = originalString.index(startStr) + len(startStr)
			endIndex = originalString.index( endStr, startIndex )
			return originalString[startIndex:endIndex]
		except ValueError:
			return ""

	def findLargestIndexOfIFCFile(self):
		ifcFile = self.openFile(IFC_FILE_PATH)
		# count = 0
		for line in ifcFile:
			if ("#" in line) and (line[0]=="#"): 
				lastIndex = int(self.findStringBetweeenStrings(line, "#", "="))

		ifcFile.close()
		return lastIndex

	def parseFile(self):
		
		lastIndex = self.findLargestIndexOfIFCFile() + 10
		ifcFile = self.openFile(IFC_FILE_PATH)
		ifcList = [0]*lastIndex;
		count = 0;

		for line in ifcFile:
			if ("#" in line) and (line[0]=="#"): 
				index =  int(self.findStringBetweeenStrings(line, "#", "="))
				value = line  #self.findStringBetweeenStrings(line, "=", ";")
				ifcList[index] = value
				count = count + 1 
		ifcFile.close()
		
		return ifcList


############ Methods to Parse Each Line According to Type ############

	def findIFCSlabWidth(self, ifcList):
		ifcSlabWidth = []
		ifcSlabIndexes = []
		count = -1 
		for each in ifcList:
			count += 1 
			each = str(each)
			if "IFCSLAB" in each:
				ifcSlabWidth.append(float(each.split(",")[4].split("'")[1].split(" ")[2]))
				ifcSlabIndexes.append(count)
		return [ifcSlabWidth, ifcSlabIndexes]



	def findMaxLoad(self, sheetName, loadForProject, slabWidth):
		# print "Iteration for Load: ", loadForProject, " || Slab: ", slabWidth
		checkrow = 0
		checkcol=0
		book = xlrd.open_workbook(EXCEL_FILE_PATH)
		for name in book.sheet_names():
			if name == sheetName:
				sheet = book.sheet_by_name(name)
				for checkrow in range(1,10):
					if sheet.cell(checkrow,0).value!="":
						if loadForProject == int(sheet.cell(checkrow,0).value):
							# obtained checkrow containing the load value
							break

				for checkcol in range(1, (sheet.ncols-3)):
					if sheet.cell(0, checkcol).value != "":
						if round(slabWidth, -1) == int(sheet.cell(0, checkcol).value):
							# obtained checkcol containing the slab value
							break

				tempMaxLoad = 0
				maxOthers = []

				#check to see if you've moved into another span
				for col in xrange(checkcol,checkcol+7):
					if col > sheet.ncols:
						pass
					else:
						tempArr = re.findall('\d+',sheet.cell(checkrow, col).value)
						count = -1
						for each in tempArr:
							count += 1
							if int(each) > int(tempMaxLoad):
								# continue from here.... !! 
								if "16+2" in str(sheet.cell((checkrow+1), col+1).value):
									if tempArr[count] not in maxOthers:
										maxOthers.append(int(tempArr[count]));
								tempMaxLoad = int(each)

				maxOthers.append(tempMaxLoad)
				return maxOthers

				


					


















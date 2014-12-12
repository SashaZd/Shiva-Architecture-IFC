# -*- coding: utf-8 -*-

import uuid
import xlrd
from xlrd.sheet import ctype_text 
import re

#Global Variables ----> TODO : Take as user input later ?? 

IFC_FILE_PATH = "slab_segmentation_V4_3floors_SEEBIM.ifc"
EXCEL_FILE_PATH = "loadTables.xlsx"

# EXCEL_SHEET_1 = "slab_segmentation_V4_3floors_SE"

############################################################ 
################ HELPER CLASS ##############################
############################################################


class Helping:

############################################################
################ EXCEL CLASS METHODS #######################
############################################################


	def getExcelCellAtRowCol(self, sheetName, rows, cols):
		book = xlrd.open_workbook(EXCEL_FILE_PATH)
		for name in book.sheet_names():
			if name == sheetName:
				sheet = book.sheet_by_name(name)
				if rows<= sheet.nrows and cols<= sheet.ncols:
					return sheet.cell(rows, cols).value
				else:
					return ""

############ Sets User Defined global variables from Sheet 3 of the Excel File ############ 

	def setGlobalsForTheAlgorithm(self, callHelp):
		loadForProj = int(callHelp.getExcelCellAtRowCol("Sheet3", 3, 1))
		
		Wpref = int(re.findall('\d+', (callHelp.getExcelCellAtRowCol("Sheet3", 0, 1)))[0])
		
		Wplant_max = int(re.findall('\d+', (callHelp.getExcelCellAtRowCol("Sheet3", 1, 1)))[0])
		
		Wmin = int(re.findall('\d+', (callHelp.getExcelCellAtRowCol("Sheet3", 2, 1)))[0])

		return [loadForProj, Wpref, Wplant_max, Wmin]



################################################################
################ OTHER HELPER METHODS ########################## 
################################################################



############ Open either the Excel or IFC File ###################

	def openFile(self, filePath):
		return open(filePath, 'r')

############ Find a Substring ###################

	def findStringBetweeenStrings(self, originalString, startStr, endStr):
		try:
			startIndex = originalString.index(startStr) + len(startStr)
			endIndex = originalString.index( endStr, startIndex )
			return originalString[startIndex:endIndex]
		except ValueError:
			return ""

############ Find the Largest Index of the IFC File : Used to write at the end ###################

	def findLargestIndexOfIFCFile(self):
		ifcFile = self.openFile(IFC_FILE_PATH)
		for line in ifcFile:
			if ("#" in line) and (line[0]=="#"): 
				lastIndex = int(self.findStringBetweeenStrings(line, "#", "="))

		ifcFile.close()
		return lastIndex

############ Parse File and Return the ifcList List ###################
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


############ Find the Maximum Load that's allowed 			###################
############ Takes into consideration the 16+2 limitations 	###################


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
				tendonValue = 0
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
								if "16+2" not in str(sheet.cell((checkrow+1), col).value):
									# if tempArr[count] not in maxOthers:
										# maxOthers.append(int(tempArr[count]));
									tempMaxLoad = int(each)
									tendonValue = str(sheet.cell((checkrow+1), col).value)
									print tendonValue

				# maxOthers.insert(tempMaxLoad)
				return [[tempMaxLoad], tendonValue]



######################################################################
############ WRITING THE OUTPUT TO THE EXCEL / IFC FILE ##############
######################################################################


########### Generate Unique ID for each added slab #################

	def generateUUID(self):
		return str(uuid.uuid4().get_hex().upper()[0:22])
	


########### Find Level from Excel for Output #################

	def findLevelFromExcelForOutput(self, ifcList, slabOriginalIndex):
		indexInclHash = ""
		for ifcRel in ifcList:
			if "IFCRELCONTAINEDINSPATIALSTRUCTURE" in str(ifcRel) and str(slabOriginalIndex) in str(ifcRel):
					indexInclHash = re.split(',|\);', ifcRel)[-2]
					break


		for ifcBuildStor in ifcList:
			if "IFCBUILDINGSTOREY" in str(ifcBuildStor) and indexInclHash in str(ifcBuildStor):
				return ifcBuildStor.split(",")[2].split("'")[1].split(" ")[2]


		# return indexInclHashc



########### Write To IFC File #################

	def writeToIFCFile(self, ifcList, writeSlabIndex, Wdt, slabOriginalIndex, slabWidth, tendonValue):

		levelOfThisSlab = self.findLevelFromExcelForOutput(ifcList ,slabOriginalIndex)

		# print tendonValue

		# print "#"+str(writeSlabIndex)+"=IFCSLAB('" + self.generateUUID() + "',#41,'Floor:Precast Concrete Slab - 30 inch thick and "+str(Wdt)+" wide','"+levelOfThisSlab+" FL','slab length "+str(slabWidth)+"’  and number of tendons "+str(int(tendonValue))+",$,$, 'double_tee_slab_piece');"
	







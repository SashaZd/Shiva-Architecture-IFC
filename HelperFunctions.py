# -*- coding: utf-8 -*-
from collections import Counter
import os, random, string, re
import xlrd, xlwt
import math
# import xlsxwriter
from xlutils.copy import copy

#Global Variables ----> TODO : Take as user input later ?? 

IFC_FILE_PATH = "slab_segmentation_V4_3floors_SEEBIM.ifc"
EXCEL_FILE_PATH = "loadTables.xls"

TEMP_EXCEL_ARRAY = list()
OTHER_EXCEL_ARRAY = list()


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
		print "Iteration for Load: ", loadForProject, " || Slab: ", slabWidth
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

				# print checkrow, " -- ", checkcol

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
									tendonValue = re.split("or| ", tendonValue)[0]
									if not "+" in tendonValue:
										tendonValue = str(int(float(tendonValue)))
									else : 
										self.refactorTendonValue(tendonValue)

									# print temp

				# maxOthers.insert(tempMaxLoad)
				return [[tempMaxLoad], tendonValue]


	def findMaxStrands(self, sheetName, loadForProject, slabWidth, Wdt):

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

				# print checkrow, " -- ", checkcol

				#check to see if you've moved into another span
				for col in xrange(checkcol,checkcol+7):
					if col > sheet.ncols:
						pass
					else:
						tempArr = re.findall('\d+',sheet.cell(checkrow, col).value)
						count = -1
						for each in tempArr:
							count += 1
							WdtTemp = math.ceil(Wdt)
							if int(WdtTemp) == int(each):
								if str(slabWidth) == str(49.93) and str(Wdt) == str(8.5):
									print "Wdt ::", Wdt, " || Strands :: ", str(sheet.cell((checkrow+1), col).value), " || Coords :: ", (checkrow+1), col
								tendonValue = str(sheet.cell((checkrow+1), col).value)
								tendonValue = re.split("or| ", tendonValue)[0]
								return tendonValue


							# if int(each) > int(tempMaxLoad):
							# 	continue from here.... !! 
							# 	if "16+2" not in str(sheet.cell((checkrow+1), col).value):
							# 		# if tempArr[count] not in maxOthers:
							# 			# maxOthers.append(int(tempArr[count]));
							# 		tempMaxLoad = int(each)
							# 		tendonValue = str(sheet.cell((checkrow+1), col).value)
							# 		tendonValue = re.split("or| ", tendonValue)[0]
							# 		if not "+" in tendonValue:
							# 			tendonValue = str(int(float(tendonValue)))
							# 		else : 
							# 			self.refactorTendonValue(tendonValue)

									# print temp

				# # maxOthers.insert(tempMaxLoad)
				# return [[tempMaxLoad], tendonValue]

				# if str(slabWidth) == str(49.93):
				# 	print "Tendon : ", tendonValue, "Load : ", loadForProject, " || Length : ", slabWidth, " || Wdt : ", Wdt


				







	def refactorTendonValue(self, tendonValue):
		temp = tendonValue.split("+")[0]+" & number of rebars "+tendonValue.split("+")[1]
		if len(tendonValue.split("+"))==3:
			temp = temp + tendonValue.split("+")[2].split("k")[0] + "000 psi concrete"
		tendonValue = temp
		return tendonValue


######################################################################
############ WRITING THE OUTPUT TO THE EXCEL / IFC FILE ##############
######################################################################


########### Generate Unique ID for each added slab #################

	def generateUUID(self):
		return ''.join(random.choice(string.ascii_uppercase + string.ascii_lowercase + string.digits + "$") for i in range(22))
		# return str(uuid.uuid4().get_hex()[0:22])
	


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
		slabWidth = round(float(slabWidth), 2)
		Wdt = round(float(Wdt),2)
		levelOfThisSlab = self.findLevelFromExcelForOutput(ifcList ,slabOriginalIndex)
		addToFile = "#"+str(writeSlabIndex)+"= IFCSLAB('" + self.generateUUID() + "',#41,'Floor:Precast Concrete Slab - 30\" thick & "+str(Wdt)+" ft. wide','"+levelOfThisSlab+" FL','slab length "+str(slabWidth)+" ft. & number of strands "+str(tendonValue)+"',$,$, 'double_tee_slab_piece');\n"
		
		resultFile = open('Results.ifc', 'a')
		resultFile.write(addToFile)
		resultFile.close()


	def writeResultFileStart(self):
		originalFile =  self.openFile(IFC_FILE_PATH)
		countEndsec = 0;
		resultFile = open('Results.ifc', 'w+')

		for line in originalFile:
			if "ENDSEC" in line:
				countEndsec += 1
				if countEndsec == 2: 
					break
				else:
					resultFile.write(line)
			else:
				resultFile.write(line)

		resultFile.close()

	def writeResultFileEnd(self):
		originalFile =  self.openFile(IFC_FILE_PATH)
		countEndsec = 0
		resultFile = open('Results.ifc', 'a')

		for line in originalFile:
			if "ENDSEC" in line:
				countEndsec += 1
				if countEndsec >= 2: 
					# print line
					resultFile.write(line)

		resultFile.write("END-ISO-10303-21;")

		resultFile.close()
		originalFile.close()


############# WRITE TO EXCEL FILE ################

	def get_sheet_by_name(self, book, name):
	    # """Get a sheet by name from xlwt.Workbook, a strangely missing method.
	    # Returns None if no sheet with the given name is present.
	    # """
	    # Note, we have to use exceptions for flow control because the
	    # xlwt API is broken and gives us no other choice.
	    try:
	        for idx in itertools.count():
	            sheet = book.get_sheet(idx)
	            if sheet.name == name:
	                return sheet
	    except IndexError:
	        return None

	def getMemberElevation(self, ifcList, levelOfThisSlab):
		
		for each in ifcList:
			if "IFCBUILDINGSTOREY" in str(each) and levelOfThisSlab in str(each):
				return round(float(re.split(",|\);", each)[9]),2)


	def writeExcelFile(self, ifcList, slabOriginalIndex, slabWidth, Wdt, tendonValue, loadForProj):
		levelOfThisSlab = self.findLevelFromExcelForOutput(ifcList ,slabOriginalIndex)

		
		memberElevation = self.getMemberElevation(ifcList, levelOfThisSlab)
		

		TEMP_EXCEL_ARRAY.append(levelOfThisSlab+" FL," + str(memberElevation) + "," + str(slabWidth) + "," + str(Wdt) + "," + "1")
		OTHER_EXCEL_ARRAY.append(str(slabWidth) + "," + str(Wdt))

		currentIndex = len(TEMP_EXCEL_ARRAY)
		# print currentIndex

		

		# Find Wstruct_max list of all possibilities based on the 16+2 measure
		# [Wstruct_max_more, temp] = self.findMaxLoad("Sheet1", loadForProj, float(Wdt))
		
		# # Choose the topmost one for the rest of the calculation. Todo: Allow for more alternatives later
		# Wstruct_max = Wstruct_max_more[-1:]

		# print "Wdt: ", Wdt, " || Wstruct_max: ", Wstruct_max

		tendonValue = self.findMaxStrands("Sheet1", loadForProj, float(slabWidth), Wdt)
		# print tendonValue
		tempTendonValue = str(tendonValue).split("+")


		if len(tempTendonValue) >= 1:
			TEMP_EXCEL_ARRAY[currentIndex-1] += "," + tempTendonValue[0]
			OTHER_EXCEL_ARRAY[currentIndex-1] += "," + tempTendonValue[0]
		if len(tempTendonValue) >= 2:
			TEMP_EXCEL_ARRAY[currentIndex-1] += "," + tempTendonValue[1]
			OTHER_EXCEL_ARRAY[currentIndex-1] += "," + tempTendonValue[1]
		if len(tempTendonValue) >= 3:
			TEMP_EXCEL_ARRAY[currentIndex-1] += "," + tempTendonValue[2]
			OTHER_EXCEL_ARRAY[currentIndex-1] += "," + tempTendonValue[2]
		
		# print len(TEMP_EXCEL_ARRAY)


		# writableSheet = writableCopyFile.get_sheet(4)
		# print writableSheet.get_name()

		# writableSheet.write(0, 0, 'Combo I 3-4 year old')

		# book = xlrd.open_workbook(EXCEL_FILE_PATH)
		# for name in book.sheet_names():
		# 	if name == "Sheet4":
		# 		sheet = book.sheet_by_name(name)
		# 		# wb = copy(sheet)
		# 		sheet.write(0,0,"Hello")
						

	def writeFinalExcel(self):
		c = Counter(TEMP_EXCEL_ARRAY)
		d = Counter(OTHER_EXCEL_ARRAY)

		readOnlyFile = xlrd.open_workbook(EXCEL_FILE_PATH)
		writableCopyFile = copy(readOnlyFile)
		readOnlySheet = readOnlyFile.sheet_by_name("Sheet4")
		readOnlySheet2 = readOnlyFile.sheet_by_name("Sheet5")
		



		for i in range(0,4):
			if "Sheet4" in writableCopyFile.get_sheet(i).get_name():
				writableSheet = writableCopyFile.get_sheet(i)
				writableSheet2 = writableCopyFile.get_sheet(i+1)

				# For Sheet 4

				x=0
				while readOnlySheet.cell(x,0).value != "":
					x += 1
				
				for letter in c:
					# print 'NEW %s : %d' % (letter, c[letter])

					tempArr = letter.split(",")
					# if len(tempArr)>=7:
					# 	print tempArr
					
					# Floor Level
					writableSheet.write(x, 0, tempArr[0])

					# Top Member Elevation
					writableSheet.write(x, 1, round(float(tempArr[1]),2))

					# Member Length (Slab Width)
					writableSheet.write(x, 2, round(float(tempArr[2]),2))

					# Member Width (Wdt Or Wdt_last)
					writableSheet.write(x, 3, round(float(tempArr[3]),2))

					# Number of pieces
					writableSheet.write(x, 4, c[letter])

					#Number of Strands in each DT piece
					writableSheet.write(x, 5, round(float(tempArr[5]),2))

					#Total Number of Strands
					writableSheet.write(x, 6, round((float(tempArr[5]) * float(c[letter])),2))

					#Number of Rebars in each DT piece
					if len(tempArr)>=7:
						writableSheet.write(x, 7, round(float(tempArr[6]),2))

					#Total Number of Rebars
						temp_rebar = round((float(tempArr[6]) * float(c[letter])),2)
						# print temp_rebar
						writableSheet.write(x, 8, temp_rebar)

					#Number of Concrete PSI
					if len(tempArr)>=8:
						writableSheet.write(x, 9, tempArr[7])

					x += 1

		for i in range(0,4):
			if "Sheet5" in writableCopyFile.get_sheet(i).get_name():
				writableSheet2 = writableCopyFile.get_sheet(i)

				# For Sheet 5

				y=0
				while readOnlySheet2.cell(y,0).value != "" :
					y += 1
				
				for each in d:
					# print 'NEW %s : %d' % (each, c[each])
					tempArr2 = each.split(",")
					# print tempArr2

					# Member Length (Slab Width)
					writableSheet2.write(y, 0, round(float(tempArr2[0]),2))

					# Member Width (Wdt Or Wdt_last)
					writableSheet2.write(y, 1, round(float(tempArr2[1]),2))

					# Number of pieces
					writableSheet2.write(y, 2, d[each])

					#Number of Strands
					temp_strand = round(float(tempArr2[2]) * float(d[each]),2)
					writableSheet2.write(y, 3, temp_strand)

					#Number of Rebars
					if len(tempArr2)>=4:
						temp_rebars = round(float(tempArr2[3]) * float(d[each]),2)
						writableSheet2.write(y, 4, temp_rebars)

					#Number of Concrete PSI
					if len(tempArr2)>=5:
						writableSheet2.write(y, 5, tempArr2[4])

					# Total Number of DT Pieces in the Project
					
					if(y==1):
						writableSheet2.write(y, 6, len(TEMP_EXCEL_ARRAY))

					y += 1

		writableCopyFile.save(EXCEL_FILE_PATH)






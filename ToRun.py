from HelperFunctions import Helping
import re
import math

##################################################
################## ALGORITHM METHODS #############
##################################################

def setGlobalsForTheAlgorithm(callHelp):
	loadForProj = int(callHelp.getExcelCellAtRowCol("Sheet3", 3, 1))
	
	Wpref = int(re.findall('\d+', (callHelp.getExcelCellAtRowCol("Sheet3", 0, 1)))[0])
	
	Wplant_max = int(re.findall('\d+', (callHelp.getExcelCellAtRowCol("Sheet3", 1, 1)))[0])
	
	Wmin = int(re.findall('\d+', (callHelp.getExcelCellAtRowCol("Sheet3", 2, 1)))[0])

	return [loadForProj, Wpref, Wplant_max, Wmin]

def findWstruct_max():
	
	pass


############################################
################## CODE : MAIN #############
############################################

		# New helping class
callHelp = Helping()

		# Takes all the IFC Lines and puts it into an array so that they can be called using the index line number
ifcList = callHelp.parseFile()

		# Get Slab Width/Length & Slab Indexes to be used ahead
[ifcSlabWidth, ifcSlabIndexes] = callHelp.findIFCSlabWidth(ifcList)

[loadForProj, Wpref, Wplant_max, Wmin] = setGlobalsForTheAlgorithm(callHelp)

count = -1

# For each slab width find Wstruct_max
for slabWidth in ifcSlabWidth:
	print "------------------ new slab"
	count += 1 		#which slab is being called

	# Find Wstruct_max list of all possibilities based on the 16+2 measure
	Wstruct_max_more = callHelp.findMaxLoad("Sheet1", loadForProj, float(slabWidth))
	
	# Choose the topmost one for the rest of the calculation. Todo: Allow for more alternatives later
	Wstruct_max = Wstruct_max_more[-1:]
	print "Wstruct_max: ", Wstruct_max

	#Find slab line number
	currentSlabIndex = ifcSlabIndexes[count]

	Wbay = []

	for i in xrange(0,len(ifcList)):
		if "IFCRELAGGREGATES" in str(ifcList[i]):
			if "column_pair_and_beam_supporting_the_slab" in str(ifcList[i]).split(",")[2]:
				tempStr = str(ifcList[i]).split(",")[4]
				compareIndex = "#"+str(currentSlabIndex)
				if compareIndex in tempStr:
					tempStr2 = int(str(ifcList[i]).split(",")[3].split("'")[1].split(" ")[0])
					if tempStr2 not in Wbay:
						Wbay.append(tempStr2)

	# print Wbay

	maxW = min(Wplant_max, Wstruct_max[0])
	print "Max(W) = ", maxW

	typW = min(Wpref, Wstruct_max[0])
	print "Typ(W) = ", typW

	# RL = []
	for bay in Wbay:
		print "----- new bay : ", bay
		print "span: ", slabWidth, " || bay: ", bay
		
		# Remaining Length of bay when divided by Typ(W) calculated
		RL = bay%typW

		if RL==0:
			Wdt = typW
			DTquant = bay/typW
			DTquant_typ = bay/typW

			# print "Found RL==0"     # ," ::: Wdt=",Wdt," || DTquant=",DTquant," || DTquant_typ=",DTquant_typ

		else:
			if RL >= Wmin:
				Wdt = typW
				Wdt_last = RL
				DTquant_typ = math.floor(bay/typW)
				DTquant_last = 1
				DTquant = math.floor(bay/typW) + 1
				# print "Found RL >= Wmin"


			elif RL < Wmin and RL > (maxW - typW):
				Wdt = typW
				Wdt_last = (bay - typW * (math.floor(bay/typW)-1)) / 2
				DTquant = math.floor(bay/typW)+1
				DTquant_typ = math.floor(bay/typW)-1
				DTquant_last = 2

			counter = 1
			while Wdt_last < Wmin:
				print "inside loop: ", counter
				Wdt_last = (bay - typW * (math.floor(bay/typW)-1-counter)) / (2+counter)
				DTquant = str(math.floor(bay/typW)+1)+"?"
				DTquant_typ = math.floor(bay/typW)- 1 - counter
				DTquant_last = (2+counter)


			print "Wdt: ", Wdt, " || Wdt_last: ", Wdt_last

	













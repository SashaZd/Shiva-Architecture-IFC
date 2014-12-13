from collections import Counter
from HelperFunctions import Helping
import re
import math

##################################################
################## ALGORITHM METHODS #############
##################################################


def findWstruct_max():
	
	pass


############################################
################## CODE : MAIN #############
############################################

		# New helping class
callHelp = Helping()

callHelp.writeResultFileStart()

		# Takes all the IFC Lines and puts it into an array so that they can be called using the index line number
ifcList = callHelp.parseFile()

		# Get Slab Width/Length & Slab Indexes to be used ahead
[ifcSlabWidth, ifcSlabIndexes] = callHelp.findIFCSlabWidth(ifcList)

[loadForProj, Wpref, Wplant_max, Wmin] = callHelp.setGlobalsForTheAlgorithm(callHelp)

# NEED FOR FILE WRITE
largestIndexOfIFC = callHelp.findLargestIndexOfIFCFile()
writeSlabIndex = largestIndexOfIFC

count = -1
totSlabMade = []
anotherSlabCounter = []


# For each slab width find Wstruct_max
for slabWidth in ifcSlabWidth:

	slabMade = []
	print "------------------------ new slab"
	count += 1 		#which slab is being called

	# Find Wstruct_max list of all possibilities based on the 16+2 measure
	[Wstruct_max_more, tendonValue] = callHelp.findMaxLoad("Sheet1", loadForProj, float(slabWidth))
	
	# Choose the topmost one for the rest of the calculation. Todo: Allow for more alternatives later
	Wstruct_max = Wstruct_max_more[-1:]


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

		# print "------- bay :: ", bay, " || typW :: ", typW
		
		
		# Remaining Length of bay when divided by Typ(W) calculated
		RL = bay%typW

		if RL==0:
			Wdt = typW
			DTquant = bay/typW
			DTquant_typ = bay/typW
			DTquant_last = 0 

			# print "Found RL==0"     # ," ::: Wdt=",Wdt," || DTquant=",DTquant," || DTquant_typ=",DTquant_typ

		else:
			if RL >= Wmin:
				Wdt = typW
				Wdt_last = RL
				DTquant_typ = math.floor(bay/typW)
				DTquant = math.floor(bay/typW) + 1
				DTquant_last = 1
				# print "Found RL >= Wmin"

			elif RL < Wmin and RL <= (maxW - typW):
				RLnew = bay - (typW * (math.floor(bay/typW)-1))
				Wdt = typW
				Wdt_last = RLnew
				DTquant = math.floor(bay/typW)
				DTquant_typ = math.floor(bay/typW)-1
				DTquant_last = 1

			elif RL < Wmin and RL > (maxW - typW):
				Wdt = typW
				Wdt_last = (bay - (typW * (math.floor(bay/typW)-1))) / 2

				# if Wdt_last >= Wmin:
				DTquant = math.floor(bay/typW)+1
				DTquant_typ = math.floor(bay/typW)-1
				DTquant_last = 2

				counting = 1
				# print "outside while"
				while Wdt_last < Wmin:
					# print "inside loop: ", counting
					Wdt_last = (bay - typW * (math.floor(bay/typW)-1-counting)) / (2+counting)
					DTquant = math.floor(bay/typW)+1
					DTquant_typ = math.floor(bay/typW)- 1 - counting
					DTquant_last = (2+counting)
					counting = counting+1

		print "---------- bay : ", bay
		
		# print "values ::: ", DTquant, DTquant_typ, DTquant_last

		for z in range(0, (int(DTquant_typ))):
			slabMade.append(Wdt)
			# NEED FOR FILE WRITE
			writeSlabIndex = writeSlabIndex + 1

			# print ifcSlabIndexes[count]
			callHelp.writeToIFCFile(ifcList, writeSlabIndex, Wdt, ifcSlabIndexes[count], slabWidth, tendonValue)
			callHelp.writeExcelFile(ifcList, ifcSlabIndexes[count], slabWidth, Wdt, tendonValue)


		for z in range(0, (int(DTquant_last))):
			slabMade.append(Wdt_last)
			callHelp.writeToIFCFile(ifcList, writeSlabIndex, Wdt_last, ifcSlabIndexes[count], slabWidth, tendonValue)
			callHelp.writeExcelFile(ifcList, ifcSlabIndexes[count], slabWidth, Wdt_last, tendonValue)

		# print Counter(slabMade)

		for everySlab in slabMade:
			totSlabMade.append(everySlab)
		slabMade=[]

	# print Counter(totSlabMade)

	for everTotSlab in totSlabMade:
		anotherSlabCounter.append(everTotSlab)

	totSlabMade = []


# print "Final Count"
# print Counter(anotherSlabCounter)
# print "Total number :::: ", len(anotherSlabCounter)
		
callHelp.writeResultFileEnd()
callHelp.writeFinalExcel()
	













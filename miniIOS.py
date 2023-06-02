"""
MINI-IOS Simulator
"""

import math
import pandas as pd
import numpy as np
import openpyxl
from datetime import datetime
from datetime import timedelta, date

import xlsxwriter
import inspect #to print lineno
#import os
#import time
# import xlsxwriter module

#Manual Setting
SETTING_MODE = 0 # 0: Auto load from inputIOS.xlsx, 1: manual setting not from file
SETTING_PATH = "InputIOS.xlsx" #IOS files will be read from here
TARGET_TO_ACHIEVE =  3 #5 #T1 ioq1, T2 ioq2, T3 bf, T4 dec, T5 EOP (If 3 means from BF to EOP, If 5 means all, if 4 means T2~T5)
RAMP_DOWN_FROM = 1 #Minimum starting week of ramp down. 13 #999 12 #T3.Period #start from 1 not 0
RAMPDOWN_MODE = 1 #0: no ramp down, 1: 1 step rampdown
REPAIR_LINE_PCT = 0.126#0.126 #230130 avg repair line ratio to normal ~12.6%

DEBUG_LVL = 2 #DEBUG_LVL level (Higher number -> will have more details debug msg)

print("Mini IOS Started.. \n")


#################################################### Class/Function/Wrapper ######################################################

#shift and its information perday
class Shift(object):
    def __init__(self):
        self.Period = 0 #period Num (such as wk#x or day#x)
        self.Name = 0
        self.UPH = 0
        self.Yield = 0
        self.Efficiency = 0 
        self.WH = 0 #working hours per day
        self.WD = 0 #working days per week
        self.sNum = 0 
        self.IQ = self.UPH * self.WH * self.Efficiency * self.WD
        self.OQ = self.IQ * self.Yield #Output Qty
        self.sCount = 0

    def setOnline(self, name, UPH, Efficiency, Yield, WH, WD, pNum, sNum): #setOnline : no need name -- name only in init
        self.Period = pNum #period Num (such as wk#x or day#x)
        self.Name = name
        self.UPH = UPH
        self.Yield = Yield
        self.Efficiency = Efficiency 
        self.WH = WH #working hours per day
        self.WD = WD #working days per week
        self.sNum = sNum
        self.sCount = 1

        #IQ: Input Quantity, OQ: Output Quantity
        self.IQ = self.UPH * self.WH * self.Efficiency * self.WD #Input Qty
        self.OQ = self.IQ * self.Yield #Output Qty
        #print("IQ {}, OQ ".format(self.IQ, self.OQ))

    def setOffline(self):
        print("Line shift #{} @Wk# {} has been set to offline, Done!".format(self.Name, self.Period))
        self.Efficiency = 0 

        self.IQ = 0     #set this shift input qty to 0
        self.OQ = 0     #set this shift output qty to 0
        self.sCount = 0 #set this shift count to 0
        #print("IQ {}, OQ ".format(self.IQ, self.OQ))

    def output(self):

        #Output qty = UPH * Yield * Efficiency * Working Hours * Working Days * (1 - Pct of line for reinput/repair)
        self.OQ = self.UPH * self.Yield * self.Efficiency * self.WH * self.WD * (1-REPAIR_LINE_PCT)
        #print("IQ {}, OQ ".format(self.IQ, self.OQ))
        
        return self.OQ

    def show(self):
        print("Period#: ", self.Period )
        print("Name: ", self.Name )
        print("UPH: ", self.UPH )
        print("Yield: ", self.Yield)
        print("Efficiency: ", self.Efficiency)
        print("Working Hours/Day: ", self.WH)
        print("Working Days/Week: ", self.WD)
        print("Input/Week: ", self.IQ)
        print("Output/Week: ", self.OQ)
        print("\n")

#to hold Target T1, T2, T3 info
class Target(object):
    #to initiate info in Target
    def __init__(self, Name, pNum, Qty):
        self.Name = Name
        self.Period = pNum          #period Num (such as wk#x or day#x) -- Target Period
        self.Qty = Qty              #Target Qty
        self.AchievedPeriod = 0
        self.AchievedQty = 0
        self.AchievedPct = 0
        self.MetPeriod = 0
        self.MetQty = 0
        self.MetPct = 0

    #to populate achievement info (achieved period, qty and % to target qty)
    def achievementUpdate(self, AchievedPeriod, AchievedQty):
        self.AchievedPeriod = AchievedPeriod
        self.AchievedQty = AchievedQty
        self.AchievedPct = round(AchievedQty/self.Qty,4)
        #print("AchievedPeriod#: ", self.AchievedPeriod )
        #print("AchievedQty: ", self.AchievedQty )
        #print("AchievedPct: ", self.AchievedPct )

    #to update met target info (Period, qty and % to target qty)
    def metUpdate(self, MetPeriod, MetQty):
        self.MetPeriod = MetPeriod
        self.MetQty = MetQty
        self.MetPct = round(MetQty/self.Qty,4)
        #print("MetPeriod#: ", self.MetPeriod )
        #print("MetQty: ", self.MetQty )
        #print("MetPct: {:.2%}".format(self.MetPct))

    def show(self):
        print("{} \t Target Period#: {:,}, Target Qty: {:,} ::\t AchievedPeriod: {:,}, AchievedQty: {:,}, AchievedPct: {:.2%} ::\t MetPeriod: {:,}, MetQty: {:,}, MetPct: {:.2%}".format(self.Name, self.Period, self.Qty,self.AchievedPeriod, self.AchievedQty, self.AchievedPct,self.MetPeriod, self.MetQty, self.MetPct))
 

#Use to get efficencyProfile and YieldProfile
def getVal(Profile, period): #period start with 0
    if (period<len(Profile)):
        return Profile[period]
    else:
        return Profile[len(Profile)-1]

#use to update CUM WeeklyShipment
def updateCUM(PeriodicData, CUMData):
    CUMData[0] = PeriodicData[0] #CUM for data0 
    for i in range(len(CUMData)-1): # CUM data0 already done above so need deduct 1
        CUMData[i+1] = CUMData[i] + PeriodicData[i+1]
        
#Given the CUM shipment, this function check whether Target qty is met
def checkMetTarget(Target, CUMShipment, productionShipmentTime):
    #productionShipmentTime = 1 #2wks from production to shipment
    p = Target.Period-productionShipmentTime #-1
    if (DEBUG_LVL <= 1):
            print("{} Wk#{}: CUM shipment {:,}, target qty {:,}, achievement {:.2%}".format(Target.Name, Target.Period, round(CUMShipment[p],2), round(Target.Qty,2), round(CUMShipment[p]/Target.Qty,4)))

    if (CUMShipment[p] >= Target.Qty): # -1 because the index start from
        return 1
    else:  
        return 0

#To calculate: provided the period of production and the qty to achieve, this function return the number of shift required
#Note: OutputPerShift is required to count the number for shift required
def getShiftQtyToAdd(Qty, OutputPerShift, Period):
    if (DEBUG_LVL == 1):
        print("Qty: {:,}pcs".format(Qty))
        print("OutputPerShift: {:,}pcs".format(OutputPerShift))
        print("Period: {:,}pcs\n".format(Period))
    #return math.ceil(Qty/OutputPerShift/Period)
    return round(Qty/OutputPerShift/Period)

#Return the number of only shift qty for the corresponding period
def getShiftQty(Shifts, period):
    #print(len(Shifts))
    ret = 0 
    for s in range(len(Shifts)):
        ret = ret + Shifts[s][period].sCount
    #ret = ret - 1 #to return index instead of count
    print("period {}, in getShiftQty ret {}".format(period, ret))

    return ret

#Fill out achievement and met Target info
def fillAchievedToTarget(T, CUMShipment, productionShipmentTime):
    #printFrame()
    #to fill out the achieved qty and achieved period (same period as target period) 
    period = (T.Period)-productionShipmentTime #-1
    T.achievementUpdate(T.Period, CUMShipment[period])
    
    #to fill out the met qty and met period
    for i in range(len(CUMShipment)):
        if (CUMShipment[i] >= T.Qty):
            T.metUpdate(i+1, CUMShipment[i])
            #T.show()  
            break #return T

#provided start Date, Need to fill out Datelist
def fillDateList(StartDate, DateList):
    DateList[0] = StartDate
    for p in range(1,len(DateList)): #care DateList[0] assumed already filled out with the start date
         #DateList[p] = DateList[p-1] + timedelta(days=7)
         DateList[p] = datetime.strptime(DateList[p-1], "%m/%d/%y") + timedelta(days=7)
         DateList[p] = DateList[p].strftime("%m/%d/%y") #convert to datetime

def getNegativeSum(numbers):
    return sum(1 for number in numbers if number < 0)

def printFrame():
  callerframerecord = inspect.stack()[1]    # 0 represents this line
                                            # 1 represents line at caller
  frame = callerframerecord[0]
  info = inspect.getframeinfo(frame)
  #print(info.filename)                      # __FILE__     -> Test.py #print(__file__)
  print("{}(): line({})".format(info.function, info.lineno))    # __FUNCTION__ -> Main, print(info.lineno)      

#################################################### End of Class/Function/Wrapper ##################################################


###########################################################Load Setting #############################################################


if (SETTING_MODE == 0): #Automaticly load from setting file in SETTING_PATH
    InputSetting_df = pd.read_excel(SETTING_PATH).T
    #print(InputSetting_df)
    #below row and column location need to be updated to parameter in future
    FileName =  InputSetting_df[3][3] #line 8, column E [6][4]
    print(FileName)
    Description =  InputSetting_df[3][4] #line 8, column E [6][4]
    print(Description)
    StartDate =  InputSetting_df[5][4] #line 7, column E [5][4]
    LBU =  InputSetting_df[6][4] #line 8, column E [6][4]
    WH =  InputSetting_df[7][4] #line 9, column E [7][4]
    WD =  InputSetting_df[8][4] #line 10, column E [8][4]
    UPH =  InputSetting_df[9][4] #line 10, column E [8][4]
    EfficiencyProfile =  InputSetting_df[15][4:].values.tolist() #line 16, column E [16][4:]
    YieldProfile =  InputSetting_df[16][4:].values.tolist() #line 17, column E [17][4:]
    MPSWeekly = InputSetting_df[17][4:].values.tolist() #line 18, column E [18][4:] #TODO: Note setting files MPS need to be 0 don't put Empty, need to be able to replace NaN to 0 in future
    #print("{}, contains: {}".format(InputSetting_df[16][3], InputSetting_df[16][4:]))
    DateList = [0]*len(MPSWeekly) #for putting date info
    #print("MPS weekly {}, contains: {}".format(len(MPSWeekly), MPSWeekly))
    DateList[0] = StartDate.strftime("%m/%d/%y") #from datetime to string format
    fillDateList(DateList[0], DateList)

    #InputSetting_df[6][4] column E row 8, 
    T1 = Target('T1:IOQ1',InputSetting_df[10][5], InputSetting_df[10][7])
    T2 = Target('T2:IOQ2',InputSetting_df[11][5], InputSetting_df[11][7])
    T3 = Target('T3:BFDay',InputSetting_df[12][5], InputSetting_df[12][7])
    T4 = Target('T4:Dec-End',InputSetting_df[13][5], InputSetting_df[13][7])
    TEOP = Target('TEOP:EOP',InputSetting_df[14][5], InputSetting_df[14][7])

    #create and update MPS CUM
    TTLPeriod = len(MPSWeekly) 
    MPSCUM = [0]*TTLPeriod #MPSTotal = sum(MPSWeekly)
    updateCUM(MPSWeekly, MPSCUM)  
else:
    LBU = 8 #LBU per wk (LBU 8 shifts means 4 lines)
    WH = 10*0.9 #10 hours (-10% repair line)
    WD = 6
    UPH = 200
    EfficiencyProfile = [0.2286, 0.675, 0.9, 0.9] #0.925, 0.95]
    YieldProfile = [0.75, 0.8, 0.825, 0.85, 0.86, 0.87 , 0.88, 0.88, 0.89, 0.89, 0.9, 0.9, 0.91, 0.91, 0.915, 0.915, 0.92]
    #MPS profile 
    MPSWeekly= [ 103394,258484,310181,310181,310181,310181,310181,258484,258484,258484,206787,206787,103394,103394,103394,51697,51697,103394,103394,90986,51697,51697,51697,51697,51697,51697,64621,64621,64621,64621,64621,64621,64621,64621,51697,51697,51697,51697,51697,46297,0,0,0 ]
    #From https://docs.google.com/spreadsheets/d/1igfITfqrIvHPiqhq9bLGPKHSeEM7ZIZozW5t36W2Gn8/edit#gid=1066378962
    #MPSWeekly = [51501,76590,47169,128931,122168,109931,112823,113095,135178,104261,146887,137975,112692,71246,75447,72471,122037,102427,104873,157587,191040,175120,127360,185340,39004,143468,144998,138438,143886,127155,122990,103447,84201,85528,98720,63626,56366,43718,44644,46081,45333,41211,6260]
    DateList = [0]*len(MPSWeekly)
    DateList[0] = "07/16/23" #StartDate
    """
    for p in range(1,len(MPSWeekly)):
         #DateList[p] = DateList[p-1] + timedelta(days=7)
         DateList[p] = datetime.strptime(DateList[p-1], "%m/%d/%y") + timedelta(days=7)
         DateList[p] = DateList[p].strftime("%m/%d/%y")
    #print(DateList)
    """
    fillDateList(DateList[0], DateList)

    #create and update MPS CUM
    TTLPeriod = len(MPSWeekly) 
    MPSCUM = [0]*TTLPeriod #MPSTotal = sum(MPSWeekly)
    updateCUM(MPSWeekly, MPSCUM)  

    #with above MPS what MPS version?
    T1 = Target('T1:IOQ1',12, MPSCUM[11])
    T2 = Target('T2:IOQ2',13, MPSCUM[12])
    T3 = Target('T3:BFDay',19, MPSCUM[18])
    T4 = Target('T4:Dec-End',29, MPSCUM[28])
    TEOP = Target('TEOP:EOP',54, MPSCUM[53])

#print(getVal(YieldProfile, 18))
TargetList  = [0]* 5 #from T1, T2, T3, T4, TEOP, total 5 to reserve
TargetList[0] = T1
TargetList[1] = T2
TargetList[2] = T3
TargetList[3] = T4
TargetList[4] = TEOP

print("TTL Demand WKs: "+ str(TTLPeriod))
print("Total MPS: {:,}pcs".format(MPSCUM[len(MPSCUM)-1]))
print("UPH ",UPH)
print("WD ",WD)
print("WH ",WH)
print("LBU: "+ str(LBU))

#Setting for print out
DateList_df = pd.DataFrame({'Date':DateList}).T
LBU_df = pd.DataFrame({'LBU':[LBU]}).T
WH_df = pd.DataFrame({'Working Hours':[WH]}).T
WD_df = pd.DataFrame({'Working Days':[WD]}).T
UPH_df = pd.DataFrame({'UPH':[UPH]}).T
Eff_df = pd.DataFrame({'Efficiency':EfficiencyProfile}).T
Yield_df = pd.DataFrame(YieldProfile, columns=['Yield Profile']).T
MPSWeekly_df = pd.DataFrame(MPSWeekly, columns=['MPS Weekly']).T
MPSCUM_df = pd.DataFrame(MPSCUM, columns=['CUM MPS']).T
T1_df = pd.DataFrame(["Target Wk#:", T1.Period, "Target Qty", T1.Qty], columns=[T1.Name]).T
T2_df = pd.DataFrame(["Target Wk#:", T2.Period, "Target Qty", T2.Qty], columns=[T2.Name]).T
T3_df = pd.DataFrame(["Target Wk#:", T3.Period, "Target Qty", T3.Qty], columns=[T3.Name]).T
T4_df = pd.DataFrame(["Target Wk#:", T4.Period, "Target Qty", T4.Qty], columns=[T4.Name]).T
TEOP_df = pd.DataFrame(["Target Wk#:", TEOP.Period, "Target Qty:", TEOP.Qty], columns=[TEOP.Name]).T
combined = [DateList_df, LBU_df, WH_df, WD_df, UPH_df, T1_df, T2_df, T3_df, T4_df, TEOP_df, Eff_df, Yield_df, MPSWeekly_df, MPSCUM_df]

setting_df = pd.concat(combined) #.fillna(method='ffill') <-- avoid fillna all, use below direct method
setting_df.loc['LBU'] = setting_df.loc['LBU'].fillna('')
setting_df.loc['Working Hours'] = setting_df.loc['Working Hours'].fillna('')
setting_df.loc['Working Days'] = setting_df.loc['Working Days'].fillna('')
setting_df.loc['UPH'] = setting_df.loc['UPH'].fillna('')
#setting_df.loc['T1'] = setting_df.loc['T1'].fillna('')
#setting_df.loc['T2'] = setting_df.loc['T2'].fillna('')
#setting_df.loc['T3'] = setting_df.loc['T3'].fillna('')
#setting_df.loc['T4'] = setting_df.loc['T4'].fillna('')
#setting_df.loc['TEOP'] = setting_df.loc['TEOP'].fillna('')
setting_df.loc['Efficiency'] = setting_df.loc['Efficiency'].fillna(method = 'ffill')
setting_df.loc['Yield Profile'] = setting_df.loc['Yield Profile'].fillna(method = 'ffill')
print(setting_df)
############################################################### End of Setting #############################################################

#MaxShift = LBU * T3.Period
WeeklyShipment = [0]*TTLPeriod
CUMShipment = [0]*TTLPeriod

if (DEBUG_LVL == 1):
    print("row "+ str(len(Shifts)))
    print("col "+ str(len(Shifts[0])))
    print("\n")

#Average output per shift (based on the ramp profile of yield and efficiency)
AvgOuput = [0]*TTLPeriod
output = 0

#calc AvgOutput per period
for pNum in range(TTLPeriod):
    output = output + UPH * getVal(EfficiencyProfile, pNum) * getVal(YieldProfile, pNum) * WH * WD * (1-REPAIR_LINE_PCT)  #weekly output
    AvgOuput[pNum] = round(output/(pNum+1), 2)

#Calculate number of shifts required to fullfil T1, T2, T3, TEOP
#Set output period from production to shipment
productionShipmentTime = 1  #Assume 1 wk gap, need to shift 1 wk from GB output production to shipment
TargetShift = [0]* 5 #from T1, T2, T3, T4, TEOP, total 5 to reserve
#For each target milestone (IOQ1/IOQ2/BF) and its qty, calculate how many shift required
for i in range(len(TargetList)):
    TargetShift[i] = getShiftQtyToAdd(TargetList[i].Qty, AvgOuput[TargetList[i].Period-productionShipmentTime-1], (TargetList[i].Period-productionShipmentTime-2)) #deduct -1 due to index start from 0
    print("Target#{}, Shift est: {}, avg output{}".format(TargetList[i].Name, TargetShift[i], AvgOuput[TargetList[i].Period-productionShipmentTime-2]))

# Get max shift est based on each provided target
MaxShift = 0 #initialized to 0
LoopTo = min(len(TargetList), TARGET_TO_ACHIEVE)
print("Target criteria: ", LoopTo)
#Get the max shift 
for i in range(LoopTo):
    curIndex = len(TargetShift) - i -1
    MaxShift = max(MaxShift, TargetShift[curIndex]) #to achieved the target, how many shifts required, the max shift is the most number of shifts required
print("\nEst Max Shift: {:}\n".format( MaxShift))

############################################################### Ramp Up Start #############################################################

#Initialized
InitShift = 2 #initial shift that already exist before ramp (usually from PVT lines)
CurMaxShift = [0]*TTLPeriod
PeriodArray = [[0 for x in range(TTLPeriod)] for y in range(MaxShift)]
EffArray = [[0 for x in range(TTLPeriod)] for y in range(MaxShift)]
TempEffArray = EffArray.copy()
OutputArray = [[0 for x in range(TTLPeriod)] for y in range(MaxShift)]
TempOutputArray = OutputArray.copy()
YieldArray = [[0 for x in range(TTLPeriod)] for y in range(MaxShift)]
Shifts = [[Shift() for x in range(TTLPeriod)] for y in range(MaxShift)]

#Ramp-up: Add line shifts
#start s shift loop
for s in range(MaxShift): #loop from 0 to Maxshift - 1 (Total loop# is MaxShift)
    #start pNum period loop (for each shift, loop into each period, period by period)
    for pNum in range(TTLPeriod-1): #max week = shipment week -1
        #print("shift {}, day# {}".format( s, pNum))
        #Set maximum number of shift per each period for the ramping
        if (CurMaxShift[pNum] == 0): #if first time not yet initiated
            if (pNum == 0):
                CurMaxShift[0] = InitShift + LBU #max period is initial shift qty + number of new line bring up (LBU) shift
            else:
                CurMaxShift[pNum] = CurMaxShift[pNum-1] + LBU #max period is previous shift qty + LBU shift qty
            
        if (s >= CurMaxShift[pNum]): #max shift reached, no more ramp-up (no more shift can be added)
            continue;

        #calculate the relative production wk for each shift and each period 
        #this information will be used to load/calculate the corresponding yield and efficiency ramp 
        if (pNum == 0): #initialized the periodArray for the first time
            PeriodArray[s][0] = 1 #start from 0, index 0 is period#1 (week#1) #221205 1 make periodArray work but it bypass the first value of efficiencyRmamp
        else: #increase counter from previous period
             PeriodArray[s][pNum] = PeriodArray[s][pNum-1] + 1

        p = PeriodArray[s][pNum]-1 #for yield and efficiency ramp period, index start from 0
        #(name, UPH, Efficiency, Yield, WH, WD, pNum, sNum):
        #setOnline
        Shifts[s][pNum] = Shift() #create shift for this shift s and this period pNum
        Shifts[s][pNum].setOnline("FATP-"+str(s), UPH, getVal(EfficiencyProfile, p), getVal(YieldProfile, p), WH, WD, pNum, s) #set online
        WeeklyShipment[pNum+1] = WeeklyShipment[pNum+1] + round(Shifts[s][pNum].output()) #Output of this week, will be the shipment of the next week
        EffArray[s][pNum] = getVal(EfficiencyProfile, p)
        OutputArray[s][pNum] = round(Shifts[s][pNum].output()) #230111 check if need to be recalculated after ramp down
        YieldArray[s][pNum] = getVal(YieldProfile, p)

#in case need to check the PeriodArray, EffArray, YieldArray
if DEBUG_LVL == 1:
    print(PeriodArray)
    print("\n")
    print(EffArray)
    print("\n")
    print(YieldArray)
    print("\n")

#update WeeklyShipment to CUMShipment array
updateCUM(WeeklyShipment, CUMShipment)

EffArray_np = np.array(EffArray)
EffArray_df = pd.DataFrame(EffArray_np)
OutputArray_np = np.array(OutputArray)
#OutputArray_df = pd.DataFrame(OutputArray_np) #outputArray_df will be created after RampDown, for calculation OutputArray will be use (_df use for write to excel)

#Update CUMBuffer 
CUMBuffer = [ CUMShipment[x] - MPSCUM[(x)] for x in range(TTLPeriod)]

############################################################### End of Ramp Up #############################################################

############################################################# Start of RampDown ############################################################
#Preparing for ramping down
if (RAMPDOWN_MODE >= 1): #Currently is forward reduction 
    print("RAMPDOWN_MODE: {}, 1st RampDown started..".format(RAMPDOWN_MODE))
    printFrame()

    #begin loop to get the valley and reduce the buffer up to 1 week buffer
    #T2 
    #TODO 221206 Need to apply Temp to actual only once per each p loop
    for p in range(RAMP_DOWN_FROM-1, TTLPeriod): #line reduction up to final week -1 #why need -1 in 230111 remove it
        
        BUFFER_MODE = 0 #ramp down from the valley
        if (BUFFER_MODE == 0):
            #if do ramp down from the valley
            minIndex = CUMBuffer.index(min(CUMBuffer[p:])) #221205 update s to p
            minBuf = min(CUMBuffer[p:]) 
        else:
            #if just using current buffer 
            minBuf = CUMBuffer[p] #230116
            minIndex = p #230116

        print("current week idx {}, min index {}, min buffer {}, weeklyQty {}, avgOutput {}".format(p, minIndex, minBuf, WeeklyShipment[p], AvgOuput[minIndex]))
        deltaPeriod = minIndex - p  #if have not reach threshold find the minimum valley
        
        #This is parameter that need to be optimized
        #bufThreshold = round(MPSCUM[len(MPSCUM)-1]*0.04) #0.08, 0.02, 0.04
        
        bufThreshold = round(max(WeeklyShipment) * (2/6)) #2/6 2days of total 6days/week
        print("Buffer Threshold {}".format(bufThreshold))
        #if CUMBuffer[minIndex]>=bufThreshold: #should be compared to weekly qty or direct compare to buffer
        if minBuf>=bufThreshold: #should be compared to weekly qty or direct compare to buffer
            print("reach Threshold")
            deltaQtyToReduce =  minBuf - bufThreshold #Target qty to reduce, also need to reduce if minBuf exceed certain level
            deltaPeriod = 1 #delta period from the valley to current period, force to reduce)    #230110 need to check why previously
        else: 
            deltaQtyToReduce =  minBuf - WeeklyShipment[minIndex-1] #min buffer threshold to be maintained up to last week shipment qty
        

        print("deltaPeriod {}, minIndex {}, p {}, deltaQtyToReduce {:}".format(deltaPeriod, minIndex, p, deltaQtyToReduce))
        if ((deltaQtyToReduce >= 0) & (deltaPeriod >0)):    #deltaPeriod > 0, this criteria will cause it to stop Ramp if current period is the valley of the buffer qty
            deltaShiftToReduce = math.ceil(deltaQtyToReduce/AvgOuput[minIndex]/deltaPeriod)
            print("current week idx {} to reduce {} shift \n".format(p, deltaShiftToReduce)) 
            print("period p {}: deltaShiftToReduce {} = deltaQtyToReduce {} / AvgOuput[minIndex] {}  / deltaPeriod {}\n".format(p, deltaShiftToReduce, deltaQtyToReduce, AvgOuput[minIndex], deltaPeriod)) 

            #start to reduce shift from p to minIndex for total qty shift deltaShiftToReduce
            #print(getShiftQty(Shifts, p))
            for shiftToReduce in range(0, deltaShiftToReduce, 1):  #230111 change range(1, to range(0 start from 0 index instead of 1
                print ("(1)in shiftToReduce loop..\n")
                for period in range(p,TTLPeriod-1): #-1): #shipment and production delta 1 week #230111 need to double check -1
                    print("\n")
                    print("period {}, TTLPeriod {}".format(period, TTLPeriod))
                    #reduce shift from the last active shift
                    #ex: max 16 to reduce 1 
                    #setOffline
                    #impactedShiftIndex = getShiftQty(Shifts, period)-shiftToReduce # 221205 remove -1 #because index start from 0 
                    impactedShiftIndex = getShiftQty(Shifts, period) -1 # 221205 remove -1 #because index start from 0 
                    #if impactedShiftIndex < 0:
                    #    print("impactedShiftIndex < 0 %d", impactedShiftIndex)
                    #    continue
                    print("shift to reduce {} of {}, period {}, impactedShiftIndex {}".format(shiftToReduce, deltaShiftToReduce, period, impactedShiftIndex))
                    
                    print("plan set offline shift index {} @period {}".format(impactedShiftIndex, period))
                    #Do buffer projection after reduction
                    tempOutput = Shifts[impactedShiftIndex][period].output()*1.05 #230112 overdo 5% for the projection
                    CurReduceQty = 0
                    projWeekly = 0
                    result = 1 #result 1 means OK, result 0 means the reduction will cause negative buffer 
                    # up to TTLPeriod-1 the last week is not needed as last week of production is shipment week -1
                    for periodCheck in range (period, TTLPeriod-1): # start from period + 1, shift 1   #230111 need to double check -1 or -2
                        CurReduceQty = CurReduceQty + tempOutput    
                        #print("periodCheck ", periodCheck)
                        #if periodCheck+1 == TTLPeriod-1: #230111 need to check 
                        #    continue
                        if (CUMBuffer[periodCheck+1] - CurReduceQty < 0):
                            result = 0
                            print("shift check#{},est in period {}, negative buffer {}, cancel shift removal".format(impactedShiftIndex, periodCheck, CUMBuffer[periodCheck] - CurReduceQty))
                            break
                        else:
                            print("plan OK2set: in period {}, OK buffer {}".format(periodCheck, CUMBuffer[periodCheck] - CurReduceQty))
                        
                        projWeekly = WeeklyShipment[periodCheck+1] - tempOutput
                        projMaxQty = projWeekly * (TTLPeriod-periodCheck)


                    if (result == 0): #if buffer projection become negative, don't do the reduction
                        continue 

                                  
                    #print("before update: period+1 {}, WeeklyShipment {}, Shifts[impactedShiftIndex][period].output() {}".format(period+1, WeeklyShipment[period+1], Shifts[impactedShiftIndex][period].output() ))
                    WeeklyShipment[period+1] = WeeklyShipment[period+1] - Shifts[impactedShiftIndex][period].output() #current week line down, the shipment qty impact in the next week
                    updateCUM(WeeklyShipment, CUMShipment)
                    #print("afer update: period+1 {}, WeeklyShipment {}, Shifts[impactedShiftIndex][period].output() {}".format(period+1, WeeklyShipment[period+1], Shifts[impactedShiftIndex][period].output() ))

                    CUMBuffer = [ CUMShipment[x] - MPSCUM[(x)] for x in range(TTLPeriod)] #MPSCUM[(x+1)] -> MPSCUM[(x)]

                    EffArray[impactedShiftIndex][period] = 0
                    OutputArray[impactedShiftIndex][period] = 0
                    print("## actual set offline shift index {} @period {}\n".format(impactedShiftIndex, period))
                    Shifts[impactedShiftIndex][period].setOffline()   
                    #print("==> Shift {}, Period {}, CUM buffer {}".format(shiftToReduce, period, CUMBuffer[(RAMP_DOWN_FROM-1):]))
                #End of for period in range(p,TTLPeriod):
            #End of for shiftToReduce in range(1, deltaShiftToReduce, 1): 
        else:
            deltaShiftToReduce = 0  #TODO need to check
            print ("nothing to reduce..(deltaQtyToReduce,\n")

        #print ("exit shiftToReduce loop..\n")
        
        updateCUM(WeeklyShipment, CUMShipment)
        CUMBuffer = [ CUMShipment[x] - MPSCUM[(x)] for x in range( len(MPSCUM) )] #MPSCUM[(x+1)] -> MPSCUM[(x)]
        

############################################################# End of RampDown ############################################################


for i in range(len(TargetList)):
    checkMetTarget(TargetList[i], CUMShipment, productionShipmentTime)
    fillAchievedToTarget(TargetList[i], CUMShipment, productionShipmentTime)     
    TargetList[i].show()

#ShiftQty = [getShiftQty(Shifts,p) for p in range(TTLPeriod)]
Description_df = pd.DataFrame({'Description' : [FileName + " (" + Description + "; Peak Shift QTY: " + str(MaxShift) + " (" + str(math.ceil(MaxShift/2)) + "lines) )"]}, index = ["Scenario"] )
TargetPeriod_df = pd.DataFrame({'Target Period' : [TargetList[i].Period for i in range(len(TargetList))]}, index = [TargetList[i].Name for i in range(len(TargetList))])
#TargetPeriod_df = pd.DataFrame({'Target Period' : [T1.Period, T2.Period, T3.Period, T4.Period, TEOP.Period]}, index = [T1.Name, T2.Name, T3.Name, T4.Name, TEOP.Name])
TargetQty_df = pd.DataFrame({'Target Qty' : [TargetList[i].Qty for i in range(len(TargetList))]}, index = [TargetList[i].Name for i in range(len(TargetList))])
#TargetQty_df = pd.DataFrame({'Target Qty' : [T1.Qty, T2.Qty, T3.Qty, T4.Qty, TEOP.Qty]}, index = [T1.Name, T2.Name, T3.Name, T4.Name, TEOP.Name])
AchievedPeriod_df = pd.DataFrame({'Achieved Period' : [TargetList[i].AchievedPeriod for i in range(len(TargetList))]}, index = [TargetList[i].Name for i in range(len(TargetList))])
#AchievedPeriod_df = pd.DataFrame({'Achieved Period' : [T1.AchievedPeriod, T2.AchievedPeriod, T3.AchievedPeriod, T4.AchievedPeriod, TEOP.AchievedPeriod]}, index = [T1.Name, T2.Name, T3.Name, T4.Name, TEOP.Name])
AchievedQty_df = pd.DataFrame({'Achieved Qty' : [TargetList[i].AchievedQty for i in range(len(TargetList))]}, index = [TargetList[i].Name for i in range(len(TargetList))])
#AchievedQty_df = pd.DataFrame({'Achieved Qty' : [T1.AchievedQty, T2.AchievedQty, T3.AchievedQty, T4.AchievedQty, TEOP.AchievedQty]}, index = [T1.Name, T2.Name, T3.Name, T4.Name, TEOP.Name])
AchievedPct_df = pd.DataFrame({'Achieved Pct' : [TargetList[i].AchievedPct for i in range(len(TargetList))]}, index = [TargetList[i].Name for i in range(len(TargetList))])
#AchievedPct_df = pd.DataFrame({'Achieved Pct' : [T1.AchievedPct, T2.AchievedPct, T3.AchievedPct, T4.AchievedPct, TEOP.AchievedPct]}, index = [T1.Name, T2.Name, T3.Name, T4.Name, TEOP.Name])
MetPeriod_df = pd.DataFrame({'Met Period' : [TargetList[i].MetPeriod for i in range(len(TargetList))]}, index = [TargetList[i].Name for i in range(len(TargetList))])
MetQty_df = pd.DataFrame({'Met Qty' : [TargetList[i].MetQty for i in range(len(TargetList))]}, index = [TargetList[i].Name for i in range(len(TargetList))])
MetPct_df = pd.DataFrame({'Met Pct' : [TargetList[i].MetPct for i in range(len(TargetList))]}, index = [TargetList[i].Name for i in range(len(TargetList))])
combined = [TargetPeriod_df, TargetQty_df, AchievedPeriod_df, AchievedQty_df, AchievedPct_df, MetPeriod_df, MetQty_df, MetPct_df]
Summary_df = pd.concat(combined, axis=1)

print(Summary_df)

#print( EffArray)
EffArray_np = np.array(EffArray)
#print(EffArray_np)
EffArray_df = pd.DataFrame(EffArray_np)
#print(EffArray_df)

#https://www.aivia-software.com/post/python-quick-tip-3-thresholding-with-numpy
OutputArray_np = (EffArray_np > 0.5) * OutputArray_np  
OutputArray_df = pd.DataFrame(OutputArray_np)

#Append MPS and Shipment info at end of line
ShiftQty = [getShiftQty(Shifts,p) for p in range(TTLPeriod)]
ShiftQty_df = pd.DataFrame({'Shift QTY' : ShiftQty}).T
WeeklyShipment_df = pd.DataFrame({'Weekly Shipment' : WeeklyShipment}).T
CUMShipment_df = pd.DataFrame({'CUM Shipment' : CUMShipment}).T
MPSCUM_np = MPSCUM_df.to_numpy()
CUMShipment_np = CUMShipment_df.to_numpy()
DeltaShipment_np = np.subtract(CUMShipment_np, MPSCUM_np)
DeltaShipment_df = pd.DataFrame(DeltaShipment_np, index = ["Delta (CUM Shipment-CUM MPS)"])
if (DEBUG_LVL == 1):
    print(DeltaShipment_df)
    print("MPS max {:n}".format(np.amax(MPSCUM_np) ))
    print("MaxShift {}".format(MaxShift))
    print("UPH {}".format(UPH))
    print("EfficiencyProfile {:.1%}".format(max(EfficiencyProfile) ))
    print("YieldProfile {:.1%}".format(max(YieldProfile) ))
    print("TEOP.Period) {}".format(TEOP.Period) )
Utilization = np.amax(MPSCUM_np) /(MaxShift * UPH * max(EfficiencyProfile) * max(YieldProfile) * WH * WD * TEOP.Period)#Note: WH incl. reinput/repair 
Utilization_df = pd.DataFrame({'TTL Utilization':[round(Utilization,4)]}).T
print("Utilization {:.1%}".format(Utilization)) 

combined = [DateList_df, ShiftQty_df, MPSWeekly_df, MPSCUM_df, WeeklyShipment_df, CUMShipment_df, DeltaShipment_df, Utilization_df, EffArray_df, OutputArray_df]
CapacityParameter_df = pd.concat(combined)
#print("### Weekly shipment/shift #17 ", WeeklyShipment[47])#/ShiftQty[46])

############################################################### Writing output file ###########################################################################################
T_stamp = datetime.now().strftime('_%Y%m%d_%H%M%S')

OutputFile = FileName + '_OutputIOS' + T_stamp + '.xlsx'
with pd.ExcelWriter(OutputFile) as writer:
    Summary_df.to_excel(writer, sheet_name = 'Summary', startcol = 3, startrow = 5)  
    Description_df.to_excel(writer, sheet_name = 'Summary', startcol = 4, startrow = 3, header = False, index = False)  
    CapacityParameter_df.to_excel(writer, sheet_name ='CapacityParameter', startcol = 3, startrow = 5)
    setting_df.to_excel(writer, sheet_name = 'CapacityParameter', startcol = 3, startrow = 20+MaxShift*2)  #index = False, header = False
    setting_df.to_excel(writer, sheet_name = 'Setting', startcol = 3, startrow = 5)  #index = False, header = False
    
    #Formating the workbook
    workbook  = writer.book
    formatDate = workbook.add_format({'num_format': 'mm/dd/yy'}) #  #,##0.00
    formatCommas = workbook.add_format({'num_format': '#,##0'}) #  #,##0.00
    formatPct = workbook.add_format({'num_format': '0%'})
    formatBold = workbook.add_format({'bold': 'True', 'font_size' : '14'})
    
    #all index start from 0
    worksheet = writer.sheets['CapacityParameter']
    
    #set column size
    worksheet.set_column(3, 3, 26)
    worksheet.set_column(4, 4, 12)
    worksheet.set_column(6, 6, 12)

    #set row attribute
    worksheet.set_row(6, 15, formatDate) #row#7, height size 15, formatDate 
    
    WSfirstRow = 7 #row 7, index   (header: 7(blank)+7, lineQty*2 + 7(blank) +)
    for i in range(6):
        worksheet.set_row(WSfirstRow+i, 15, formatCommas)
    
    WSutilizationRow = WSfirstRow + 6 
    worksheet.set_row(WSutilizationRow, 15, formatPct) 

    WSfirstRow = WSutilizationRow + MaxShift + 1
    for i in range(MaxShift):
        worksheet.set_row(WSfirstRow+i, 15, formatCommas)


    #for the setting row
    SettingLineInWS_Eff = 7+7 + 2*MaxShift + 7 + 12
    for i in range(2):
        worksheet.set_row(SettingLineInWS_Eff+i, 15, formatCommas)
    
    worksheet.freeze_panes(6,4)

    worksheet2 = writer.sheets['Setting']
    worksheet2.set_column(3, 3, 18)
    worksheet2.set_column(4, 4, 12)
    worksheet2.set_column(6, 6, 12)
    
    
    worksheet2.set_row(10, 15, formatCommas)
    worksheet2.set_row(11, 15, formatCommas)
    worksheet2.set_row(12, 15, formatCommas)
    worksheet2.set_row(13, 15, formatCommas)
    worksheet2.set_row(14, 15, formatCommas)
    
    worksheet2.set_row(18, 15, formatCommas)
    worksheet2.set_row(19, 15, formatCommas)
    worksheet2.freeze_panes(6,4)


    worksheet3 = writer.sheets['Summary']
    worksheet3.set_row(3, 15, formatBold)
    worksheet3.set_column(3, 11, 14)
    worksheet3.set_column(4, 7, 14, formatCommas)
    worksheet3.set_column(8, 8, 14, formatPct)
    worksheet3.set_column(9, 10, 14, formatCommas)
    worksheet3.set_column(11, 11, 14, formatPct)
    worksheet3.freeze_panes(6,4)

########################################################### Start Graph Drawing ################################################################################
    # Here we create a column chart object .
    # This will use as the primary chart.
    column_chart2 = workbook.add_chart({'type': 'column'}) 
    column_chart2.add_series({
        'name':       '=CapacityParameter!$D$8',      # "shift qty"
        'categories': '=CapacityParameter!$E$7:$BF$7', #x "period"
        'values':     '=CapacityParameter!$E$8:$BF$8', #y "number of shift"
    })

    # Create a new line chart.
    # This will use as the secondary chart.
    line_chart2 = workbook.add_chart({'type': 'line'})

    # Configure the data series for the secondary chart.
    # We also set a secondary Y axis via (y2_axis).
    line_chart2.add_series({
        'name':       '=CapacityParameter!$D$10',
        'categories': '=CapacityParameter!$E$7:$BF$7',
        'values':     '=CapacityParameter!$E$10:$BF$10',
         'y2_axis':    True,
    })

    line_chart2.add_series({
        'name':       '=CapacityParameter!$D$12',
        'categories': '=CapacityParameter!$E$7:$BF$7',
        'values':     '=CapacityParameter!$E$12:$BF$12',
         'y2_axis':    True,
    })
 
    # Combine both column and line charts together.
    column_chart2.combine(line_chart2)
     
    # Add a chart title 
    column_chart2.set_title({ 'name': 'Line Shift QTY and CUM MPS vs CUM Shipment'})
     
    # Add x-axis label
    column_chart2.set_x_axis({'name': 'Week Start Date'})
     
    # Add y-axis label
    column_chart2.set_y_axis({'name': 'Line Shift QTY (Shift)'})
     
    # Note: the y2 properties are on the secondary chart.
    line_chart2.set_y2_axis({'name': 'CUM Shipment (pcs)'})
    
    #https://xlsxwriter.readthedocs.io/chart.html
    column_chart2.set_size({'width': 920, 'height': 376})
    column_chart2.set_legend({'position': 'bottom'})

    # add chart to the worksheet with given
    # offset values at the top-left corner of
    # a chart is anchored to cell E15
    worksheet3.insert_chart('E15', column_chart2, {'x_offset': 25, 'y_offset': 10})
 
########################################################### End of Graph Drawing ################################################################################


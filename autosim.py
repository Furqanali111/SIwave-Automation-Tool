#!/usr/bin/env python
# coding: utf-8

# In[43]:


import csv
import os
import sys

script_path=os.path.abspath(__file__)
script_path = os.path.dirname(script_path)
filepath=script_path+"\path.txt"

with open(filepath, 'r') as f:
    content = f.read().replace('\n', '')

arg101=content
file_path_padstack_usage =arg101+ '/padstack_usage.csv'
# file_path_padstack_usage = 'padstack_usage.csv'
print('arg101:', arg101)

# Read the CSV file from the starting line
def read_csv_file_after_string(file_path_padstack_usage,search_string):
    padstack_usage = []
    found_string = False
    with open(file_path_padstack_usage, 'r') as csvfile:
        csvreader = csv.reader(csvfile)
        for row in csvreader:
            if not found_string:
                if search_string in row:
                    found_string = True
            else:
                padstack_usage.append(row)
    return padstack_usage

search_string = 'Detailed Padstack Usage'
padstack_usage = read_csv_file_after_string(file_path_padstack_usage, search_string)
# print('padstack_usage:', padstack_usage)

# Search for BGA padstack name
keyword = 'BGA'
BGA_padstack = None

for row in padstack_usage:
    if keyword in row[0]:
        BGA_padstack = row[1]
        break

if BGA_padstack:
    print("BGA_padstack:", BGA_padstack)
else:
    print("Keyword not found in the padstack_usage.")


# Search for DIE padstack name
keyword = 'DIE'
DIE_padstack = None

for row in padstack_usage:
    if keyword in row[0]:
        DIE_padstack = row[1]
        break

if DIE_padstack:
    print("DIE_padstack:", DIE_padstack)
else:
    print("Keyword not found in the padstack_usage.")


#extract BGA padstack data
padstack_csv_file = arg101+ '/Padstack_Definition_Report.csv'
start_keyword = 'Padstack: ' + BGA_padstack + ' '
end_keyword = 'Drill Data for ' + BGA_padstack + ''

def read_csv_file_between_keywords(csv_file, start_keyword, end_keyword):
    data = []
    reading_data = False
    with open(csv_file, 'r') as csvfile:
        csvreader = csv.reader(csvfile)
        for row in csvreader:
            if not row:  # Skip empty rows
                continue

            if not reading_data:
                if row and row[0].startswith(start_keyword):
                    reading_data = True
            elif row and row[0].startswith(end_keyword):
                break
            else:
                data.append(row)
    return data

padstack_data = read_csv_file_between_keywords(padstack_csv_file, start_keyword, end_keyword)
print(padstack_data)

#find BGA radius and height
BGA_radius = None

for row in padstack_data:
    if row[0].startswith('L') and row[2] == 'CIRCLE':
        BGA_radius = float(row[3]) / 1000
        break

if BGA_radius is not None:
    print("BGA_radius:", BGA_radius)
else:
    print("The required conditions were not met in the data.")
BGA_height = ((0.5 / 1000) * ((BGA_radius / 1000) * (BGA_radius / 1000) * 3.1415926) / (1 / 7000000)) * 1000

#extract DIE padstack data
def read_csv_file_between_keywords(csv_file, start_keyword, end_keyword):
    data = []
    reading_data = False
    with open(csv_file, 'r') as csvfile:
        csvreader = csv.reader(csvfile)
        for row in csvreader:
            if not row:  # Skip empty rows
                continue

            if not reading_data:
                if row and row[0].startswith(start_keyword):
                    reading_data = True
            elif row and row[0].startswith(end_keyword):
                break
            else:
                data.append(row)
    return data

# padstack_csv_file = arg101+'\Padstack_Definition_Report.csv'
start_keyword = 'Padstack: ' + DIE_padstack + ' '
end_keyword = 'Drill Data for ' + DIE_padstack + ''

padstack_data = read_csv_file_between_keywords(padstack_csv_file, start_keyword, end_keyword)
print(padstack_data)

#find DIE radius and height
DIE_radius = None

for row in padstack_data:
    if row[0].startswith('L') and row[2] == 'CIRCLE':
        DIE_radius = float(row[3]) / 1000
        break

if DIE_radius is not None:
    print("BGA_radius:", DIE_radius)
else:
    print("The required conditions were not met in the data.")

DIE_height = ((0.5 / 1000) * ((DIE_radius / 1000) * (DIE_radius / 1000) * 3.1415926) / (1 / 7000000)) * 1000



# Edit/set materials (step 4 - 6)

# first edit copper globally
oDoc.ScrEditMaterial('conductor', 'copper', 5.0E+07, 0.999991)
if oDoc.ScrEditMaterial('conductor', 'copper', 5.0E+07, 0.999991) == 0:
    oDoc.ScrLogMessage('Successful: Edited copper conductivity to 5E+7')

# step 6
layernamelist = oDoc.ScrGetLayerNameList()
# delete air at top
for i in range(len(layernamelist)):
    if layernamelist[i:i + 2] == ['TOP_SM', 'L1']:
        deleteabove = layernamelist[0:i]
        break
for layer in deleteabove:
    oDoc.ScrDeleteLayer(layer)
    # if oDoc.ScrDeleteLayer(layer) == 1:
    oDoc.ScrLogMessage('Successful: Deleted layer: ' + layer)
    # else:
    #     oDoc.ScrLogMessage('Failed: Deleted layer: ' + layer)

# delete air at bottom
for i in range(len(layernamelist)):
    if layernamelist[i] == 'BOT_SM' and layernamelist[i - 1].startswith('L'):
        deletebelow = layernamelist[i + 1:]

for i in range(len(deletebelow)):
    oDoc.ScrDeleteLayer(deletebelow[i])
    # if oDoc.ScrDeleteLayer(deletebelow[i]) == 1:
    oDoc.LogMessage('Successful: Deleted layer: ' + deletebelow[i])
    # else:
    #     oDoc.LogMessage('Failed: Deleted layer: ' + deletebelow[i])

# step 6
# assign new_copper to all metal layers
layernamelist_left = oDoc.ScrGetLayerNameList()
for i in range(1, len(layernamelist_left), 2):
    oDoc.ScrSetLayerMaterial(layernamelist_left[i], 'copper')

# PTH (step 7)
# set PTH padstack via plating
# padstacknamelist = oDoc.ScrGetPadstackNameList()
# for string in padstacknamelist:
#     if "PTH" in string:
#         oDoc.ScrSetPadstackViaPlatingAbsolute(string, 0.015mm)


# power_nets = oDoc.ScrGetNetNameList()
# power_nets = ['VDDCR_CPU', 'VDDCR_SOC', 'VSS']
power_nets = ['VDDCR', 'VDD_MEM']
# power_nets = ['VDD', 'VDDIO', 'VDDM_0', 'VDDM_1', 'VDDM_7']


##(delete existing pin groups)
partName_refDes = oDoc.ScrGetComponentList('integrated circuits, input/output')
intCir = partName_refDes[0]  # get the first element of the list
partName_DIE = intCir.split()[0]  # split the first element into words and get the first word
refDes_DIE = intCir.split()[1]
exist_pingroups_DIE = oDoc.ScrGetPinGroupNameList(partName_DIE, refDes_DIE)
for i in exist_pingroups_DIE:
    oDoc.ScrDeletePinGroup(i, True)
    # if oDoc.ScrDeletePinGroup(i, True) == 1:
    #     oDoc.ScrLogMessage('Successful: Deleted existing duplicated pin groups: ' + i)
    # else:
    #     oDoc.ScrLogMessage('Failed: Deleted existing duplicated pin groups: ' + i)

inPutoutPut = partName_refDes[1]
partName_BGA = inPutoutPut.split()[0]  # split the first element into words and get the first word
refDes_BGA = inPutoutPut.split()[1]
exist_pingroups_BGA = oDoc.ScrGetPinGroupNameList(partName_BGA, refDes_BGA)
for i in exist_pingroups_BGA:
    oDoc.ScrDeletePinGroup(i, True)
    # if oDoc.ScrDeletePinGroup(i, True) ==1:
    #     oDoc.ScrLogMessage('Successful: Deleted existing duplicated pin groups: ' + i)
    # else:
    #     oDoc.ScrLogMessage('Failed: Deleted existing duplicated pin groups: ' + i)

## (delete existing current source) ##
# power_nets = oDoc.ScrGetNetNameList()
exist_csourcelist = oDoc.ScrGetComponentList('Current Sources')
for i in power_nets:
    if ('DIE_DIE_' + i + '_Group_BGA_' + i + '_Group') in exist_csourcelist:
        oDoc.ScrDeleteCktElem('DIE_DIE_' + i + '_Group_BGA_' + i + '_Group')
        # if oDoc.ScrDeleteCktElem('DIE_DIE_' + i + '_Group_BGA_' + i + '_Group') ==1:
        oDoc.ScrLogMessage(
            'Successful: Deleted existing duplicated current source: ' + 'DIE_DIE_' + i + '_Group_BGA_' + i + '_Group')
        # else:
        #     oDoc.ScrLogMessage('Failed: Deleted existing duplicated current source: ' + 'DIE_DIE_' + i + '_Group_BGA_' + i + '_Group')


#set solder ball for BGA and DIE (step 9)
oDoc.ScrSetSolderballMaterial(BGA_padstack, 'solder')
oDoc.ScrAssignSimpleSolderballProfile(BGA_padstack, BGA_height, BGA_radius, 1, 0)

oDoc.ScrSetSolderballMaterial(DIE_padstack, 'solder')
oDoc.ScrAssignSimpleSolderballProfile(BGA_padstack, BGA_height, BGA_radius, 0, 1)

# Set DIE/BGA pin groups  (step 11 - 12)
for i in power_nets:
    oDoc.ScrCreatePinGroupByNet(partName_DIE, refDes_DIE, i, 'DIE_' + i + '_Group', False)
    # if oDoc.ScrCreatePinGroupByNet(partName_DIE,refDes_DIE, i , 'DIE_'+i+'_Group',False) ==1:
    oDoc.ScrLogMessage('Successful: Created pin groups: ' + 'DIE_' + i + '_Group')
    # else:
    #     oDoc.ScrLogMessage('Failed: Created pin groups: ' + 'DIE_' + i + '_Group')
    oDoc.ScrCreatePinGroupByNet(partName_BGA, refDes_BGA, i, 'BGA_' + i + '_Group', False)
    # if oDoc.ScrCreatePinGroupByNet(partName_BGA,refDes_BGA, i , 'BGA_'+i+'_Group',False) ==1:
    oDoc.ScrLogMessage('Successful: Created pin groups: ' + 'BGA_' + i + '_Group')
    # else:
    #     oDoc.ScrLogMessage('Failed: Created pin groups: ' + 'BGA_'+i+'_Group')
# Set Simulation Options
path = os.getcwd()
path = path.decode("utf-8")
oDoc.ScrImportSIwaveSimulationOptions(path + '\\DCSim_4P_Settings.sws')
# if oDoc.ScrImportSIwaveSimulationOptions(path+'\\DCSim_4P_Settings.sws')==1:
oDoc.ScrLogMessage('Successful: Set simulation options')
# else:
#     oDoc.ScrLogMessage('Failed: Set simulation options')

# Set current sources (step 14) ##
for i in power_nets:
    oDoc.ScrPlaceCircuitElement('DIE_DIE_' + i + '_Group_BGA_' + i + '_Group', 'CurrentSourceGroup', 4, 1, 'DIE', 'DIE',
                                'DIE_' + i + '_Group', 1, 'BGA', 'BGA', 'BGA_' + i + '_Group', 0.0, 0.0, 5E7, 0.0, 1.0,
                                0.0)
    # if oDoc.ScrPlaceCircuitElement('DIE_DIE_' + i + '_Group_BGA_' + i + '_Group', 'CurrentSourceGroup', 4,1,'DIE', 'DIE','DIE_' + i + '_Group',1, 'BGA','BGA','BGA_' +i+'_Group',0.0,0.0,5E7,0.0,1.0,0.0) ==1:
    oDoc.ScrLogMessage('Successful: Created current sources: ' + 'DIE_DIE_' + i + '_Group_BGA_' + i + '_Group')
    # if oDoc.ScrPlaceCircuitElement('DIE_DIE_' + i + '_Group_BGA_' + i + '_Group', 'CurrentSourceGroup', 4,1,'DIE', 'DIE','DIE_' + i + '_Group',1, 'BGA','BGA','BGA_' +i+'_Group',0.0,0.0,5E7,0.0,1.0,0.0) ==1:
    #     oDoc.ScrLogMessage('Failed: Created current sources: ' + 'DIE_DIE_' + i + '_Group_BGA_' + i + '_Group')

# Auto run simulation ##
csourcelist = oDoc.ScrGetComponentList('Current Sources')
csourcelist = [s.split(' ')[1] for s in csourcelist]

for i in csourcelist:
    oDoc.ScrActivateCktElem(i, 'csource', False)

for i, j in zip(csourcelist, power_nets):
    index = csourcelist.index(i)
    oDoc.ScrActivateCktElem(i, 'csource', True)
    oDoc.ScrActivateCktElem(csourcelist[index - 1], 'csource', False)
    oDoc.ScrSetSimulationName('dc', 'DC IR Sim ' + j)
    oDoc.ScrSetIdealGroundNodeInDcSimulation(i, 1)
    oDoc.ScrRunDcSimulation(1)

# ## Extract simulation results (step 19 - 21)
for i in power_nets:
    if i != 'VSS':
        checksim = oDoc.ScrReadDCLoopResInfo("DC IR Sim " + i, [], [])
        if checksim == 0:
            oDoc.ScrExportElementData("DC IR Sim " + i, i + "_IMAX.dc", "Vias")

for i in power_nets:
    if i != 'VSS':
        checksim = oDoc.ScrReadDCLoopResInfo("DC IR Sim " + i, [], [])
        if checksim == 0:
            oDoc.ScrExportElementData("DC IR Sim " + i, i + "_DCR.dc", "Current Sources")



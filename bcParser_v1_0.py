import sys
import fileinput
import os
import getopt
import re
import numpy as np
import configparser
from termcolor import colored,cprint #REQUIRES PIP INSTALL termcolor
from tabulate import tabulate #requires pip install tabulate
import math


#path = os.path.split(os.getcwd())[0]
#case = os.path.split(os.getcwd())[1]

def assignVar(variable):
   
   if variable in varList:
      varIndex = np.where(varList == variable)
      variableValue = varVal[varIndex[0]]
      return variableValue[0]
   else:
      variableValue = False
      return variableValue

def magnitude(vector):
    return math.sqrt(sum(pow(element, 2) for element in vector))

def bcParser(path,case):
    global varList,varVal
    caseSetupPath="%s/%s/caseSetup" % (path,case)
    config = configparser.ConfigParser()
    config.optionxform = str
    config.read_file(open(caseSetupPath))
    configSections = config.sections()
    variableSet = {}
    varList = []
    varVal = []
    for section in configSections:
        variableSet[section] = np.array(config.items(section)) 
        if not "GEOM" in section:
            #print("  Reading %s values:" % (section))
            pass
        for item in np.array(config.items(section)):
            if not "GEOM" in section:
                #print("     %s: %s" % (item[0],item[1]))
                varList = np.append(varList,item[0])
                varVal = np.append(varVal,item[1])
   
    with open('%s/%s/system/controlDict' % (path,case)) as controlDict:

        lines = controlDict.readlines()
        time = []
        for line in lines:

            if line.startswith("endTime"):

                time.append(re.findall(r"[-+]?(?:\d*\.*\d+)", line)[0])
                lastTime = time[0]
                

            elif line.startswith("application"):

                application = re.search('application\s* (.+?);',line).group(1)

    with open('%s/%s/system/caseProperties' % (path,case)) as caseProperties:

        lines = caseProperties.readlines()

        wheelRotation = False

        for line in lines:
            if "U" in line:

                inletVel = np.array(re.findall('[0-9]+',line))

                n=0
				
                inletVel = map(float,inletVel)

                inletVel = list(inletVel)
				#print(inletVel)

            elif "rotating" in line:
                wheelRotation = True

            elif "ground" in line:
                groundFlag = 1
                wheelFlag = 0

            elif "*wh*" in line:
                wheelFlag = 1 
                groundFlag = 0
		
        inletMag = assignVar('INLETMAG')
        yaw = assignVar('YAW')
		#print(wheelRotation)

    with open('%s/%s/constant/turbulenceProperties' % (path,case)) as turbProperties:
        lines = turbProperties.readlines()

        for line in lines:
            if "simulationType" in line:
                simType = re.search('simulationType\s* (.+?);',line).group(1)
                #print(simType)
            elif "RASModel" in line:
                turbModel = re.search('RASModel\s* (.+?);',line).group(1)
                #print(turbModel)
            elif "LESModel" in line:
                turbModel = re.search('LESModel\s* (.+?);',line).group(1)



    return inletMag,lastTime,yaw,wheelRotation,simType,turbModel






import os
import sys
import numpy as np
import pandas as pd
from pptx import Presentation
from datetime import date
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from bcParser_v1_0 import *
from pptx.oxml.xmlchemy import OxmlElement



#### Functions ####
def SubElement(parent, tagname, **kwargs):
    element = OxmlElement(tagname)
    element.attrib.update(kwargs)
    parent.append(element)
    return element


def _set_cell_border(cell, border_color="000000", border_width='500'):
    """ Hack function to enable the setting of border width and border color
        - bottom border only at present
        (c) Steve Canny
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    lnR = SubElement(
        tcPr, 'a:lnR', w=border_width, cap='flat', cmpd='sng', algn='ctr')
    solidFill = SubElement(lnR, 'a:solidFill')
    srgbClr = SubElement(solidFill, 'a:srgbClr', val=border_color)
    return cell
################################################################################################

print("#### Running Post Pro Report Generator v1.0 ####")


todays_date = date.today()
path = os.path.split(os.getcwd())[0]
case = os.path.split(os.getcwd())[1]
casePath = os.path.split(path)[0]
job = os.path.split(casePath)[1]


#### TEMPLATE PATH ####
templatePath = "/home/openfoam/openFoam/templates/pptTemplate/XXXXXX-Template_Presentation.pptx"
print("TEMPLATE PATH: %s" % (templatePath))

#Getting number of trials compared
numTrials = len(sys.argv)


caseArray = [case]

#making the list of trials compared
print("Getting data for trials:")
print("%s" % (case))
for i in sys.argv[1:]:
	caseArray.append(i)
	print(i)



#initializing the arrays for the trial boundary conditions
inletMagArray = []
lastTimeArray = []
yawArray = []
wheelRotationArray = []
simTypeArray = []
turbModelArray = []
caseSymArray = []

for trial in caseArray:

	[inletMag,lastTime,yaw,wheelRotation,simType,turbModel]=bcParser(path,trial)
	inletMagArray.append(inletMag)
	lastTimeArray.append(lastTime)
	yawArray.append(yaw)
	wheelRotationArray.append(wheelRotation)
	simTypeArray.append(simType)
	turbModelArray.append(turbModel)
	if "_half" in trial:
		caseSym = "Half Car"
		caseSymArray.append(caseSym)
	else:
		caseSym = "Full Car"
		caseSymArray.append(caseSym)




#### Setting up Presentation ####

#getting the presentation template
prs = Presentation(templatePath)

print("Making title slide...")
#setting up the title slide
titleSlideLayout = prs.slide_layouts[0]
titleSlide = prs.slides.add_slide(titleSlideLayout)
title = titleSlide.shapes.title
subtitle = titleSlide.placeholders[1]
title.text = "%s - JOB NAME" % (job)
trialsInReport = ' | '.join(caseArray)
subtitle.text = "%s" % (trialsInReport)


#### INFO SLIDE SET UP ####
print("Making info slide...")
infoTableLayout = prs.slide_layouts[5]
infoTableSlide = prs.slides.add_slide(infoTableLayout)
infoTableTitle = infoTableSlide.shapes.title
infoTableTitle.text = "Trial Setup and BC" #set the title of the trial set up and boundary conditions slide
table_placeholder = infoTableSlide.shapes[1]
nCols = numTrials + 1 #number of columns is number of trials plus 1
shape = table_placeholder.insert_table(rows=7, cols=nCols)
table = shape.table

#making row lables
rowLabelText = ['Trial','Run Type','Velocity (m/s)','Turb. Model','Wheel Rot.','Yaw Angle (deg)','Sym. Cond.']
#adding text to the rowLabels
i = 0
for text in rowLabelText:
	table.cell(i,0).text = text
	i = i + 1

#adding data for each trials column
i = 0 
while i < numTrials:
	table.cell(0,i+1).text = caseArray[i]
	table.cell(1,i+1).text = simTypeArray[i]
	table.cell(2,i+1).text = str(inletMagArray[i])
	table.cell(3,i+1).text = turbModelArray[i]
	table.cell(4,i+1).text = str(wheelRotationArray[i])
	table.cell(5,i+1).text = str(yawArray[i])
	table.cell(6,i+1).text = caseSymArray[i]
	i = i + 1


for j in range(numTrials + 1):
	for i in range(len(rowLabelText)):
		
		cell = table.cell(i,j)
		cell = _set_cell_border(cell)
		table.cell(i,j).fill.solid()
		table.cell(i,j).text_frame.paragraphs[0].font.size = Pt(14)
		table.cell(i,j).text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)
		if i == 0:
			table.cell(i,j).fill.fore_color.rgb = RGBColor(192,192,192)
		else:
			table.cell(i,j).fill.fore_color.rgb = RGBColor(225,225,225)


#### RESULT TABLE ####
print("Checking for existance of averaged results...")

coeffs = np.zeros((6,numTrials))
n = 0
for trial in caseArray:
	avgFile = "trial%s_AVG_all_coeff.csv" % (trial)
	if os.path.isfile("%s/%s/%s" % (path,trial,avgFile)):
		print("		%s found... importing data..." % (avgFile))
		# Defining the columns to read
		
		
		avgData = np.loadtxt("%s/%s/%s" % (path,trial,avgFile), dtype='float', comments='#', delimiter=",",skiprows = 1,usecols=(1,2,3,4,7,8), unpack=False, ndmin=0)
		
		coeffs[:,n] = avgData

	else: 
		print("		Cannot find average data... skipping...")

	n = n + 1





print("Making results slide...")
resultsTableLayout = prs.slide_layouts[5]
resultsTableSlide = prs.slides.add_slide(resultsTableLayout)
resultsTableTitle = resultsTableSlide.shapes.title
resultsTableTitle.text = "Results" #set the title of the trial set up and boundary conditions slide
table_placeholder = resultsTableSlide.shapes[1]
nCols = numTrials + 1 #number of columns is number of trials plus 1
shape = table_placeholder.insert_table(rows=8, cols=nCols)
table = shape.table

#making row lables
rowLabelText = ['Trial','CdA','ClA','ClfA','ClrA','0.95 CI. - CdA','0.95 CI. - ClA','% Front']
#adding text to the rowLabels
i = 0
for text in rowLabelText:
	table.cell(i,0).text = text
	i = i + 1

#adding data for each trials column
i = 0 
percentFrontArray = []

while i < numTrials:

	cl = (coeffs[1,i])
	clf = (coeffs[2,i])
	percentFront = 100*clf/cl
	percentFrontArray.append(percentFront)

	i = i + 1
	

i = 0
while i < numTrials:
	#naming first row in column as trial name
	table.cell(0,i+1).text = caseArray[i]

	j = 1
	while j < 7:
		#assigning each value a number
		
		if i == 0:
			if "_half" in caseArray[i] and j < 5:
				table.cell(j,i+1).text = str("%0.4f" % (coeffs[j-1,i]*2))
			else:
				table.cell(j,i+1).text = str("%0.4f" % (coeffs[j-1,i]))
		else:
			if "_half" in caseArray[i] and j < 5:
				table.cell(j,i+1).text = str("%0.4f (%0.2f %%)" % (coeffs[j-1,i]*2,100*(coeffs[j-1,i]-coeffs[j-1,0])/coeffs[j-1,0]))
			else:
				table.cell(j,i+1).text = str("%0.4f" % (coeffs[j-1,i]))
		j = j + 1
	
	table.cell(7,i+1).text = str("%0.2f"%(percentFrontArray[i]))
	
	i = i + 1


for j in range(numTrials + 1):
	for i in range(len(rowLabelText)):
		
		cell = table.cell(i,j)
		cell = _set_cell_border(cell)
		table.cell(i,j).fill.solid()
		table.cell(i,j).text_frame.paragraphs[0].font.size = Pt(14)
		table.cell(i,j).text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)
		if i == 0:
			table.cell(i,j).fill.fore_color.rgb = RGBColor(192,192,192)
		else:
			table.cell(i,j).fill.fore_color.rgb = RGBColor(225,225,225)


#### Confidence Plot SLIDES ####
print("Checking for Confidence Plot image existance...")
plotArray = ['cdConfPlot','clConfPlot']

for trial in caseArray:
	print("		Checking in trial%s" % (trial))
	for plot in plotArray	:
		plotImage = "trial%s_%s.png" % (trial,plot)
		if os.path.isfile("%s/%s/%s" % (path,trial,plotImage)):
			print("			Found %s..." % (plotImage))
		else:
			print("			Cannot find %s... will skip in report..." % (plotImage))

print("Making confidence plot slides...")
for plot in plotArray:
	for trial in caseArray:
		confPlotLayout = prs.slide_layouts[6]
		confPlotSlide = prs.slides.add_slide(confPlotLayout)
		confPlotSlideTitle = confPlotSlide.shapes.title
		confPlotSlideTitle.text = "%s - %s" % (plot,trial) #set the title of the geom slides
		confPlot_placeholder = confPlotSlide.shapes[1]
		plotImage = "trial%s_%s.png" % (trial,plot)
		if os.path.isfile("%s/%s/%s" % (path,trial,plotImage)):
			insertConfPlotImage = confPlot_placeholder.insert_picture("%s/%s/%s" % (path,trial,plotImage))
		else:
			pass

#### Development Plot SLIDES ####
print("Creating development plots...")
if len(caseArray) < 2:
    caseCommandString = ""
    command = "python3.8 /home/openfoam/openFoam/scripts/binPlotForces_v2_0.py -s -n"
else:
    caseCommandString = " ".join(caseArray[1:])
    command = "python3.8 /home/openfoam/openFoam/scripts/binPlotForces_v2_0.py -t %s -s -n" % (caseCommandString)
    
caseString = "_".join(caseArray)
os.system(command)
plotArray = ["cd-development_%s" % (caseString),"cl-development_%s" % (caseString)]


print("		Checking in trial%s" % (trial))
for plot in plotArray:
    plotImage = "%s.png" % (plot)
    if os.path.isfile("%s/%s/%s" % (path,case,plotImage)):
        print("			Found %s..." % (plotImage))
    else:
        print("			Cannot find %s... will skip in report..." % (plotImage))

print("Making development plot slides...")
for plot in plotArray:

    confPlotLayout = prs.slide_layouts[6]
    confPlotSlide = prs.slides.add_slide(confPlotLayout)
    confPlotSlideTitle = confPlotSlide.shapes.title
    confPlotSlideTitle.text = "%s" % (plot) #set the title of the geom slides
    confPlot_placeholder = confPlotSlide.shapes[1]
    plotImage = "%s.png" % (plot)
    if os.path.isfile("%s/%s/%s" % (path,case,plotImage)):
        insertConfPlotImage = confPlot_placeholder.insert_picture("%s/%s/%s" % (path,case,plotImage))
    else:
        pass




#### GEOMETRY SLIDES ####
print("Checking for geometry image existance...")
viewsArray = ['front','frontLeft','left','bottom','rearLeft','rear']

for trial in caseArray:
	print("		Checking in trial%s" % (trial))
	for view in viewsArray:
		geomImage = "%s_geom_%s.png" % (trial,view)
		if os.path.isfile("%s/%s/postProcessing/images/%s" % (path,trial,geomImage)):
			print("			Found %s..." % (geomImage))
		else:
			print("			Cannot find %s... will skip in report..." % (geomImage))

print("Making geometry slides...")

for view in viewsArray:
	for trial in caseArray:
		geomLayout = prs.slide_layouts[4]
		geomSlide = prs.slides.add_slide(geomLayout)
		geomSlideTitle = geomSlide.shapes.title
		geomSlideTitle.text = "Geometry - %s - %s" % (trial,view) #set the title of the geom slides
		geom_placeholder = geomSlide.shapes[1]
		geomImage = "%s_geom_%s.png" % (trial,view)
		if os.path.isfile("%s/%s/postProcessing/images/%s" % (path,trial,geomImage)):
			insertGeomImage = geom_placeholder.insert_picture("%s/%s/postProcessing/images/%s" % (path,trial,geomImage))
		else:
			pass


#### CP SLIDES ####
print("Checking for CP image existance...")
viewsArray = ['front','frontLeft','left','bottom','rearLeft','rear']

for trial in caseArray:
	print("		Checking in trial%s" % (trial))
	for view in viewsArray:
		cpImage = "%s_cP_%s.png" % (trial,view)
		if os.path.isfile("%s/%s/postProcessing/images/%s" % (path,trial,cpImage)):
			print("			Found %s..." % (cpImage))
		else:
			print("			Cannot find %s... will skip in report..." % (cpImage))

print("Making Cp slides...")
for view in viewsArray:
	for trial in caseArray:
		cpLayout = prs.slide_layouts[4]
		cpSlide = prs.slides.add_slide(cpLayout)
		cpSlideTitle = cpSlide.shapes.title
		cpSlideTitle.text = "Cp Plot - %s - %s" % (trial,view) #set the title of the geom slides
		cp_placeholder = cpSlide.shapes[1]
		cpImage = "%s_cP_%s.png" % (trial,view)
		if os.path.isfile("%s/%s/postProcessing/images/%s" % (path,trial,cpImage)):
			insertCpImage = cp_placeholder.insert_picture("%s/%s/postProcessing/images/%s" % (path,trial,cpImage))
		else:
			pass

#### CPX SLIDES ####
print("Checking for CPX image existance...")
viewsArray = ['front','frontLeft','left','bottom','rearLeft','rear']

for trial in caseArray:
	print("		Checking in trial%s" % (trial))
	for view in viewsArray:
		cpXImage = "%s_cPx_%s.png" % (trial,view)
		if os.path.isfile("%s/%s/postProcessing/images/%s" % (path,trial,cpXImage)):
			print("			Found %s..." % (cpXImage))
		else:
			print("			Cannot find %s... will skip in report..." % (cpXImage))

print("Making CpX slides...")
for view in viewsArray:
	for trial in caseArray:
		cpXLayout = prs.slide_layouts[4]
		cpXSlide = prs.slides.add_slide(cpXLayout)
		cpXSlideTitle = cpXSlide.shapes.title
		cpXSlideTitle.text = "CpX Plot - %s - %s" % (trial,view) #set the title of the geom slides
		cpX_placeholder = cpXSlide.shapes[1]
		cpXImage = "%s_cPx_%s.png" % (trial,view)
		if os.path.isfile("%s/%s/postProcessing/images/%s" % (path,trial,cpXImage)):
			insertCpXImage = cpX_placeholder.insert_picture("%s/%s/postProcessing/images/%s" % (path,trial,cpXImage))
		else:
			pass

#### CPZ SLIDES ####
print("Checking for CPZ image existance...")
viewsArray = ['front','frontLeft','left','bottom','rearLeft','rear']

for trial in caseArray:
	print("		Checking in trial%s" % (trial))
	for view in viewsArray:
		cpZImage = "%s_cPz_%s.png" % (trial,view)
		if os.path.isfile("%s/%s/postProcessing/images/%s" % (path,trial,cpZImage)):
			print("			Found %s..." % (cpZImage))
		else:
			print("			Cannot find %s... will skip in report..." % (cpZImage))

print("Making CpZ slides...")
for view in viewsArray:
	for trial in caseArray:
		cpZLayout = prs.slide_layouts[4]
		cpZSlide = prs.slides.add_slide(cpZLayout)
		cpZSlideTitle = cpZSlide.shapes.title
		cpZSlideTitle.text = "CpZ Plot - %s - %s" % (trial,view) #set the title of the geom slides
		cpZ_placeholder = cpZSlide.shapes[1]
		cpZImage = "%s_cPz_%s.png" % (trial,view)
		if os.path.isfile("%s/%s/postProcessing/images/%s" % (path,trial,cpZImage)):
			insertCpZImage = cpZ_placeholder.insert_picture("%s/%s/postProcessing/images/%s" % (path,trial,cpZImage))
		else:
			pass



#### UMeanNear SLIDES ####
print("Checking for UMeanNear image existance...")
viewsArray = ['frontLeft','left','bottom','rearLeft','rear']

for trial in caseArray:
	print("		Checking in trial%s" % (trial))
	for view in viewsArray:
		UMeanNearImage = "%s_UMeanNear_%s.png" % (trial,view)
		if os.path.isfile("%s/%s/postProcessing/images/%s" % (path,trial,UMeanNearImage)):
			print("			Found %s..." % (UMeanNearImage))
		else:
			print("			Cannot find %s... will skip in report..." % (UMeanNearImage))

print("Making UMeanNear slides...")
for view in viewsArray:
	for trial in caseArray:
		UMeanNearLayout = prs.slide_layouts[4]
		UMeanNearSlide = prs.slides.add_slide(UMeanNearLayout)
		UMeanNearSlideTitle = UMeanNearSlide.shapes.title
		UMeanNearSlideTitle.text = "UMeanNear Plot - %s - %s" % (trial,view) #set the title of the geom slides
		UMeanNear_placeholder = UMeanNearSlide.shapes[1]
		UMeanNearImage = "%s_UMeanNear_%s.png" % (trial,view)
		if os.path.isfile("%s/%s/postProcessing/images/%s" % (path,trial,UMeanNearImage)):
			insertUMeanNearImage = UMeanNear_placeholder.insert_picture("%s/%s/postProcessing/images/%s" % (path,trial,UMeanNearImage))
		else:
			pass



#### Ctp Iso SLIDES ####
print("Checking for CTP Iso image existance...")
viewsArray = ['frontLeft','left','bottom','rearLeft','rear']

for trial in caseArray:
	print("		Checking in trial%s" % (trial))
	for view in viewsArray:
		CtpIsoImage = "%s_isoCtp_%s.png" % (trial,view)
		if os.path.isfile("%s/%s/postProcessing/images/%s" % (path,trial,CtpIsoImage)):
			print("			Found %s..." % (CtpIsoImage))
		else:
			print("			Cannot find %s... will skip in report..." % (CtpIsoImage))

print("Making CTPIso slides...")
for view in viewsArray:
	for trial in caseArray:
		CtpIsoLayout = prs.slide_layouts[4]
		CtpIsoSlide = prs.slides.add_slide(CtpIsoLayout)
		CtpIsoSlideTitle = CtpIsoSlide.shapes.title
		CtpIsoSlideTitle.text = "Ctp = 0 Iso Plot - %s - %s" % (trial,view) #set the title of the geom slides
		CtpIso_placeholder = CtpIsoSlide.shapes[1]
		CtpIsoImage = "%s_isoCtp_%s.png" % (trial,view)
		if os.path.isfile("%s/%s/postProcessing/images/%s" % (path,trial,CtpIsoImage)):
			insertCtpIsoImage = CtpIso_placeholder.insert_picture("%s/%s/postProcessing/images/%s" % (path,trial,CtpIsoImage))
		else:
			pass




reportName = '_'.join(caseArray)


print("Saving file to: %s/03_reports/%s_report_%s.pptx" % (casePath,reportName,todays_date))
prs.save("%s/03_reports/%s_report_%s.pptx" % (casePath,reportName,todays_date))


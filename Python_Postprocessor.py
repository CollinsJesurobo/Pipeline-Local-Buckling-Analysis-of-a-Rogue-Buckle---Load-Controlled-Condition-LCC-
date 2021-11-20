##########################################################################################################################################################################################
########################################     Pipeline Local Buckling POSTPROCESSOR    ###############################################################################
########################################     Subject:    Abaqus FEA Postprocessing   ###############################################################################
########################################     Author :    Engr.Jesurobo Collins       #####################################################################################    #################################################################################################
########################################     Project:    Personal project            ##############################################################################################
########################################     Tools used: Python,xlsxriter     ##############################################################################################
########################################     Email:      collins4engr@yahoo.com      ##############################################################################################
#########################################################################################################################################################################################
import sys,os
from abaqus import*
from abaqusConstants import*
from math import*
import xlsxwriter
import glob
  
# CHANGE TO CURRENT WORKING DIRECTORY
os.chdir('C:/temp/LCC')
###CREATE EXCEL WORKBOOK, SHEETS AND ITS PROPERTIES####
execFile = 'Results.xlsx'
workbook = xlsxwriter.Workbook(execFile)
workbook.set_properties({
    'title':    'This is Abaqus postprocessing',
    'subject':  'Pipeline Lateral Buckling Analysis',   
    'author':   'Collins Jesurobo',
    'company':  'Personal Project',
    'comments': 'Created with Python and XlsxWriter'})

# Create a format to use in the merged range.
merge_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'yellow'})

SHEET1 = workbook.add_worksheet('Summary_sheet')
SHEET1.center_horizontally()
SHEET1.fit_to_pages(1, 1)
SHEET1.set_column(0,3,20)
SHEET1.set_column(3,4,30)
SHEET1.set_column(5,7,30)
SHEET1.merge_range('A1:D1', 'SUMMARY - MAXIMUM VALUES OF LCC',merge_format)


SHEET2 = workbook.add_worksheet('Empty')
SHEET2.center_horizontally()
SHEET2.fit_to_pages(1, 1)
SHEET2.set_column(0,2,16)
SHEET2.set_column(3,4,35)
SHEET2.set_column(5,6,16)
SHEET2.merge_range('A1:F1', ' LOCAL BUCKLING RESULTS FOR INSTALLATION CASE WITH EXTERNAL OVERPRESSURE -  LOAD CONTROLLED CONDITION (LCC) ',merge_format)

SHEET3 = workbook.add_worksheet('Hydrotest')
SHEET3.center_horizontally()
SHEET3.fit_to_pages(1, 1)
SHEET3.set_column(0,2,16)
SHEET3.set_column(3,4,35)
SHEET3.set_column(5,6,16)
SHEET3.merge_range('A1:F1', 'LOCAL BUCKLING RESULTS FOR OPERATING CASE WITH INTERNAL OVERPRESSURE-  LOAD CONTROLLED CONDITION (LCC) ',merge_format)

SHEET4 = workbook.add_worksheet('Operating')
SHEET4.center_horizontally()
SHEET4.fit_to_pages(1, 1)
SHEET4.set_column(0,2,16)
SHEET4.set_column(3,4,35)
SHEET4.set_column(5,6,16)
SHEET4.merge_range('A1:F1', 'LOCAL BUCKLING RESULTS FOR HYDROTEST CASE WITH INTERNAL OVERPRESSURE-  LOAD CONTROLLED CONDITION (LCC) ',merge_format)

SHEET5 = workbook.add_worksheet('Shutdown')
SHEET5.center_horizontally()
SHEET5.fit_to_pages(1, 1)
SHEET5.set_column(0,2,16)
SHEET5.set_column(3,4,35)
SHEET5.set_column(5,6,16)
SHEET5.merge_range('A1:F1', 'LOCAL BUCKLING RESULTS FOR HYDROTEST CASE WITH INTERNAL OVERPRESSURE-  LOAD CONTROLLED CONDITION (LCC)',merge_format)
# defines the worksheet formatting (font name, size, cell colour etc.)
format_title = workbook.add_format()
format_title.set_bold('bold')
format_title.set_align('center')
format_title.set_align('vcenter')
format_title.set_bg_color('#F2F2F2')
format_title.set_font_size(10)
format_title.set_font_name('Arial')
format_table_headers = workbook.add_format()
format_table_headers.set_align('center')
format_table_headers.set_align('vcenter')
format_table_headers.set_text_wrap('text_wrap')
format_table_headers.set_bg_color('#F2F2F2')
format_table_headers.set_border()
format_table_headers.set_font_size(10)
format_table_headers.set_font_name('Arial')

###WRITING THE TITLES TO SHEET1,SHEET2###
SHEET1.write_row('B2',['WorstElement@LCC','WorstLoadStep@LCC','LocalBuckling Utilization - LCC'],format_title)
SHEET1.write('A3', 'Installation case',format_title)
SHEET1.write('A4', 'Hydrotest',format_title)
SHEET1.write('A5', 'Operating',format_title)
SHEET1.write('A5', 'Shutdown',format_title)


SHEET2.write_row('A2',['Element','LoadStep','Distance (m)','Design Effective Axial Force - Ssd(KN)','Design Moment -Msd (kNm)','LCC'],format_title)
SHEET3.write_row('A2',['Element','LoadStep','Distance (m)','Design Effective Axial Force - Ssd(KN)','Design Moment -Msd (kNm)','LCC'],format_title)
SHEET4.write_row('A2',['Element','LoadStep','Distance (m)','Design Effective Axial Force - Ssd(KN)','Design Moment -Msd (kNm)','LCC'],format_title)
SHEET5.write_row('A2',['Element','LoadStep','Distance (m)','Design Effective Axial Force - Ssd(KN)','Design Moment -Msd (kNm)','LCC'],format_title)
###PIPE INPUT PARAMETERS####
D=0.3667      #Outside Diameter, m
t1=0.0243  #Nominal thickness,m
tcorr=0.003   #Corrosion allowance,m
g=9.81        #Acceleration due to gravity,m/s2
v=0.3         #Poisson ration
fo=0.01       #Out of roundness
alfafab=1     #Fabrication factor
Tdes = 92     #Design temperature,oC
To = 4        #seawater temperature,oC
   
###DNVGLSTF101 LRFD INPUT DATA####
yc=1.00         #Condition load effect factor
yc_hydro=0.93   #Condition load effect factor hydrotest
ym=1.15         #Material resistance factor
yF=1.1          #Functional load effect factor
ySC_inst=1.04   #Low safety class for installation case
ySC =1.14       #medium safety class for operating case
alfaU=0.96      #Material strength factor 
alfaU_hydro=1   #Material strength factor hydrotest                                                
SMYS=450000000  #Specified mimimum yield strength,Pa
SMTS= 535000000 #Specified mimimum tensile strength,Pa
E=207*10**9     #Young modulus,pa
###DENSITY OF SEA WATER & CONTENT#####
rho_sw=1025	#Seawater density,kg/m^3 
rho_inst=0	#Content density during installation,kg/m^3
rho_hydro=1025	#Content density during hydrotest,kg/m^3               
rho_op=700	#Content density during operation,kg/m^3

###WATER DEPTH#####
WD=1600		#water depth,m

###PRESSURE INPUT DATA#####
p_inst=0           #Internal pressure installation,Pa
p_hydro=25800000   #Internal pressure hydrotest,Pa                                                        
p_design=21500000  #Internal pressure design @msl,Pa
###PRESSURE OUTPUT DEFINITIONS#####
Pext=(rho_sw*g*WD)		      #External hydrostatic pressure,Pa
Pint_inst=p_inst+(rho_sw*g*WD)        #Internal pressure@installation 
Pint_hydro=(p_hydro)+(rho_hydro*g*WD) #System test pressure @ seabed
Pint=(p_design)+(rho_op*g*WD)         #Design pressure @ seabed
p_min=0                               #Minimum internal pressure that can be sustained

###PIPE OUPUT PARAMETERS####
R=D/2
t2=t1-tcorr
d_nom=D-(2*t1)
d_op=D-(2*t2)
A_nom=(pi/4)*((D**2)-(d_nom**2))
A_op=(pi/4)*((D**2)-(d_op**2))
z_nom=pi*((D**4)-(d_nom**4))*(32*D)**-1
z_op=pi*((D**4)-(d_op**4))*(32*D)**-1
J_nom=pi*((D**4)-(d_nom**4))*(32)**-1
J_op=pi*((D**4)-(d_op**4))*(32)**-1
###LBUC OUTPUT DATA DEFINITIONS FOR INTERNAL OVERPRESSURE####
def fytemp(Tdes):       #Derating value due to temperature, Pa
        if Tdes < 50:
                fytemp = 0
        elif 50<Tdes<100:
                fytemp = (Tdes-50)*(30/50)
        else:
               fytemp=30+(Tdes-100)*(40/100)      
        return fytemp      
                
def beta(t): 
	if (D/t) <15:                           
		beta= 0.5  
	elif 15<=(D/t)<=60:
		beta=(60-(D/t))/90
	elif (D/t)>60:
		beta=0 
	return beta

def fy(alfaU):	     #Derated Ultimate yield strength   
	fy = (SMYS - fytemp(Tdes))*alfaU
	return fy

def fu(alfaU):       #Derated Tensile yield strength
	fu = (SMTS - fytemp(Tdes))*alfaU
	return fu

def Mp(t,alfaU):     #Plastic capacity Moment
	Mp = fy(alfaU)*((D-t)**2)*t
        return Mp

def Sp(t,alfaU):     #Plastic capacity Force
	Sp = fy(alfaU) *pi * (D-t)*t                                 
        return Sp

def alfac(t,alfaU):  #Flow stress parameter
	alfac =(1-beta(t))+ (beta(t) * (float(fu(alfaU))/fy(alfaU)))
	return alfac

def fcb(alfaU):
	fcb = min(fy(alfaU),fu(alfaU)/1.15)
	return fcb

def Pb(t,alfaU):              #Burst pressure
	Pb =((2*t/(D-t))*(fcb(alfaU))*2/sqrt(3))                                          
	return Pb	   

def yp(t):
        if (Pint-Pext)/(Pb(t,alfaU))<2.0/3:
		alfaP =1-beta(t)
        else:
		alfaP =1-((3*beta(t)*(1-(Pint-Pext)))/(Pb(t,alfaU)))
	return alfaP

###LRFD OUTPUT DATA DEFINITIONS FOR EXTERNAL OVERPRESSURE####

def Pel(t):              #Elastic collapse pressure 
	Pel=(2*E*(t/D)**3)/(1-v**2)
	return Pel

def Pp(t):               #Plastic collapse pressure                               
	Pp= fy(alfaU)*(alfafab)*(2*t/D)
	return Pp

def Pc(t):	        #Characteristic  collapse pressure
	import numpy as np
	Poly=[1.0, -Pel(t),-(Pp(t)**2 +\
	(Pel(t)*Pp(t)* fo*D/t)),(Pel(t)*Pp(t)**2)]
	Par=np.roots(Poly)
	Pc=min(abs(Par))
        return Pc

###LOOP THROUGH THE ODB AND EXTRACT RESULTS SPECIFIED NODESETS FOR ALL STEPS###
def output1():
        row=1
        col=0
        for i in glob.glob('*.odb'):     # loop  to access all odbs in the folder
                odb = session.openOdb(i) # open each odb
                step = odb.steps.keys()  # probe the content of the steps object in odb, steps object is a dictionary, so extract the step names with keys()
                section = odb.rootAssembly.instances['PART-1-1'].nodeSets['NODES']         # extract section for pipeline nodeset
                ###DEFINE RESULT OUTPUT####
                coor = odb.steps[step[0]].frames[-1].fieldOutputs['COORD'].getSubset(region=section).values # results for x coordinate
                for kp in coor:
                        coor1 = kp.data[0]                                                   # extract kp distance
                        ### WRITE OUT MAIN RESULT OUTPUT####    
                        SHEET2.write(row+1,col+2,round(coor1,0),format_table_headers)        # write distance KP to sheet2
                        SHEET3.write(row+1,col+2,round(coor1,0),format_table_headers)        # write distance to sheet3
                        SHEET4.write(row+1,col+2,round(coor1,0),format_table_headers)        # write distance to sheet4
                        SHEET5.write(row+1,col+2,round(coor1,0),format_table_headers)        # write distance to sheet5
                        row+=1                               
output1()
        
###LOOP THROUGH THE ODB AND EXTRACT RESULTS FOR SPECIFIED ELEMENTSETS FOR EMPTY CONDITION###
def output2():
        row=1
        col=0
        for i in glob.glob('*.odb'):     
                odb = session.openOdb(i) 
                step = odb.steps.keys()    
                section = odb.rootAssembly.instances['PART-1-1'].elementSets['ELEM']    
                ###DEFINE RESULT OUTPUT- CHECK FOR INSTALLATION CASE FOR EXTERNAL OVERPRESSURE####
                ESF = odb.steps[step[0]].frames[-1].fieldOutputs['ESF1'].getSubset(region=section).values # results for Effective axial force
                SM = odb.steps[step[0]].frames[-1].fieldOutputs['SM'].getSubset(region=section).values    # results for section moment
                for force,moment in zip(ESF,SM):
                        ESF1 = force.data                                               #Effective axial force
                        SM2 = moment.data[1]                                            #Section moment in lateral direction
                        Msd=SM2*yF*yc                                                   #Design moment,Nm
                        Ssd=ESF1*yF*yc                                                  #Design Effective Axial Force,N
                        e1 = force.elementLabel                                         #Element label
                        P = (ym* ySC_inst*(Pext-p_min))/Pc(t1)                          #Pressure term &its Load effect factors
                        F =((ym*ySC_inst*Ssd)/(alfac(t1,alfaU)*Sp(t1,alfaU)))**2        #Design Effective Axial Force & its Load effect factors 
			
                        M = (ym*ySC_inst*abs(Msd))/(alfac(t1,alfaU)*Mp(t1,alfaU))       #Design moment& its Load effect factors
				
                        LCC =(M +F )**2 + (P )**2
                        ### WRITE OUT MAIN RESULT OUTPUT####
                        SHEET2.write(row+1,col,e1,format_table_headers)                 # write all element in the pipeline to sheet2
                        SHEET2.write(row+1,col+1,step[0],format_table_headers)
                        SHEET2.write(row+1,col+3,Msd/1000,format_table_headers)         #Design moment,kNm
			SHEET2.write(row+1,col+4,Ssd/1000,format_table_headers)         #Design moment,kNm
                        SHEET2.write(row+1,col+5,LCC,format_table_headers)              #Local buckling utilization check        
                        row+=1
                                             
output2()
                                             
###LOOP THROUGH THE ODB AND EXTRACT RESULTS FOR SPECIFIED ELEMENTSETS FOR HYDROTEST CONDITION###
def output3():
        row=1
        col=0
        for i in glob.glob('*.odb'):     
                odb = session.openOdb(i) 
                step = odb.steps.keys()  
                section = odb.rootAssembly.instances['PART-1-1'].elementSets['ELEM']    
                ###DEFINE RESULT OUTPUT- CHECK FOR HYDROTEST CASE FOR INTERNAL OVERPRESSURE####
                ESF = odb.steps[step[5]].frames[-1].fieldOutputs['ESF1'].getSubset(region=section).values # results for Effective axial force
                SM = odb.steps[step[5]].frames[-1].fieldOutputs['SM'].getSubset(region=section).values    # results for section moment
                for force,moment in zip(ESF,SM):
                        ESF1 = force.data                                               #Effective axial force
                        SM2 = moment.data[1]                                            #Section moment in lateral direction
                        Msd=SM2*yF*yc                                                   #Design moment,Nm
                        Ssd=ESF1*yF*yc                                                  #Design Effective Axial Force,N
                        e1 = force.elementLabel                                         # Element label
                        P = (yp(t1)*(Pint_hydro-Pext))/(alfac(t1,alfaU_hydro)
                                                          *Pb(t1,alfaU_hydro))           #Pressure term &its Load effect factors
				    
                        F =((ym*ySC*Ssd)/(alfac(t1,alfaU_hydro)*Sp(t1,alfaU_hydro)))**2                 #Design Effective Axial Force & its Load effect factors 
				   
                        M = (ym*ySC*abs(Msd))/(alfac(t1,alfaU_hydro)*Mp(t1,alfaU_hydro))                #Design moment& its Load effect factors

                        LCC =(M +F )**2 + (P )**2
                        ### WRITE OUT MAIN RESULT OUTPUT####
                        SHEET3.write(row+1,col,e1,format_table_headers)                 # write all element in the pipeline to sheet2
                        SHEET3.write(row+1,col+1,step[5],format_table_headers)
                        SHEET3.write(row+1,col+3,Msd/1000,format_table_headers)         #Design moment,kNm
			SHEET3.write(row+1,col+4,Ssd/1000,format_table_headers)         #Design moment,kNm
                        SHEET3.write(row+1,col+5,LCC,format_table_headers)              #Local buckling utilization check        
                        row+=1      
output3()
###LOOP THROUGH THE ODB AND EXTRACT RESULTS FOR SPECIFIED ELEMENTSETS FOR OPERATING CONDITION###
def output4():
        row=1
        col=0
        for i in glob.glob('*.odb'):     
                odb = session.openOdb(i) 
                step = odb.steps.keys()    
                section = odb.rootAssembly.instances['PART-1-1'].elementSets['ELEM']    
                ###DEFINE RESULT OUTPUT- CHECK FOR HYDROTEST CASE FOR INTERNAL OVERPRESSURE####
                ESF = odb.steps[step[7]].frames[-1].fieldOutputs['ESF1'].getSubset(region=section).values # results for Effective axial force
                SM = odb.steps[step[7]].frames[-1].fieldOutputs['SM'].getSubset(region=section).values    # results for section moment
                for force,moment in zip(ESF,SM):
                        ESF1 = force.data                                               #Effective axial force
                        SM2 = moment.data[1]                                            #Section moment in lateral direction
                        Msd=SM2*yF*yc                                                   #Design moment,Nm
                        Ssd=ESF1*yF*yc                                                  #Design Effective Axial Force,N
                        e1 = force.elementLabel                                         # Element label
                        P = ((yp(t2))*(Pint-Pext))/(alfac(t2,alfaU)*Pb(t2,alfaU))       #Pressure term &its Load effect factors
			       
                        F =((ym*ySC*Ssd)/(alfac(t2,alfaU)*Sp(t2,alfaU)))**2             #Design Effective Axial Force & its Load effect factors 
				   
                        M = (ym*ySC*abs(Msd))/(alfac(t2,alfaU)*Mp(t2,alfaU))            #Design moment& its Load effect factors
				   
                        LCC =(M +F )**2 + (P )**2
                        ### WRITE OUT MAIN RESULT OUTPUT####
                        SHEET4.write(row+1,col,e1,format_table_headers)                 # write all element in the pipeline to sheet2
                        SHEET4.write(row+1,col+1,step[7],format_table_headers)
                        SHEET4.write(row+1,col+3,Msd/1000,format_table_headers)         #Design moment,kNm
			SHEET4.write(row+1,col+4,Ssd/1000,format_table_headers)         #Design moment,kNm
                        SHEET4.write(row+1,col+5,LCC,format_table_headers)              #Local buckling utilization check        
                        row+=1 
                                             
output4()
###LOOP THROUGH THE ODB AND EXTRACT RESULTS FOR SPECIFIED ELEMENTSETS FOR SHUTDOWN CONDITION###
def output5():
        row=1
        col=0
        for i in glob.glob('*.odb'):     
                odb = session.openOdb(i) 
                step = odb.steps.keys()   
                section = odb.rootAssembly.instances['PART-1-1'].elementSets['ELEM']    
                ###DEFINE RESULT OUTPUT- CHECK FOR INSTALLATION CASE FOR SHUTDOWN####
                ESF = odb.steps[step[-1]].frames[-1].fieldOutputs['ESF1'].getSubset(region=section).values # results for Effective axial force
                SM = odb.steps[step[-1]].frames[-1].fieldOutputs['SM'].getSubset(region=section).values    # results for section moment
                for force,moment in zip(ESF,SM):
                        ESF1 = force.data                                               #Effective axial force
                        SM2 = moment.data[1]                                            #Section moment in lateral direction
                        Msd=SM2*yF*yc                                                   #Design moment,Nm
                        Ssd=ESF1*yF*yc                                                  #Design Effective Axial Force,N
                        e1 = force.elementLabel                                         # Element label
                        P = (ym* ySC*(Pext-p_min))/Pc(t2)                               #Pressure term &its Load effect factors
                        F =((ym*ySC*Ssd)/(alfac(t2,alfaU)*Sp(t2,alfaU)))**2             #Design Effective Axial Force & its Load effect factors 
			
                        M = (ym*ySC*abs(Msd))/(alfac(t2,alfaU)*Mp(t2,alfaU))            #Design moment& its Load effect factors
				
                        LCC =(M +F )**2 + (P )**2
                        ### WRITE OUT MAIN RESULT OUTPUT####
                        SHEET5.write(row+1,col,e1,format_table_headers)                 # write all element in the pipeline to sheet2
                        SHEET5.write(row+1,col+1,step[-1],format_table_headers)
                        SHEET5.write(row+1,col+3,Msd/1000,format_table_headers)         #Design moment,kNm
			SHEET5.write(row+1,col+4,Ssd/1000,format_table_headers)         #Design moment,kNm
                        SHEET5.write(row+1,col+5,LCC,format_table_headers)              #Local buckling utilization check        
                        row+=1
                                             
output5()
### WRITE THE MAXIMUM VALUES INTO SUMMARY SHEET(SHEET1)
def output6():
        SHEET1.write('D3', '=ROUND(max(Empty!F3:F200000),2)',format_table_headers)      # localbuckling utilization for empty case
        SHEET1.write('D4', '=ROUND(max(Hydrotest!F3:F200000),2)',format_table_headers)   # localbuckling utilization for Hydrotest case
        SHEET1.write('D5', '=ROUND(max(Operating!F3:F200000),2)',format_table_headers)  # localbuckling utilization for Operating case
        SHEET1.write('D6', '=ROUND(max(Shutdown!F3:F200000),2)',format_table_headers)   # localbuckling utilization for Shutdown case


        ### WRITE WORST LOADSTEP CORRESPONDING TO MAXIMUM LCC VALUES INTO SUMMARY SHEET(SHEET1)
        SHEET1.write('C3','=INDEX(Empty!B3:B200000,MATCH(MAX(Empty!F3:F200000),Empty!F3:F200000,0))',format_table_headers)
        SHEET1.write('C4','=INDEX(Hydrotest!B3:B200000,MATCH(MAX(Hydrotest!F3:F200000),Hydrotest!F3:F200000,0))',format_table_headers)
        SHEET1.write('C5','=INDEX(Operating!B3:B200000,MATCH(MAX(Operating!F3:F200000),Operating!F3:F200000,0))',format_table_headers)
        SHEET1.write('C6','=INDEX(Shutdown!B3:B200000,MATCH(MAX(Shutdown!F3:F200000),Shutdown!F3:F200000,0))',format_table_headers)
        
        ### WORST ELEMENT CORRESPONDING TO MAXIMUM LCC VALUES INTO SUMMARY SHEET(SHEET1)
        SHEET1.write('B3','=INDEX(Empty!A3:A200000,MATCH(MAX(Empty!F3:F200000),Empty!F3:F200000,0))',format_table_headers)
        SHEET1.write('B4','=INDEX(Hydrotest!A3:A200000,MATCH(MAX(Hydrotest!F3:F200000),Hydrotest!F3:F200000,0))',format_table_headers)        
        SHEET1.write('B5','=INDEX(Operating!A3:A200000,MATCH(MAX(Operating!F3:F200000),Operating!F3:F200000,0))',format_table_headers)
        SHEET1.write('B6','=INDEX(Shutdown!A3:A200000,MATCH(MAX(Shutdown!F3:F200000),Shutdown!F3:F200000,0))',format_table_headers)
# PLOT CHARTS
chart1 = workbook.add_chart({'type': 'line'})

# plot lateral dsplacement with kp'''
chart1.set_x_axis({'line':{'none':True}})
chart1.add_series({
        'name': 'Empty',
        'categories':'=Empty!$C$450:$C$2500',           # Distance ,m
        'values': '=Empty!$F$450:$F$2500',              # Local buckling utilization values(empty case)
        'line':{'color':'blue'}})
chart1.add_series({
        'name': 'Hydrotest',
        'categories':'=Hydrotest!$C$450:$C$2500',       # Distance ,m
        'values': '=Hydrotest!$F$450:$F$2500',          # Local buckling utilization values(hydrotest case)
        'line':{'color':'red'}})
chart1.add_series({
        'name': 'Operating',
        'categories':'=Operating!$C$450:$C$2500',       # Distance ,m
        'values': '=Operating!$F$450:$F$2500',          # Local buckling utilization values(operating case)
        'line':{'color':'green'}})
chart1.add_series({
        'name': 'Shutdown',
        'categories':'=Shutdown!$C$450:$C$2500',       # Distance ,m
        'values': '=Shutdown!$F$450:$F$2500',          # Local buckling utilization values(operating case)
        'line':{'color':'magenta'}})
chart1.set_title({'name': 'LCC Plot ',})
chart1.set_x_axis(
        {'name': 'Distance (m)'})
chart1.set_y_axis(
        {'name': 'LCC Utilization'})
chart1.set_style(9)
chart1.set_size({'x_scale': 1.5, 'y_scale': 1.0})



# Insert the chart into the worksheet.
SHEET1.insert_chart('A8', chart1)
output6()

# closes the workbook once all data is written
workbook.close()
# opens the resultant spreadsheet
os.startfile(execFile)
# Lateral Buckling study completed
























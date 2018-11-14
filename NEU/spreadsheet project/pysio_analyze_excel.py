# -*- coding: utf-8 -*-
"""
Created on Fri Jun 13 10:06:52 2014

@author: zsiwy
"""

import xlrd #http://pypi.python.org/pypi/xlrd
import xlwt
import math
from sympy import *

'''Copy paste all raw data from Clampfit into .xls excel file (note the extension!  Not available for office 2013 yet).
When using this script, make sure excel file is closed.'''

# Change this stuff
raw_filename = 'C:/Users/zsiwy/Desktop/Test/laura.xls'  #Excel file path of raw data
save_filename = 'C:/Users/zsiwy/Desktop/Test/lauraOutput.xls' #Excel file path of saved data



# Stuff you don't have to change.  Open raw data file, create new excel workbook, some global variables
wb = xlrd.open_workbook(raw_filename)
new_wb = xlwt.Workbook()

raw_data = []

# functions for analysis.  The meat of the script
def getRawData(sht1, data):
    #copies and outputs raw data
    # sht1: sheets from opened excel workbook/file
    # data: list of data being copied
    rowdata = []
    for numrows in range(sht1.nrows):
        for numcols in range(sht1.ncols):
            rowdata.append(sht1.cell(numrows,numcols).value)
        data.append(rowdata)
        rowdata = []
    return data
    
def calcEffectiveSize(etchRate, minEtched, poreDia, sht1, data):
    # calculate effective particle size    
    # etchRate: either 200 or 1200 nm/min
    # minEtched: minutes etched pore
    # poreDia: pore diameter or constriction point (nm)
    # sht1: sheets from opened excel workbook/file
    # data: list of data being copied
    # see comment below if effective particle size is negative

    poreDiaNM = poreDia*10**-9
    
    if etchRate == 200: 
        poreLenNM = (12000 - (200 / 60 * minEtched)) * 10 ** -9
    elif etchRate == 1200:
        poreLenNM = (12000 - (1200 / 60 * minEtched)) * 10 ** -9
    else:
        print 'pore length default to 11000 nm'
        poreLenNM = 11000 * 10 ** -9

    headers = ['', '', 'Empty Current', 'Particle Current', '(Ie-Ip)/Ip','Pore Length', 'Pore Diameter', 'Effective Particle Size']    

    for i in range(len(headers)):
        data[0].append(headers[i])
    
    for i in range(1, sht1.nrows):
        data[i].append('')
        data[i].append('')
        data[i].append(abs(sht1.cell(i, 6).value * 10 ** -12))
        data[i].append(data[i][len(data[i])-1] - sht1.cell(i, 7).value * 10 ** -12)  #if effective particle size is negative, switch between subtraction and addition.  Also double check if shoudl be using peak or antipeak
        data[i].append((data[i][len(data[i])-2] - data[i][len(data[i])-1]) / data[i][len(data[i])-1])
        data[i].append(poreLenNM)
        data[i].append(poreDiaNM)
        data[i].append((data[i][len(data[i])-3] * poreLenNM * (poreDiaNM ** 2) / (data[i][len(data[i])-3] * poreLenNM * 0.8 / poreDiaNM + 1)) ** (1.0 / 3) / 10 ** -9)       
        
    return data

def calcShapeFactorSphere(particleDia, poreDia, sht1, data):
    # Calculate shape factor as seen in delR/R = f * v / V
    # where delR/R is fractional resistance change
    # v is the volume of the particle
    # V is the volume of the pore
    # f is the shape factor, what this function finds.  
    # for spheres, f is 3/2

    # shape: currently, particle shapes are 'sphere' or 'cylindrical'
    # particleDia:  particle diameter (nm)
    # poreDia: pore diameter (nm)
    # sht1: sheets from opened excel workbook/file
    # data: list of data being copied

    headers = ['', '', 'Empty Current', 'Particle Current', '(Ie-Ip)/Ip', 'Pore Volume', 'Particle Volume', 'Calculated Shape Factor']
    
    for i in range(len(headers)):
        data[0].append(headers[i])

    for i in range(1, sht1.nrows):        
        data[i].append('')
        data[i].append('')
        data[i].append(abs(sht1.cell(i, 6).value * 10 ** -12))
        data[i].append(data[i][len(data[i])-1] + sht1.cell(i, 7).value * 10 ** -12)  # Double check if should be using peak (i, 7) or antipea(i, 10)k
        data[i].append((data[i][len(data[i])-2] - data[i][len(data[i])-1]) / data[i][len(data[i])-1])
        data[i].append(getCylinderVolume(poreDia, 12000)) # pore length default 12000 nm
        data[i].append(getSphereVolume(particleDia))
        data[i].append(data[i][len(data[i])-3] * data[i][len(data[i]) - 2] / data[i][len(data[i]) - 1])
    return data

def calcShapeFactorCylinder(particleDia, particleLen, poreDia, fper, fpar, sht1, data):
    # Calculate shape factor as seen in delR/R = f * v / V
    # where delR/R is fractional resistance change
    # v is the volume of the particle
    # V is the volume of the pore
    # f is the shape factor, what this function finds.  
    # for spheres, f is 3/2

    # shape: currently, particle shapes are 'sphere' or 'cylindrical'
    # particleDia:  particle diameter (nm)
    # poreDia: pore diameter (nm)
    # fper: 
    # fpar: 
    # sht1: sheets from opened excel workbook/file
    # data: list of data being copied

    headers = ['', '', 'Empty Current', 'Particle Current', '(Ie-Ip)/Ip', 'Pore Volume', 'Particle Volume', '', 'cos^2(a)','a']
    
    for i in range(len(headers)):
        data[0].append(headers[i])

    for i in range(1, sht1.nrows):        
        data[i].append('')
        data[i].append('')
        data[i].append(abs(sht1.cell(i, 6).value * 10 ** -12))
        data[i].append(data[i][len(data[i])-1] - sht1.cell(i, 7).value * 10 ** -12)  # Double check if should be using peak (i, 7) or antipea(i, 10)k
        data[i].append((data[i][len(data[i])-2] - data[i][len(data[i])-1]) / data[i][len(data[i])-1])
        data[i].append(getCylinderVolume(poreDia, 12000)) # pore length default 12000 nm
        data[i].append(getCylinderVolume(particleDia, particleLen))
        data[i].append(data[i][len(data[i])-3] * data[i][len(data[i]) - 2] / data[i][len(data[i]) - 1])
        data[i].append((data[i][len(data[i])-1] - fper) / (fpar - fper))
        try:
            data[i].append(2 * math.pi + math.acos(math.sqrt(data[i][len(data[i])-5])))
        except:
            data[i].append('Math domain error')
            
    return data

def matt(particleDia, particleLen, poreDia, fper, fpar, sht1, data):
    # Calculate alpha from matt's MATLAB output
    headers = ['', '', 'Empty Current', 'Particle Current', '(Ie-Ip)/Ip', 'Pore Volume', 'Particle Volume', '', 'cos^2(a)','a']
    
    for i in range(len(headers)):
        data[0].append(headers[i])

    for i in range(1, sht1.nrows):        
        data[i].append('')
        data[i].append('')
        Ie = sht1.cell(i, 3).value * 10 ** -12
        Ip = sht1.cell(i, 4).value * 10 ** -12
        left = (Ie - Ip)/Ip
        v = getCylinderVolume(particleDia, particleLen)        
        V = getCylinderVolume(poreDia, 12000)
        a = symbols('a')
        right = (fper + (fpar - fper) * cos(a) * cos(a)) * v/V
        alpha = solve(left - right,a)
        print alpha
#        data[i].append(sht1.cell(i, 3).value * 10 ** -12)
#        data[i].append(sht1.cell(i, 4).value * 10 ** -12)  # Double check if should be using peak (i, 7) or antipea(i, 10)k
#        data[i].append((data[i][len(data[i])-2] - data[i][len(data[i])-1]) / data[i][len(data[i])-1])
#        data[i].append(getCylinderVolume(poreDia, 12000)) # pore length default 12000 nm
#        data[i].append(getCylinderVolume(particleDia, particleLen))
        
        
#        data[i].append(data[i][len(data[i])-3] * data[i][len(data[i]) - 2] / data[i][len(data[i]) - 1])
#        data[i].append((data[i][len(data[i])-1] - fper) / (fpar - fper))
        try:
            data[i].append(alpha)
        except:
            data[i].append('Error')
            
    return data
    
def laura(poreDia1, poreDiaMD, poreDia2, sht1, data):
    # Calculate alpha from matt's MATLAB output
    headers = ['', '', 'cell diameter1', 'cell diameterMD', 'cell diameter2']
    
    poreDiameter1 = poreDia1 * 10 ** -6
    poreDiameterMD = poreDiaMD * 10 ** -6
    poreDiameter2 = poreDia2 * 10** -6
    
    for i in range(len(headers)):
        data[0].append(headers[i])

    for i in range(1, sht1.nrows):        
        data[i].append('')
        data[i].append('')
        
        left = (0.3/(sht1.cell(i,1).value * 10 ** -9)) -(0.3/(sht1.cell(i,0).value * 10 ** -9))
        d1 = symbols('d1')
        right = (4 * d1**3 * (1 / (1 -0.8 * (d1/poreDiameter1)**3))) / (3.14 * (poreDiameter1 ** 4) * 1.228)
        alpha1 = solve(left - right, d1)

        try:
            data[i].append(float(alpha1[0]))
        except:
            data[i].append('Error')
            
        left = (0.3/(sht1.cell(i,2).value * 10 ** -9)) -(0.3/(sht1.cell(i,0).value * 10 ** -9))
        dMD = symbols('dMD')
        right = (4 * dMD**3 * (1 / (1 -0.8 * (dMD/poreDiameterMD)**3))) / (3.14 * (poreDiameterMD ** 4) * 1.228)
        alphaMD = solve(left - right, dMD)

        try:
            data[i].append(float(alphaMD[0]))
        except:
            data[i].append('Error')
            
        left = (0.3/(sht1.cell(i,3).value * 10 ** -9)) -(0.3/(sht1.cell(i,0).value * 10 ** -9))
        d2 = symbols('d2')
        right = (4 * d2**3 * (1 / (1 -0.8 * (d2/poreDiameter2)**3))) / (3.14 * (poreDiameter2 ** 4) * 1.228)
        alpha2 = solve(left - right, d2)

        try:
            data[i].append(float(alpha2[0]))
        except:
            data[i].append('Error')
            
    return data

def getSphereVolume(dia):
    # returns the volume of a sphere given the diameter
    return 4.0 /3 * math.pi * (dia / 2) ** 3

def getCylinderVolume(dia, length):
    #returns the volume of a cylinder given the diameter and length
    return math.pi * length * (dia / 2) ** 2



# Change - insert the functions (see above) you want.  Goes through each sheet and applies data analysis, saves into new Excel file
# Maybe next version, analyze data into same file?

for ind in range(wb.nsheets):
    # Don't change.  This reads your notebook sheet by sheet and creates a new sheet in the notebook it saves to
    sheet1 = wb.sheet_by_index(ind)
    sheet2 = new_wb.add_sheet(sheet1.name)
    data = []
    
    # Don't change.  Functions depends on reading raw data
    getRawData(sheet1, data)
    
    # Change these variables.  For calculating effective size of particles
    etchRate = 200
    minEtched = 240
    poreDiameter = 920
#    calcEffectiveSize(etchRate, minEtched, poreDiameter, sheet1, data)
    
    # Change these values.  For calcultaing shape factor of sphere (Maxwell's expression)
    particleDiameter= 280
    poreDiameter = 920
#    calcShapeFactorSphere(particleDiameter, poreDiameter, sheet1, data)
    
    # Change these values.  For calculating angle of rod in pore
    particleDiameter = 229
    particleLength = 592
    poreDiameter = 1100
    fperpendicular = 1.75 #1.75
    fparallel = 1.17 #1.17
#    matt(particleDiameter, particleLength, poreDiameter, fperpendicular, fparallel, sheet1, data)
 #   calcShapeFactorCylinder(particleDiameter, particleLength, poreDiameter, fperpendicular, fparallel, sheet1, data)

# Change these values.  For calculating particle diameter - laura's pore
    poreDiameter1 = 24.25
    poreDiameterMiddle = 25.5
    poreDiameter2 = 24.25
    laura(poreDiameter1, poreDiameterMiddle, poreDiameter2, sheet1, data)    
    
    # Don't change.  This writes everything to your current sheet
    for j in range(len(data)):
        for index, value in enumerate(data[j]):
            sheet2.write(j, index, value)
    #data = [] #reset data list
    print data
new_wb.save(save_filename)






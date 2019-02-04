# -*- coding: utf-8 -*-
"""
Created on Thu Jun 14 14:16:49 2018

@author: eliza
"""

from gurobipy import *

import xlrd


#Import distance between districts from excel file (originally from Eugene Lykhovyd, www.lykhovyd.com)

file_location = ("C:\\Users\\eliza\\Desktop\\Wentz 2018-2019\\OK_distances_01172019.xlsx")  

wb = xlrd.open_workbook(file_location) #open workbook
sheet = wb.sheet_by_index(0) #indexing of sheet
rows = sheet.nrows #n rows in sheet
columns = sheet.ncols #n rows in sheet
distance = [] #a list for distance matrix
for i in range(1, rows): #i is the row number
    tmpList = [] #temporary list
    for j in range(1,columns): #j is column number
        tmpList.append(sheet.cell_value(i,j)) #add the value of cell (i,j) to tmpList
    distance.append(tmpList) #add tmpList to distance list
    tmpList.clear #clear tmpList

#Import the population of Oklahoma counties from excel file (originally from Eugene Lykhovyd, www.lyhovyd.com)

file_location = ("C:\\Users\\eliza\\Desktop\\Wentz 2018-2019\\OK-Population-Counties.xls")  

wb = xlrd.open_workbook(file_location) #open workbook
sheet = wb.sheet_by_index(0) #indexing of sheet
rows = sheet.nrows #n rows in sheet
columns = sheet.ncols #n cols in sheet
population = [] #a list for population matrix
for i in range(1, rows): #i is the row
        population.append(sheet.cell_value(i, 1)) #add the value of cell (i,j) to tmpList

#Lower bound of counties/population units allowed in district
LB = 725000

#Upper bound of counties/population units allowed in district
UB = 775000

#K = number of districts
K = 5

#Range of vertices. The vertices are 
vertices = range(len(population))

# Model
m = Model("districting")

#Decision variable to decide if population unit i is assigned to region j
Z = m.addVars(vertices, 
              vertices, 
              vtype=GRB.BINARY, 
              obj = distance,
              name = "assign") 
 
#The objective is to minimize the distance between all 
#population units and their district centers 

#adding objective function

expr = LinExpr()
    
for i in vertices:
    for j in vertices:
        expr = expr + distance[i][j]**2*population[i]*Z[i,j]

m.setObjective(expr, GRB.MINIMIZE)
 
#Constraints 
 
#Each population unit is assigned to 1 district 
#(2) for all i, the sum of Zij over j = 1.

#Zij = 1 when unit i is assigned to unit j 
 
m.addConstrs(Z.sum(i, '*') == 1 for i in vertices)
 
#(3) sum of Zii = k 
#The required number of political districts is k 
#equal to the sum of each district assigned to itself 

expr = LinExpr()
for j in vertices:
    expr = expr + Z[j,j]
    
m.addConstr((expr == K), "3")

#(4) Population boundaries: minimum and maximum population allowed in each district

for j in vertices:
    expr1 = LinExpr()
    for i in vertices:
        expr1 = expr1 + population[i] * Z[i,j]    
    m.addConstr(expr1 <= UB * Z[j,j] )
    print(expr1)

for j in vertices:
    expr1 = LinExpr()
    for i in vertices:
        expr1 = expr1 + population[i] * Z[i,j]
    m.addConstr(expr1 >= LB * Z[j,j])
    
#(5) The number of assigned units is less than the
#number of districts

for i in vertices:
    for j in vertices:
        if (i != j):
            m.addConstr(Z[i,j] <= Z[j,j])

m.write('districtingprogram.lp')

#Optimize
m.optimize()

#Results: which districts are assigned to each district center?

for i in vertices:
    for j in vertices:
        if Z[i,j].x > 0.5:
            print('assign %g to %s' % \
            (i, j))

# -*- coding: utf-8 -*-
"""
Created on Thu Jun 14 14:16:49 2018

@author: eliza
"""
#The purpose of this code is to determine the optimal counties to be political district centers in Oklahoma.
#Using an objective function with population of counties and distance between counties as inputs.
#Require inputs: excel file containing distances between Oklahoma counties, excel file containing the populations of 
#Oklahoma counties, upper and lower population bounds for political districts, and the number of political districts
#in Oklahoma (K), which is 5.

from gurobipy import Model, LinExpr, GRB

import xlrd

#import arcpy


def optimize_district(distance_location, population_location, upper_bound, lower_bound, number_of_districts):
        
    
    #Import distance between districts from excel file (originally from Eugene Lykhovyd, www.lykhovyd.com)
    
    #file_location = ("C:\\Users\\eliza\\Desktop\\Wentz 2018-2019\\OK_distances_01172019.xlsx")  
    
    wb = xlrd.open_workbook(distance_location) #open workbook
    sheet = wb.sheet_by_index(0) #indexing of sheet
    rows = sheet.nrows #n rows in sheet
    columns = sheet.ncols #n cols in sheet
    distance = [] #a list for distance matrix
    for i in range(1, rows): #i is the row number
        tmpList = [] #temporary list
        for j in range(1,columns): #j is column number
            tmpList.append(sheet.cell_value(i,j)) #add the value of cell (i,j) to tmpList
        distance.append(tmpList) #add tmpList to distance list
        tmpList.clear #clear tmpList
    
    #Import the population of Oklahoma counties from excel file (originally from Eugene Lykhovyd, www.lyhovyd.com)
    
    #file_location = ("C:\\Users\\eliza\\Desktop\\Wentz 2018-2019\\OK-Population-Counties.xls")  
    
    wb = xlrd.open_workbook(population_location) #open workbook
    sheet = wb.sheet_by_index(0) #indexing of sheet
    rows = sheet.nrows #n rows in sheet
    columns = sheet.ncols #n cols in sheet
    population = [] #a list for population matrix
    for i in range(1, rows): #i is the row
            population.append(sheet.cell_value(i, 1)) #add the value of cell (i,j) to tmpList
    
    #Lower bound of counties/population units allowed in district
    LB = lower_bound #705253.988
    
    #Upper bound of counties/population units allowed in district
    UB = upper_bound #795286
    
    #K = number of districts
    K = number_of_districts #5
    
    #The vertices are individual counties (population units)
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
            expr += distance[i][j]**2*population[i]*Z[i,j]
    
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
        expr += Z[j,j]
        
    m.addConstr((expr == K), "3")
    
    #(4) Population bounds: minimum and maximum population allowed in each district
    
    for j in vertices:
        expr1 = LinExpr()
        for i in vertices:
            expr1 += population[i] * Z[i,j]    
        m.addConstr(expr1 <= UB * Z[j,j] )
    
    for j in vertices:
        expr1 = LinExpr()
        for i in vertices:
            expr1 += population[i] * Z[i,j]
        m.addConstr(expr1 >= LB * Z[j,j])
        
    #(5) The number of assigned units is less than the
    #number of districts
    
    for i in vertices:
        for j in vertices:
            if (i != j):
                m.addConstr(Z[i,j] <= Z[j,j])
    
    #m.write('districtingprogram.lp')
    
    #Optimize
    m.optimize()
    
    #Results: which districts are assigned to each district center?
    result = []
    for i in vertices:
        for j in vertices:
            if Z[i,j].x > 0.5:
                result.append('assign %g to %s' % \
                (i, j))
                
    return result

optimized = optimize_district("C:\\Users\\eliza\\Desktop\\Wentz 2018-2019\\OK_distances_01172019.xlsx", 
                              "C:\\Users\\eliza\\Desktop\\Wentz 2018-2019\\OK-Population-Counties.xls",
                              795286, 705253.988, 5)

print(optimized)
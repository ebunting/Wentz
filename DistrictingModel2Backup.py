# -*- coding: utf-8 -*-
"""
Created on Thu Jun 14 14:16:49 2018
Last edit Feb 27 2019

@author: eliza
"""
#The purpose of this code is to determine the optimal counties to be political district centers in Oklahoma,
#And assigns each county to its optimal district center based on population and distance, resulting
#in optimally compact political districts for Oklahoma
#Using an objective function with population of counties and distance between counties as inputs.
#Require inputs: excel file containing distances between Oklahoma counties, excel file containing the populations of 
#Oklahoma counties, upper and lower population bounds for political districts, and the number of political districts
#in Oklahoma (K), which is 5.

from gurobipy import Model, LinExpr, GRB

import xlrd

def build_and_solve_hess_model(distance_location, population_location, upper_bound, lower_bound, number_of_districts):
         
    #Import distance between districts from excel file (originally from Eugene Lykhovyd, www.lykhovyd.com) 
    distance = read_data(distance_location) #a list for distance matrix
    
    #Import the population of Oklahoma counties from excel file (originally from Eugene Lykhovyd, www.lykhovyd.com) 
    population = read_data(population_location)
   
    m, Z = build_model(lower_bound, upper_bound, number_of_districts, distance,
                       population, model_name='District Model')
    
    #Optimize
    m.optimize()
    
    #Results: which districts are assigned to each district center?
    result = []
    for i in range(len(population)):
        for j in range(len(population)):
            if Z[i,j].x > 0.5:
                result.append('assign %g to %s' % \
                (i, j))
                
    return result

def read_data(filename):
    wb = xlrd.open_workbook(filename) #open workbook
    sheet = wb.sheet_by_index(0) #indexing of sheet
    rows = sheet.nrows #n rows in sheet
    columns = sheet.ncols #n cols in sheet
    result = [] #a list for distance matrix
    for i in range(1, rows): #i is the row number
        tmpList = [] #temporary list
        for j in range(1,columns): #j is column number
            val = sheet.cell_value(i, j)
            if val != '':
                tmpList.append(val) #add the value of cell (i,j) to tmpList
        if len(tmpList) > 1:
            result.append(tmpList) #add tmpList to distance list
        else:
            result.append(tmpList[0])
        tmpList.clear #clear tmpList
    return result

def build_model(lower_bound, upper_bound, number_of_districts, obj_distance,
                obj_population, model_name='district assignments'):
    
    population = obj_population
    distance = obj_distance
    
    #Lower bound of counties/population units allowed in district
    LB = lower_bound
    
    #Upper bound of counties/population units allowed in district
    UB = upper_bound
    
    #K = number of districts
    K = number_of_districts
    
    #The vertices are individual counties (population units)
    vertices = range(len(population))
    
    # Model
    m = Model(model_name)
    
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
    
    return m, Z

optimized = build_and_solve_hess_model("C:\\Users\\eliza\\Desktop\\Wentz 2018-2019\\OK_distances_01172019.xlsx", 
                              "C:\\Users\\eliza\\Desktop\\Wentz 2018-2019\\OK-Population-Counties.xls",
                              795286, 705254, 5)

print(optimized)
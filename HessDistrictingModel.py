# -*- coding: utf-8 -*-
"""
Created on Thu Jun 14 14:16:49 2018
Last edit Feb 27 2019

@author: eliza

The purpose of this code is to determine the optimal counties to be political district centers in Oklahoma,
And assigns each county to its optimal district center based on population and distance, resulting
in optimally compact political districts for Oklahoma
Using an objective function with population of counties and distance between counties as inputs.
Require inputs: excel file containing distances between Oklahoma counties, excel file containing the populations of 
Oklahoma counties, upper and lower population bounds for political districts, and the number of political districts
in Oklahoma (K), which is 5.
"""

from gurobipy import Model, LinExpr, GRB

import xlrd, xlwt

def build_and_solve_hess_model(distance_filename, population_filename, L, U, K):
    """
    
    :param distance_filename: (list) the distance between each county
    :param population_filename: (list) the population of each county
    :param L: (float) lower bound of counties allowed in a district
    :param U: (float) upper bound of counties allowed in a district
    :param K: (int) number of districts
    :return: (tuple) returns model and decision variable
    """
    #Import distance between districts from excel file (originally from Eugene Lykhovyd, www.lykhovyd.com) 
    distance = read_distance_data(distance_filename) #a list for distance matrix
    
    #Import the population of Oklahoma counties from excel file (originally from Eugene Lykhovyd, www.lykhovyd.com) 
    population = read_population_data(population_filename)
   
    m, Z = build_model(L, U, K, distance, population, model_name='District Model')
    
    #Optimize
    m.optimize()
    
    #Results: which districts are assigned to each district center?
    HessSolution = []
    if m.status == GRB.OPTIMAL:
        for i in range(len(population)):
            for j in range(len(population)):
                if Z[i,j].x > 0.5:
                    HessSolution.append((i, j))
        return HessSolution
    else:
        return "Model is infeasible"

def HessSolution_Excel(district_assignments, file_name='Hess Solution'):
    
    book = xlwt.Workbook()
    sheet1 = book.add_sheet('Hess Solutions')
    sheet1.write(0, 0, 'County')
    sheet1.write(0, 1, 'District')
    for row, tup in enumerate(district_assignments):
        for col, val in enumerate(tup):
            sheet1.write(row + 1, col, val)
    book.save(file_name + '.xls')

def read_distance_data(filename):
    """
    
    :param filename: (string) name of Excel file to open
    :return: (list) Excel sheet 1 as a 2D matrix (or 1D if applicable)
    """
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


def read_population_data(filename):
    """
    
    :param filename: (string) name of Excel file to open
    :return: (list) Excel sheet 1 as a 2D matrix (or 1D if applicable)
    """
    wb = xlrd.open_workbook(filename) #open workbook
    sheet = wb.sheet_by_index(0) #indexing of sheet
    rows = sheet.nrows #n rows in sheet
    result = [] #a list for distance matrix
    for i in range(1, rows): #i is the row number
        result.append(sheet.cell_value(i, 1))
    return result


def build_model(L, U, K, distance, population, model_name='district assignments'):
    """
    
    :param L: (float) lower bound of counties allowed in a district
    :param U: (float) upper bound of counties allowed in a district
    :param K: (int) number of districts
    :param distance: (list) the distance between each county
    :param population: (list) the population of each county
    :param model_name: (string) optional name for model
    :return: (tuple) returns model and decision variable
    """
  
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
     
    for i in vertices:
        expr = LinExpr()
        for j in vertices:
            expr += Z[i,j]
        m.addConstr(expr == 1)
        
     
    #(3) sum of Zjj = k 
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
        m.addConstr(expr1 <= U * Z[j,j] )
    
    for j in vertices:
        expr1 = LinExpr()
        for i in vertices:
            expr1 += population[i] * Z[i,j]
        m.addConstr(expr1 >= L * Z[j,j])
        
    #(5) The number of assigned units is less than the
    #number of districts
    
    for i in vertices:
        for j in vertices:
            if (i != j):
                m.addConstr(Z[i,j] <= Z[j,j])
    
    return m, Z


def print_to_string(district_assignment):
    for tup in district_assignment:
        print('assign %g to %s' % tup)


if __name__ == '__main__':
    HessSolution = build_and_solve_hess_model("C:\\Users\\eliza\\Desktop\\Wentz 2018-2019\\OK_distances_01172019.xlsx", 
                                  "C:\\Users\\eliza\\Desktop\\Wentz 2018-2019\\OK-Population-Counties.xls",
                                  746519, 754022, 5)
    
    HessSolution_Excel(HessSolution)
    
    print_to_string(HessSolution)
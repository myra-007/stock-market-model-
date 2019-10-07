# -*- coding: utf-8 -*-
"""
Created on Wed Sep 25 22:32:52 2019

@author: Ishita
"""

home = "C:\\Users\\Ishita\\Desktop\\Thesis2\\Raw Data\\Before merger\\Daily\\Public sector banks\\"
global file_col_length
def count_null(list1, num):
    count = 0
    for element in list1:
        if(element == num):
            count = count + 1
    print("count of null is: ", count)
    return count

def create_text_file(file_col_write, file_col):
    with open(file_col_write, "w") as f:
        for item in file_col:
            f.write(str(item) + "\n")
            
def create_dataset_file(dataset, file_col):
    with open("dataset.txt", "w") as f:
        f.write(str(file_col))
            
def plot_and_move(file_col, file_col_write, location_save, file_col_copy):
    location_save_plot = home + name_of_file + "_" + file_col_copy + ".png"
    #print("location_save_plot: ", location_save_plot)
    #print(type(x), type(file_col))
    if(j == 1):
        plt.plot(x, file_col, color="red")
    elif(j == 2):
        plt.plot(x, file_col, color="blue")
    elif(j == 3):
        plt.plot(x, file_col, color="yellow")
    elif(j == 4):
        plt.plot(x, file_col, color="green")
    plt.xlabel("Days")
    plt.ylabel("Price")
    print("def")
    plt.savefig(location_save_plot, bbox_inches="tight")
    if os.path.exists(location_save_plot):
        print("File exists")
        if os.path.exists(location_save):
            print("Folder exists")
            try:
                shutil.move(location_save_plot, location_save)
                shutil.move(file_col_write, location_save)
                shutil.move("dataset.txt", location_save)
            except:
                pass
        else:
            os.makedirs(location_save)
            print("Folder created")
            try:
                shutil.move(location_save_plot, location_save)
                shutil.move(file_col_write, location_save)
                shutil.move("dataset.txt", location_save)
            except:
                pass
    plt.show()
   
def create_array(file_col, j, file_col_copy):
      file_col = []
      for i in range(1, sheet.nrows):
          file_col.append(sheet.cell_value(i,j))
      #print(file_col)
      for i in range(count_null(file_col, "null")):
          file_col.remove("null")
      #print(file_col)
      file_col_write = home + file_col_copy + ".txt"
      create_text_file(file_col_write, file_col)
      print("file_col length: ",len(file_col))
      file_col_length = len(file_col)
      create_dataset_file("dataset.txt", file_col_length)
      #print(file_col_length)
      a = (file_col, file_col_length)
      #print("a ", type(a))
      return a
      
import xlrd
import matplotlib.pyplot as plt
import os
import shutil


location = home + "testfile-532276.BO.xlsx"
print(location)
name_of_file = location.split("\\")[-1]
name_of_file = name_of_file.split(".")[0] + name_of_file.split(".")[1]
print(name_of_file)

wb = xlrd.open_workbook(location) 
sheet = wb.sheet_by_index(0) 

print("Rows:", sheet.nrows)
print("Column:", sheet.ncols)
location_save = home + name_of_file

for j in range(1,5):
    file_col = sheet.cell_value(0,j)
    file_col_copy = sheet.cell_value(0,j)
    print(file_col)
    a = create_array(file_col, j, file_col_copy)    
    x = []
    #print("x:", x)
    file_col_write = home + file_col_copy + ".txt"
    file_col = a[0]
    file_col_length = a[1]
    for i in range(0, file_col_length):
        x.append(i)
    print("length of x:", len(x))
    plot_and_move(file_col, file_col_write, location_save, file_col_copy) 

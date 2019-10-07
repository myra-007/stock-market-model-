# -*- coding: utf-8 -*-
"""
Created on Wed Sep 25 16:08:26 2019

@author: Ishita
"""
def count_null(list1, num):
    count = 0
    for element in list1:
        if(element == num):
            count = count + 1
    print("count of null is: ", count)
    return count
                
import xlrd
import matplotlib.pyplot as plt
import os
import shutil

x = []
home = "C:\\Users\\Ishita\\Desktop\\Thesis2\\Raw Data\\Before merger\\Daily\\Public_sector_banks\\"

location = home + "testfile-532276.BO.xlsx"

name_of_file = location.split("\\")[-1]
name_of_file = name_of_file.split(".")[0] + name_of_file.split(".")[1]
print(name_of_file)

wb = xlrd.open_workbook(location) 
sheet = wb.sheet_by_index(0) 

print("Rows:", sheet.nrows)
print("Column:", sheet.ncols)
location_save = home + name_of_file
        
open_1 = []    
for i in range(1, sheet.nrows):
    open_1.append(sheet.cell_value(i,1))
null_count = count_null(open_1, "null")

for i in range(null_count):
    open_1.remove("null")
open_write = home + "open_1" + ".txt"
#print(open_1)
with open(open_write, "w") as f:
    for item in open_1:
        f.write(str(item) + "\n")        
print("Length of array after removing nulls: ", len(open_1))
with open("dataset.txt", "w") as m:
    m.write(str(len(open_1)))
a = len(open_1)-1
for i in range(0,a+1):
    x.append(i)
#print(x)
location_save_plot1 = home + name_of_file + "_open" + ".png"
plt.plot(x, open_1, color="green")
plt.xlabel("Days")
plt.ylabel("Price")
plt.savefig(location_save_plot1, bbox_inches="tight")
if os.path.exists(location_save_plot1):
    print("File exists")
    if os.path.exists(location_save):
        print("Folder exists")
        shutil.move(location_save_plot1, location_save)
        shutil.move(open_write, location_save)
        shutil.move("dataset.txt", location_save)
    else:
        os.makedirs(location_save)
        print("Folder created")
        shutil.move(location_save_plot1, location_save)
        shutil.move(open_write, location_save)
        shutil.move("dataset.txt", location_save)
plt.show()


highest_1 = []
for i in range(1, sheet.nrows):
    highest_1.append(sheet.cell_value(i,2))    
count_null(highest_1, "null")
for i in range(null_count):
    highest_1.remove("null")
highest_write = home + "highest_1" + ".txt"
#print(highest_1)
with open(highest_write, "w") as g:
    for item in highest_1:
        g.write(str(item) + "\n")
print("Length of array after removing nulls: ", len(highest_1))
a = len(highest_1)-1
location_save_highest = home + name_of_file + "_highest"
location_save_plot2 = home + name_of_file + "_highest" + ".png"
plt.plot(x, highest_1, color="blue")
plt.xlabel("Days")
plt.ylabel("Price")
plt.savefig(location_save_plot2, bbox_inches="tight")
if os.path.exists(location_save_plot2):
    print("File exists")
    if os.path.exists(location_save):
        print("Folder exists")
        shutil.move(location_save_plot2, location_save)
        shutil.move(highest_write, location_save)
    else:
        os.makedirs(location_save)
        print("Folder created")
        shutil.move(location_save_plot2, location_save)
        shutil.move(highest_write, location_save)
plt.show()

lowest_1 = []
for i in range(1, sheet.nrows):
    lowest_1.append(sheet.cell_value(i,3))

count_null(lowest_1, "null")

for i in range(null_count):
    lowest_1.remove("null")    
lowest_write = home + "lowest_1" + ".txt"
#print(lowest_1)
with open(lowest_write, "w") as h:
    for item in lowest_1:
        h.write(str(item) + "\n")
print("Length of array after removing nulls: ", len(lowest_1))
a = len(lowest_1)-1
location_save_lowest = home + name_of_file + "_lowest"
location_save_plot3 = location_save_lowest + ".png"
plt.plot(x, lowest_1, color="red")
plt.xlabel("Days")
plt.ylabel("Price")
plt.savefig(location_save_plot3, bbox_inches="tight")
if os.path.exists(location_save_plot3):
    print("File exists")
    if os.path.exists(location_save):
        print("Folder exists")
        shutil.move(location_save_plot3, location_save)
        shutil.move(lowest_write, location_save)
    else:
        os.makedirs(location_save)
        print("Folder created")
        shutil.move(location_save_plot3, location_save)
        shutil.move(lowest_write, location_save)
plt.show()
        
close_1 = []
for i in range(1, sheet.nrows):
    close_1.append(sheet.cell_value(i,4))
count_null(close_1, "null")
for i in range(null_count):
    close_1.remove("null")
close_write = home + "close_1" + ".txt"
#print(close_1)    
with open(close_write, "w") as j:
    for item in close_1:
        j.write(str(item) + "\n")
print("Length of array after removing nulls: ", len(close_1))
a = len(close_1)-1
location_save_close = home + name_of_file + "_close"
location_save_plot4 = location_save_close + ".png"
plt.plot(x, close_1, color="yellow")
plt.xlabel("Days")
plt.ylabel("Price")
plt.savefig(location_save_plot4, bbox_inches="tight")
if os.path.exists(location_save_plot4):
    print("File exists")
    if os.path.exists(location_save):
        print("Folder exists")
        shutil.move(location_save_plot4, location_save)
        shutil.move(close_write, location_save)
    else:
        os.makedirs(location_save)
        print("Folder created")
        shutil.move(location_save_plot4, location_save)
        shutil.move(close_write, location_save)
plt.show()
        
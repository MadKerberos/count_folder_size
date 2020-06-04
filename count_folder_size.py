# !/usr/bin/python3
import os
from pathlib import Path
import xlsxwriter
import functions

folder_root = 'C:\\temp\\'
output_excel = 'C:\\Users\\Giuseppe\\Desktop\\GitHub\\count_folder_size\\'

print(" - - - - - - - Starting Script - - - - - - - -  ")
print("Folder root: " + folder_root)

# (1) Print info about Root Folder

folder_root_size = functions.GetSingleFolderSize(folder_root)
print("Folder Root Size: " + str(functions.GetSingleFolderSize(folder_root)))

# (2) Getting info about Root SubFolders
os.chdir(folder_root) #< ------------------------------ Changing pwd
tree_size_folder_info = functions.GetFolderTreeInfo(folder_root_size)

#(3) Write Excel Output
os.chdir(output_excel) #< ------------------------------ Changing pwd
functions.WriteToExcel(tree_size_folder_info) 
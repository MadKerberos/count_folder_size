# !/usr/bin/python3
import os
from pathlib import Path
import xlsxwriter

def FolderSize(folder_name):
    #os.chdir(folder_name) 
    root_directory = Path('./' + str(folder_name) )
    #print(root_directory)
    size_in_gb = float( format( (sum(f.stat().st_size for f in root_directory.glob('**/*') if f.is_file()) ) / (1024**3), ".2f" ) ) # Folder Size in GB
    return size_in_gb
    #print(size_in_gb)


folder_root = 'C:\\temp\\'
print(FolderSize(folder_root))

# (1) Print info about Root Folder
dir_size_info = {}

dir_size = 0
folder_root_size = FolderSize(folder_root)
print("Folder Root Size: " + str(FolderSize(folder_root)))
#os.chdir(folder_root)

dir_size_info["folder_root"] = folder_root_size

# (2) Print info about Root SubFolders
for folder in os.listdir(folder_root):    
    if os.path.isdir(folder): 
        dir_size += FolderSize(folder)
        dir_size_info[folder] = FolderSize(folder)
        print("\t" + folder + " Size: "  + str(FolderSize(folder)))

other_files_size = folder_root_size - dir_size
dir_size_info["other_files_in_folder"] = float(other_files_size)
print("\tOther Files size in folder: "  + '{:.2f}'.format(other_files_size) )

print(dir_size_info)

# (3) Write to Excel
workbook = xlsxwriter.Workbook('Expenses02.xlsx')
worksheet = workbook.add_worksheet()
# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 40)

# Write some simple text.
worksheet.write('A1', 'Ciao')

workbook.close()
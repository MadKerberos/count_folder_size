from pathlib import Path
import xlsxwriter
import os 

def GetSingleFolderSize(folder_name):
    #os.chdir(folder_name) 
    root_directory = Path('./' + str(folder_name) )
    size_in_gb = float( format( (sum(f.stat().st_size for f in root_directory.glob('**/*') if f.is_file()) ) / (1024**3), ".2f" ) ) # Folder Size in GB
    return size_in_gb

def GetFolderTreeInfo(folder_root_size):
    dir_size_info = {}
    dir_size_info["folder_root"] = folder_root_size
    dir_size = 0
    for folder in os.listdir('.'):    
        if os.path.isdir(folder): 
            dir_size += GetSingleFolderSize(folder)
            dir_size_info[folder] = GetSingleFolderSize(folder)
            print("\t" + folder + " Size: "  + str(GetSingleFolderSize(folder)))

    other_files_size = folder_root_size - dir_size
    dir_size_info["other_files_in_folder"] = float(other_files_size)
    print("\tOther Files size in folder: "  + '{:.2f}'.format(other_files_size) )
    return dir_size_info

def WriteToExcel(dicDirSizeInfo):
    workbook = xlsxwriter.Workbook('Folder_Info.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})

    # Widen the first column to make the text clearer.
    worksheet.set_column('A:A', 30)
    worksheet.set_column('B:B', 30)

    # Write some simple text.
    #worksheet.write('A1', 'Ciao')
    worksheet.write('A1', "Folder" , bold )
    worksheet.write('B1', "Size GB", bold )

    i = 2
    for dir in dicDirSizeInfo:
        #print('A'+str(i) + " " + str(dir) + " " + str(dicDirSizeInfo[dir]))
        worksheet.write('A'+str(i), dir)
        worksheet.write('B'+str(i), dicDirSizeInfo[dir])
        i+=1
        
    workbook.close()
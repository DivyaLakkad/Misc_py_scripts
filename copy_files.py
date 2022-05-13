import os
import shutil

rootdir = r"\\projects2016.graham.ca\sites\e20147\GrahamExecution\21_Accounting\Timecards\0. Data Source Files"
dest_dir = r"C:\Users\divyal\Desktop\projects\NOVA Equip Py\All TS"

for subdir, dirs, files in os.walk(rootdir):
    for file in files:
        if file.find('CV') > 0 and file.endswith('.xlsx'):
            
            shutil.copy(os.path.join(subdir, file), dest_dir)

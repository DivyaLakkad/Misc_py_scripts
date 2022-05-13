import os
import shutil
import datetime

rootdir = r"C:\Users\divyal\Desktop\projects\NOVA Equip Py\TS_By_Week"

for subdir, dirs, files in os.walk(rootdir):
    for file in files:
        if file.find('CV') > 0 and file.endswith('.xlsx'):
            pass
        else:
            os.remove(os.path.join(subdir, file))

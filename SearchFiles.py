import os


wpath = "C:\Samcef\Caesam\StrenBox_V2.4\customer\workspace"
files = os.listdir(wpath)

for file in files:
    ppath = os.path.join(wpath, file)
    hpath = os.path.join(ppath, "help")
    if os.path.isdir(hpath):
        hfiles = os.listdir(hpath)
        for hfile in hfiles:
            if "~$" in hfile:
                hfileos = os.path.join(hpath, hfile) 
                print(hfileos)
                os.remove(hfileos)
    
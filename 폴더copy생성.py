import os

for dirpath, dirnames, filenames in os.walk(r"E:\2020학년도"):

    for dirname in dirnames:
        new_dir = dirpath.replace("2020", "2021")
        new_dirname = os.path.join(new_dir, dirname)
        os.makedirs(new_dirname)
        # print ("\t", dirname, " ", new_dirname)
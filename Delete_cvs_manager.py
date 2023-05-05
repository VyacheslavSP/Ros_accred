import os
import fnmatch


def delete_cvs_manager():
    path_to_delete_cvs = 'C:/Unitess/TEMP/'
    listOfFiles = os.listdir(path_to_delete_cvs)
    pattern = "*.csv"
    for entry in listOfFiles:
        if fnmatch.fnmatch(entry, pattern):
            os.remove(path_to_delete_cvs+entry)

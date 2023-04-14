import os
import shutil


def newest():
    path = 'C:/Users/VecheslavSP/Downloads'
    destination_path = "C:/Users/VecheslavSP/Desktop/Python/Ros_accred/NetWork/for_rosaccredit.txt"
    files = os.listdir(path)
    paths = [os.path.join(path, basename) for basename in files]
    new_location = shutil.move(
        max(paths, key=os.path.getctime), destination_path)
    return new_location


# newest()

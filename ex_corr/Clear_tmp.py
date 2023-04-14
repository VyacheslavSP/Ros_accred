import os

dir = 'C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/tmpXLS'


def clear():
    for f in os.listdir(dir):
        os.remove(os.path.join(dir, f))




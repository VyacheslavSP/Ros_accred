import ClassPoveritel


def Build_poveritel_list():
    list_poveritrl = []
    dict_pov = {}
    f = open('C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/listOfPoveritel.txt',
             'r', encoding='UTF-8')
    l = [line.strip() for line in f]
    for line in l:
        line = line.split(";")
        tmp = ClassPoveritel.Poveriteli(line[0], line[1], line[2], line[3])
        dict_pov[line[1]] = line[2]
        list_poveritrl.append(tmp)
    return list_poveritrl, dict_pov

import os
from ConfigFileExcelBeforeArsh import configFileBeforeArsh
from AddSnils import addSnilsAndDeleteNameSoname
from Check_sender_files import check_send_files_for_Stas_Xls
from uniteTmpFiles import union
from Clear_tmp import clear
from ReBuildNewFileStas import ReBuild_New_stas_file
# import sys
# sys.path.insert(0, "C:/Users/VecheslavSP/Desktop/Python/Ros_accred/NetWork/")
from Selenium_arsh import find_reestr


def Main_New_correct_excel():
    need_action = False
    path_tmp = union()
    path_old_file = 'C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/Выгрузка_v5.4.4.xlsm'
    path_new_file = 'C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/Vigruzka v6.0_standalone.xlsm'
    addSnilsAndDeleteNameSoname(path_tmp)
    configFileBeforeArsh(path_tmp)
    check_send_files_for_Stas_Xls(
        path_old_file, path_tmp)  # проверка старого файла
    if (check_send_files_for_Stas_Xls(path_new_file, path_tmp)):  # проверка нового файла
        find_reestr(path_tmp)
        ReBuild_New_stas_file(path_new_file, path_tmp)
        need_action = True
    else:
        need_action = False
    clear()  # очистка временной папки
    os.remove(path_tmp)
    return need_action
# Main_New_correct_excel()

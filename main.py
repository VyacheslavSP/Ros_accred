

import sys_path
import traceback
import Ex_correct
import Start_export_manager
import Selenium_arsh
import copy_rename_for_accred
import selenium_accred_build
import Build_correct_excel
import Check_sender_files
import Selenium_rosaccredit

# Ex_correct.build_correct_excel(path_of_excell, Ex_correct.build_2d_array())
# print(Ex_correct.check_send_files(
#   path_of_excell, Ex_correct.build_arrey_for_check()))
path_of_excel_Stas = 'C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/Выгрузка_v5.4.4.xlsm'


def full_operation():
    # запуск манагера и первая выгрузга в XLS
    Start_export_manager.main_work_manager()
# Корректировка Экселя (форматы даты, пустота и прочее)
    path_of_excell = Ex_correct.Work_TMP_Excel()
# проверка на уже отправленные файлы
    need_action = Check_sender_files.check_send_files_for_Stas_Xls(
        path_of_excel_Stas, path_of_excell)
    if (need_action):
        # проверка результатов в аршине и корректировка отправлений
        Selenium_arsh.Main_res(path_of_excell)
    # корректировка  экселя данными из аршина
        Ex_correct.write_rows_in_Stas(
            path_of_excel_Stas, path_of_excell, need_action)
        # коррекция фала Стаса
        print("открой файл Стаса")

    else:
        print("ничего не надо")
    return


full_operation()

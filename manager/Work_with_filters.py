import time

import keyboard
import pyautogui


def mouse_pozition(x, y, sleep):
    pyautogui.click(x, y)
    time.sleep(float(sleep))
    return


def many_filters(arr_cont_press_down):
    mouse_pozition(1129, 369, 0.1)  # Фильтр сотрудники
    keyboard.send("Home")   # возврат в начало списка
    for i in range(int(arr_cont_press_down)):
        keyboard.send("Down")
    # финальное нажатие фильтра
    mouse_pozition(797, 819,  0.5)


def once_filters():
    mouse_pozition(789, 654,  0.5)  # по дате регистрации
    # из списка-по дате завершения
    mouse_pozition(845, 735,  0.5)
    mouse_pozition(883, 676,  0.5)  # список за сегодня
    mouse_pozition(814, 789, 1)  # из списка-за месяц
  #  mouse_pozition(860, 728,  0.5)  # из списка в диапазоне
  #  mouse_pozition(921, 680,  0.5)  # первая дата
  #  for i in range(10):
  #      keyboard.send("Delete")
    # временно с 28 апреля планируется за месяц
  #  keyboard.write("28.03.2023")
    time.sleep(0.1)
    mouse_pozition(1078, 728,  0.5)  # список фильтров
    # из списка- фильтр Zсектор
    mouse_pozition(823, 791,  0.5)


# once_filters()
# many_filters([534, 528, 529, 507])

# mouse_pozition(661, 775, 2)  # фильтр
# mouse_pozition(921, 373, 0.5)  # фильтр сектор
# mouse_pozition(833, 415, 0.5)  # из списка сектор-2
#
#
#
# mouse_pozition(814, 789, 1)# из списка-за месяц
# mouse_pozition(832, 763,  0.5)  # из списка-за неделю

# . в текущей версии выбирает с начала месяца до  текущей даты
#
#
# mouse_pozition(906, 766, 1) # из списка- фильтр Тест (только пугачев)
#
#
#
# pyautogui.moveTo(602, 519, 1)

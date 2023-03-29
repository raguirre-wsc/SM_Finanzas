from pywinauto import *
from pywinauto.keyboard import send_keys
import arrow

SAP=Application().start(r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
# SAP=Application().start(r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
# time.sleep(7)
# for x in range(7):
#     send_keys("{DOWN}")
# send_keys("{ENTER}")
# time.sleep(7)
# send_keys("zcl04")
# send_keys("{ENTER}")
# time.sleep(2)
# send_keys("+{F5}")
# time.sleep(2)
# send_keys("pagosposidiari")
# for x in range(4):
#     send_keys("{TAB}")
# send_keys("{DELETE}")
# send_keys("{F8}")
# for x in range(23):
#     send_keys("{TAB}")
# tadd2 = arrow.now().shift(days=2).format("DD.MM.YYYY")
# tsus5 = arrow.now().shift(days=-5).format("DD.MM.YYYY")
# send_keys(tsus5)
# send_keys("{TAB}")
# send_keys(tadd2)
# for x in range(41):
#     send_keys("{TAB}")
# send_keys("{DELETE}")
# send_keys("{F8}")
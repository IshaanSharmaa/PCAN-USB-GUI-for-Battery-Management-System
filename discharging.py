from datetime import datetime
# from typing_extensions import ParamSpecArgs
from PyQt5 import QtCore, QtGui, QtWidgets, QtTest
from PyQt5.QtWidgets import *
from GUI_python import Ui_MainWindow

import pandas as pd
import numpy as np
import random
from datetime import datetime
import sys
import string

import time
from time import sleep
import numpy as np
import pandas as pd
from PCANBasic import *
import ProcessMessageCanFunc as pro
from ctypes import *
from datetime import datetime
from xlsxwriter import *
import openpyxl

dict = {
    0x100: ["Total Voltage", "Current", "Remaining Capacity", 100],
    0x101: ["Full Capacity", "Number Of Cycles", "RSOC", 100],
    # 0x102: [""] protection state later
    0x105: ["NTC1", "NTC2", "NTC3", 100],
    # 0X106: ["NTC4", "NTC5", "NTC6", 100],
    0X107: ["V1", "V2", "V3", 1000],
    0X108: ["V4", "V5", "V6", 1000],
    0X109: ["V7", "V8", "V9", 1000],
    0X10A: ["V10", "V11", "V12", 1000],
    0X10B: ["V13", "V14", "V15", 1000]
}

def convert(s):
    i = int(s, 16)
    cp = pointer(c_int(i))
    # fp = cast(cp, POINTER(c_float))
    return float.fromhex(s)

def convert2(s):
    bits = 16
    val = int(s, bits)
    if val & (1 << (bits - 1)):
        val -= 1 << bits
        return val
    else:
        cp = pointer(c_int(val))
        return cp.contents.value

def filter_out_junk(text):
    return ''.join(x for x in text if x in set(string.printable))


df = pd.DataFrame(columns=[
    'Total Voltage', 'Current', 'Remaining Capacity', 'Full Capacity',
    'Number Of Cycles', 'RSOC', 'NTC1', 'NTC2', 'NTC3', 'NTC4', 'NTC5', 'NTC6',
    'V1', 'V2', 'V3', 'V4', 'V5', 'V6', 'V7', 'V8', 'V9', 'V10', 'V11', 'V12',
    'V13', 'V14', 'V15'
])


class code(object):
    def __init__(self, object):
        self.ui = Ui_MainWindow()
        self.ui.setupUi(object)
        object.show()
        pass

    def start(self):
        self.refresh()
        self.ui.pushButton.setText("Start")
        self.ui.pushButton.setCheckable(True)
        self.ui.pushButton.clicked.connect(self.pcan)

        self.ui.pushButton_2.setCheckable(True)

    def pcan(self):
        #Initialize

        self.ui.pushButton.setText("Stop")

        pcan = PCANBasic()


        while (self.ui.pushButton.isChecked()):

            pcan.Uninitialize(PCAN_USBBUS1)
            initialise = pcan.Initialize(PCAN_USBBUS1, PCAN_BAUD_250K)

            if initialise != PCAN_ERROR_OK:
                l = pcan.GetErrorText(initialise)
                print(l[1])
            else:
                # timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                

                new_dict = {}

                start_time = time.time()
                for can_id, id_list in dict.items():

                    #Writing in Data
                    msg = TPCANMsg()
                    msg.ID = can_id
                    msg.MSGTYPE = PCAN_MESSAGE_STANDARD
                    msg.LEN = 1
                    msg.DATA[0] = 1

                    written_m = pcan.Write(PCAN_USBBUS1, msg)

                    #reading
                    delay = 0.05
                    sleep(delay)
                    read_m = pcan.Read(PCAN_USBBUS1)

                    timestamp, can_data = pro.ProcessMessageCan(
                        read_m[1], read_m[2])

                    print(can_data)
                    l = list(can_data.split())
                    
                    data = {}

                    if can_id == 0x105:
                        data[0] = ((convert(filter_out_junk(l[0] + l[1])) - 2731) / 10) 
                        data[1] = ((convert(filter_out_junk(l[2] + l[3])) - 2731)/ 10) 
                        data[2] = ((convert(filter_out_junk(l[4] + l[5])) - 2731) / 10)    
                    
                    else:

                        data[0] = convert(filter_out_junk(l[0] + l[1])) / id_list[3]
                        if can_id == 0x100:
                            data[1] = convert2(filter_out_junk(l[2] + l[3])) / id_list[3]
                        else:
                            data[1] = convert(filter_out_junk(l[2] + l[3])) / id_list[3]                             
                        data[2] = convert(filter_out_junk(l[4] + l[5])) / id_list[3]

                    for i in range(3):
                        new_dict[id_list[i]] = data[i]

                if self.ui.pushButton_2.isChecked():
                    
                    
                    df.loc[datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S")] = new_dict.values()

                    if d == 0:                   
                        var = str(datetime.now().strftime("%Y-%m-%d_%H-%M-%S"))
                        df.to_excel(f"readings\\test_read_{var}.xlsx",
                                    sheet_name="1",
                                    index=True)

                    elif d == 1:
                        df.to_excel(f"readings\\test_read_{var}.xlsx",
                                    sheet_name="1",
                                    index=True)
                    d = 1
                else:
                    d = 0
                    df = pd.DataFrame(columns=[
                        'Total Voltage', 'Current', 'Remaining Capacity',
                        'Full Capacity', 'Number Of Cycles', 'RSOC', 'NTC1',
                        'NTC2', 'NTC3', 'V1', 'V2', 'V3', 'V4', 'V5', 'V6', 'V7', 'V8', 'V9', 'V10', 'V11',
                        'V12', 'V13', 'V14', 'V15'
                    ])

                self.iteration(new_dict)
                print(new_dict)
                # print("time taken: ", time.time() - start_time)

            pcan.Uninitialize(PCAN_USBBUS1)                                                             

        self.start()

    def iteration(self, df):

        current = df["Current"]

        if current < 0:
            self.ui.label_charging_discharging.setText("DISCHARGING")
            self.ui.label_charging_discharging.adjustSize()
        elif current == 0:
            self.ui.label_charging_discharging.setText("STATIC")
            self.ui.label_charging_discharging.adjustSize()
        elif current > 0:
            self.ui.label_charging_discharging.setText("CHARGING")
            self.ui.label_charging_discharging.adjustSize() 

        self.progressBarValue(df["Total Voltage"],
                              self.ui.circularProgress_voltage,
                              "rgba(85, 255, 127, 255)")
        self.ui.labelPercentage_voltage.setText(str(df["Total Voltage"]))

        self.progressBarValue2(df["Current"], self.ui.circularProgress_current,
                               "rgba(85, 255, 127, 255)")
        self.ui.labelPercentage_current.setText(str(abs(df["Current"])))

        self.progressBarValue(df["RSOC"] * 100, self.ui.circularProgressSOC,
                              "rgba(85, 255, 127, 255)")
        self.ui.labelPercentageSOC.setText(str(df["RSOC"] * 100))

        self.ui.label_2.setText("V1 - " + str(df["V1"]) + "V")
        self.ui.label_2.adjustSize()
        self.progressBarHorizontal(df["V1"], self.ui.progressBar_1)

        self.ui.label_3.setText("V2 - " + str(df["V2"]) + "V")
        self.ui.label_3.adjustSize()
        self.progressBarHorizontal(df["V2"], self.ui.progressBar_2)

        self.ui.label_4.setText("V3 - " + str(df["V3"]) + "V")
        self.ui.label_4.adjustSize()
        self.progressBarHorizontal(df["V3"], self.ui.progressBar_3)

        self.ui.label_5.setText("V4 - " + str(df["V4"]) + "V")
        self.ui.label_5.adjustSize()
        self.progressBarHorizontal(df["V4"], self.ui.progressBar_4)

        self.ui.label_6.setText("V5 - " + str(df["V5"]) + "V")
        self.ui.label_6.adjustSize()
        self.progressBarHorizontal(df["V5"], self.ui.progressBar_5)

        self.ui.label_7.setText("V6 - " + str(df["V6"]) + "V")
        self.ui.label_7.adjustSize()
        self.progressBarHorizontal(df["V6"], self.ui.progressBar_6)

        self.ui.label_8.setText("V7 - " + str(df["V7"]) + "V")
        self.ui.label_8.adjustSize()
        self.progressBarHorizontal(df["V7"], self.ui.progressBar_7)

        self.ui.label_9.setText("V8 - " + str(df["V8"]) + "V")
        self.ui.label_9.adjustSize()
        self.progressBarHorizontal(df["V8"], self.ui.progressBar_8)

        self.ui.label_10.setText("V9 - " + str(df["V9"]) + "V")
        self.ui.label_10.adjustSize()
        self.progressBarHorizontal(df["V9"], self.ui.progressBar_9)

        self.ui.label_11.setText("V10 - " + str(df["V10"]) + "V")
        self.ui.label_11.adjustSize()
        self.progressBarHorizontal(df["V10"], self.ui.progressBar_10)

        self.ui.label_12.setText("V11 - " + str(df["V11"]) + "V")
        self.ui.label_12.adjustSize()
        self.progressBarHorizontal(df["V11"], self.ui.progressBar_11)

        self.ui.label_13.setText("V12 - " + str(df["V12"]) + "V")
        self.ui.label_13.adjustSize()
        self.progressBarHorizontal(df["V12"], self.ui.progressBar_12)

        self.ui.label_14.setText("V13 - " + str(df["V13"]) + "V")
        self.ui.label_14.adjustSize()
        self.progressBarHorizontal(df["V13"], self.ui.progressBar_13)

        self.ui.label_15.setText("V14 - " + str(df["V14"]) + "V")
        self.ui.label_15.adjustSize()
        self.progressBarHorizontal(df["V14"], self.ui.progressBar_14)

        self.ui.label_16.setText("V15 - " + str(df["V15"]) + "V")
        self.ui.label_16.adjustSize()
        self.progressBarHorizontal(df["V15"], self.ui.progressBar_15)

        self.ui.lcd_No_of_cycles.display(df["Number Of Cycles"]*100)

        max_v = max(df["V1"], df["V2"], df["V3"], df["V4"], df["V5"], df["V6"],
                    df["V7"], df["V8"], df["V9"], df["V10"], df["V11"],
                    df["V12"], df["V13"], df["V14"], df["V15"])

        min_v = min(df["V1"], df["V2"], df["V3"], df["V4"], df["V5"], df["V6"],
                    df["V7"], df["V8"], df["V9"], df["V10"], df["V11"],
                    df["V12"], df["V13"], df["V14"], df["V15"])

        self.ui.lcdNumber_3.display(max_v)
        self.ui.lcd_Min_voltage.display(min_v)

        self.ui.lcd_T1.display(df["NTC1"])
        self.ui.lcd_T2.display(df["NTC2"])
        self.ui.lcd_T3.display(df["NTC3"])
        # self.ui.lcd_T4.display(df["NTC4"])
        # self.ui.lcd_T5.display(df["NTC5"])
        # self.ui.lcd_T6.display(df["NTC6"])

        self.ui.textBrowser_2.setPlainText(datetime.now().strftime("%Y-%m-%d"))
        self.ui.textBrowser_3.setPlainText(datetime.now().strftime("%H:%M:%S"))

        QtTest.QTest.qWait(10)

    def progressBarValue(self, value, widget, color):
        # PROGRESSBAR STYLESHEET BASE
        styleSheet = """
        QFrame{
            border-radius: 110px;
            background-color: qconicalgradient(cx:0.5, cy:0.5, angle:90, stop:{STOP_1} rgba(255, 0, 127, 0), stop:{STOP_2} {COLOR});
        }
        """

        # GET PROGRESS BAR VALUE, CONVERT TO FLOAT AND INVERT VALUES
        # stop works of 1.000 to 0.000

        progress = (100 - value) / 100.0

        # GET NEW VALUES
        stop_1 = str(progress - 0.001)
        stop_2 = str(progress)

        # FIX MAX VALUE
        if value == 100:
            stop_1 = "1.000"
            stop_2 = "1.000"

        # SET VALUES TO NEW STYLESHEET
        newStylesheet = styleSheet.replace("{STOP_1}", stop_1).replace(
            "{STOP_2}", stop_2).replace("{COLOR}", color)

        # APPLY STYLESHEET WITH NEW VALUES
        widget.setStyleSheet(newStylesheet)

    def progressBarValue2(self, value, widget, color):
        # PROGRESSBAR STYLESHEET BASE
        styleSheet = """
        QFrame{
            border-radius: 110px;
            background-color: qconicalgradient(cx:0.5, cy:0.5, angle:90, stop:{STOP_1} rgba(255, 0, 127, 0), stop:{STOP_2} {COLOR});
        }
        """

        # GET PROGRESS BAR VALUE, CONVERT TO FLOAT AND INVERT VALUES
        # stop works of 1.000 to 0.000

        progress = (500 - abs(value)) / 500.0

        # GET NEW VALUES
        stop_1 = str(progress - 0.001)
        stop_2 = str(progress)

        # FIX MAX VALUE
        if value == 500:
            stop_1 = "1.000"
            stop_2 = "1.000"

        # SET VALUES TO NEW STYLESHEET
        newStylesheet = styleSheet.replace("{STOP_1}", stop_1).replace(
            "{STOP_2}", stop_2).replace("{COLOR}", color)

        # APPLY STYLESHEET WITH NEW VALUES
        widget.setStyleSheet(newStylesheet)

    def progressBarHorizontal(self, value, widget):
        n = 5
        percent = (value / n) * 100
        widget.setValue(int(percent))

    def refresh(self):

        self.ui.label_charging_discharging.setText("STATIC")
        self.ui.label_charging_discharging.adjustSize()

        self.progressBarValue(0, self.ui.circularProgress_voltage,
                              "rgba(255, 0, 127, 255)")
        self.ui.labelPercentage_voltage.setText(str(0))

        self.progressBarValue(0, self.ui.circularProgress_current,
                              "rgba(85, 170, 255, 255)")
        self.ui.labelPercentage_current.setText(str(0))

        self.progressBarValue(0, self.ui.circularProgressSOC,
                              "rgba(85, 255, 127, 255)")
        self.ui.labelPercentageSOC.setText(str(0))

        self.ui.label_2.setText("V1")
        self.ui.label_2.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_1)

        self.ui.label_3.setText("V2")
        self.ui.label_3.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_2)

        self.ui.label_4.setText("V3")
        self.ui.label_4.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_3)

        self.ui.label_5.setText("V4")
        self.ui.label_5.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_4)

        self.ui.label_6.setText("V5")
        self.ui.label_6.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_5)

        self.ui.label_7.setText("V6")
        self.ui.label_7.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_6)

        self.ui.label_8.setText("V7")
        self.ui.label_8.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_7)

        self.ui.label_9.setText("V8")
        self.ui.label_9.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_8)

        self.ui.label_10.setText("V9")
        self.ui.label_10.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_9)

        self.ui.label_11.setText("V10")
        self.ui.label_11.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_10)

        self.ui.label_12.setText("V11")
        self.ui.label_12.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_11)

        self.ui.label_13.setText("V12")
        self.ui.label_13.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_12)

        self.ui.label_14.setText("V13")
        self.ui.label_14.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_13)

        self.ui.label_15.setText("V14")
        self.ui.label_15.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_14)

        self.ui.label_16.setText("V15")
        self.ui.label_16.adjustSize()
        self.progressBarHorizontal(0, self.ui.progressBar_15)

        self.ui.lcd_No_of_cycles.display(0)

        self.ui.lcdNumber_3.display(0)
        self.ui.lcd_Min_voltage.display(0)

        self.ui.lcd_T1.display(0)
        self.ui.lcd_T2.display(0)
        self.ui.lcd_T3.display(0)
        self.ui.lcd_T4.display(0)
        self.ui.lcd_T5.display(0)
        self.ui.lcd_T6.display(0)

        self.ui.textBrowser_2.setPlainText(str(0))
        self.ui.textBrowser_3.setPlainText(str(0))


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = code(MainWindow)
    ui.start()
    sys.exit(app.exec_())

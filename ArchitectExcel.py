import wx

from SystemExcel import *
from FunctionForArhitectExcel import *
from OemSpecific import *
from A2LColumn import *
import global_
from GUI_Ppar import *
from Logging import *
import win32file

logging.basicConfig(filename='LogError.log', filemode='a', format='%(asctime)s - %(message)s',
                    datefmt='%d-%b-%y %H:%M:%S')


def takeFirstAddress():  # this function is used for reading the start address from Architect excel and will be written on the Output excel
    address = global_.sheetArchitect.cell(global_.readRowAddress, global_.readColumnAddress).value
    address32 = conversionInt32Address(address)
    global_.sheetOutput.cell(row=global_.writeRowAddress + 1, column=global_.writeColumnAddress, value=address32)
    global_.sheetOutput.cell(row=global_.writeRowAddress, column=global_.writeColumnAddress, value=address32)
    global_.excelWriteOutput.save(global_.pathOutput)


def parameterAdaptation(list):  # this function calls other functions to adapt the parameters from Architect excel
    wx.Yield()  # GUI Refresh                                                           #to be written in the Output excel
    file = open("LogWarning.txt", "w")
    file.write("[ARCHITECT SECTION]\n")
    for self in list:
        self.count = conversionCountFor2Value(self.type, self.count)
        auxType = conversionAndVerifyType(self.type)
        if auxType is False:
            logging.warning("[TYPE ERROR] " + "Name of the parameter: " + self.name + "  Type of the parameter: " + self.type)
            global_.stateOfTheProgram = "TYPE ERROR\n" + "Name of the parameter: " + self.name + "  Type of the parameter: " + self.type
            raise RuntimeError

        self.type = auxType
        self.value = completeNoneValue(self.type, self.value, self.defaultValue)
        self.value = conversionValue(self.type, self.count, self.value)

        if valid_data(self.type, self.count, self.value) is False:
            file.write("[OUT OF RANGE]  Name: " + str(self.name) + ", Type: " + str(self.type) + ", Count: " + str(self.count) +  ", Value: " + str(self.value) + "\n")

        if isinstance(self.value, str) is True and ("0x" in self.value or "0X" in self.value) and ("CalSet".upper() not in self.name.upper() and "AntCompValue".upper() not in self.name.upper()):
            lengthValues(self.name, self.value, self.count, self.type, file)

    file.close()


class DataFromXlsArchitect:  # this class is created to take the parameters for Architect excel
    def __init__(self, name, type, count, value, L2Architect, TreeLevel2, defaultValue):
        self.name = name
        self.type = type
        self.count = count
        self.value = value
        self.nameOem = None
        self.countOem = None
        self.valueOem = None
        self.L2Architect = L2Architect
        self.TreeLevel2 = TreeLevel2
        self.defaultValue = defaultValue

    def __repr__(self):
        return "(" + str(self.name) + "," + str(self.type) + "," + str(self.count) + "," + str(self.value) + "," + \
               str(self.L2Architect) + "," + str(self.TreeLevel2) + "," + str(self.countOem) + "," + str(
            self.valueOem) + "," + str(self.defaultValue) + ")"

    def readXMLArchitect(self, listOem):  # this function is used to read the parameters from Architect excel
        listDataFromArchitect = []  # which will be stored in an object, which will be added to a list of objects

        while global_.row:
            wx.YieldIfNeeded()  # GUI Refresh
            enableOEM = False

            self.name = global_.sheetArchitect.cell(global_.row, global_.nameColumn).value
            if self.name == global_.rowEnd:
                global_.row = False
                break

            self.type = global_.sheetArchitect.cell(global_.row, global_.typeColumn).value
            self.count = global_.sheetArchitect.cell(global_.row, global_.countColumn).value
            self.defaultValue = global_.sheetArchitect.cell(global_.row, global_.defaultColumn).value

            if isinstance(global_.sheetArchitect.cell(global_.row, global_.valueColumn).value, str):
                if global_.sheetArchitect.cell(global_.row, global_.valueColumn).value is not None:
                    self.value = global_.sheetArchitect.cell(global_.row, global_.valueColumn).value.strip()
            else:
                self.value = global_.sheetArchitect.cell(global_.row, global_.valueColumn).value

            if self.name != None:
                if self.name.upper() == "Complex channel data".upper():
                    self.type = "sint8"
                    self.count = 3616

            if self.name == "Placeholder for OemSwBlock":
                for iterator in listOem:
                    listDataFromArchitect.append(iterator)
                    enableOEM = True

            if enableOEM is False:
                if self.name is not None:
                    if "res" in self.name and len(self.name) < 6:
                        if self.type is None:
                            resCount = str(global_.sheetArchitect.cell(global_.row, global_.resColumn).value)
                            self.name = "res"
                            self.type = "uint8"
                            self.count = int(resCount)
                            if global_.sheetArchitect.cell(global_.row, global_.valueColumn).value is not None:
                                self.value = global_.sheetArchitect.cell(global_.row, global_.valueColumn).value.strip()
                            else:
                                self.value = "0x00"
                            Object = DataFromXlsArchitect(self.name, self.type, self.count, self.value, self.L2Architect, self.TreeLevel2, self.defaultValue)
                            listDataFromArchitect.append(Object)

                        elif self.type is not None:
                            if not isinstance(self.value, int):
                                if not isinstance(self.value, float):
                                    if self.value is not None:
                                        self.value = self.value.strip()

                            Object = DataFromXlsArchitect(self.name, self.type, self.count, self.value, self.L2Architect, self.TreeLevel2, self.defaultValue)
                            listDataFromArchitect.append(Object)

                    else:
                        if self.type is not None:
                            Object = DataFromXlsArchitect(self.name, self.type, self.count, self.value, self.L2Architect, self.TreeLevel2, self.defaultValue)
                            listDataFromArchitect.append(Object)

                elif global_.sheetArchitect.cell(global_.row, global_.resColumn).value != 0 and global_.sheetArchitect.cell(global_.row, global_.resColumn).value is not None:
                    if self.type is None:
                        resCount = str(global_.sheetArchitect.cell(global_.row, global_.resColumn).value)
                        self.name = "res"
                        self.type = "uint8"
                        self.count = int(resCount)

                        if global_.sheetArchitect.cell(global_.row, global_.valueColumn).value is not None:
                            self.value = self.value = global_.sheetArchitect.cell(global_.row, global_.valueColumn).value.strip()
                        else:
                            self.value = "0x00"
                        Object = DataFromXlsArchitect(self.name, self.type, self.count, self.value, self.L2Architect, self.TreeLevel2, self.defaultValue)
                        listDataFromArchitect.append(Object)

                if self.name is not None and self.type is None and self.value is None:
                    nameSecondary = str(global_.sheetArchitect.cell(global_.row + 1, global_.nameColumn).value)
                    if nameSecondary is not None and global_.sheetArchitect.cell(global_.row + 1, global_.typeColumn).value is None and global_.sheetArchitect.cell(global_.row + 1,
                                                                                                                                                                    global_.valueColumn).value is None:
                        self.L2Architect = str(self.name)
                        self.TreeLevel2 = nameSecondary
                    else:
                        self.TreeLevel2 = self.name

            global_.row = global_.row + 1

        return listDataFromArchitect

    @staticmethod
    def writeInExcel(listDataFromArchitect):  # this function writes the data retrieved from the Architect excel in the Output excel
        contorRes = 0
        listRes = []

        for self in listDataFromArchitect:
            wx.Yield()  # GUI Refresh

            if self.name is None:    # pragma: no cover
                break

            if self.name is not None:
                if "res" in self.name and len(self.name) <= 5:
                    self.name = "res" + str(contorRes) + "." + str(self.TreeLevel2)
                    contorRes += 1

            global_.sheetOutput.cell(row=global_.rowWrite, column=global_.nameColumnWrite, value=self.name)
            global_.sheetOutput.cell(row=global_.rowWrite, column=global_.typeColumnWrite, value=self.type)
            global_.sheetOutput.cell(row=global_.rowWrite, column=global_.countColumnWrite, value=self.count)
            global_.sheetOutput.cell(row=global_.rowWrite, column=global_.writeL2Architect, value=self.L2Architect)
            global_.sheetOutput.cell(row=global_.rowWrite, column=global_.writeTreeLevel2, value=self.TreeLevel2)

            if self.value is not None:
                if isinstance(self.value, str):
                    if "res" in self.name and "." in self.name and (self.type == "uint8" or self.type == "sint8") and len(self.value) > 4:
                        if self.count > 32:
                            global_.sheetOutput.cell(row=global_.rowWrite, column=global_.valueColumnWrite, value="add_PPAR_default_data")
                            listRes.clear()

                            if 2 * self.count == len(self.value[2:]):
                                startString = 2
                                stopString = 4
                                for iterator in range(0, self.count):
                                    valueAux = self.value[startString:stopString]
                                    valueAux = "0x" + valueAux
                                    listRes.append(valueAux)
                                    startString += 2
                                    stopString += 2

                            else:
                                print(self.name, "ERROR AT SIZE/COUNT RES")

                            global_.sheetOutput2.cell(row=1, column=global_.columnRes, value=self.TreeLevel2)
                            global_.sheetOutput2.cell(row=global_.rowRes, column=global_.columnRes, value=self.name)
                            rowRes2 = global_.rowRes + 1
                            for valueRes in listRes:
                                global_.sheetOutput2.cell(row=rowRes2, column=global_.columnRes, value=valueRes)
                                rowRes2 += 1
                            global_.columnRes += 1

                        elif self.count <= 32:
                            startString = 2
                            endString = 4
                            for iterator in range(0, self.count):
                                global_.sheetOutput.cell(row=global_.rowWrite,
                                                         column=global_.valueColumnWrite + (4 * iterator),
                                                         value="0x" + self.value[startString:endString])
                                startString = startString + 2
                                endString = endString + 2

                    elif "res" in self.name and "." in self.name and (self.type == "uint16" or self.type == "sint16") and len(self.value) > 6:
                        if self.count > 32:
                            global_.sheetOutput.cell(row=global_.rowWrite, column=global_.valueColumnWrite,
                                                     value="add_PPAR_default_data")
                            listRes.clear()
                            if 4 * self.count == len(self.value[2:]):
                                startString = 2
                                stopString = 6
                                for iterator in range(0, self.count):
                                    valueAux = self.value[startString:stopString]
                                    valueAux = "0x" + valueAux
                                    listRes.append(valueAux)
                                    startString += 4
                                    stopString += 4
                            else:
                                print(self.name, "ERROR AT SIZE/COUNT RES")
                            global_.sheetOutput2.cell(row=1, column=global_.columnRes, value=self.TreeLevel2)
                            global_.sheetOutput2.cell(row=global_.rowRes, column=global_.columnRes, value=self.name)
                            rowRes2 = global_.rowRes + 1
                            for valueRes in listRes:
                                global_.sheetOutput2.cell(row=rowRes2, column=global_.columnRes, value=valueRes)
                                rowRes2 += 1
                            global_.columnRes += 1

                        elif self.count <= 32:
                            startString = 2
                            endString = 6
                            for iterator in range(0, self.count):
                                global_.sheetOutput.cell(row=global_.rowWrite,
                                                         column=global_.valueColumnWrite + (4 * iterator),
                                                         value="0x" + self.value[startString:endString])
                                startString = startString + 4
                                endString = endString + 4

                    elif "res" in self.name and "." in self.name and (self.type == "uint32" or self.type == "sint32") and len(self.value) > 10:
                        if self.count > 32:
                            global_.sheetOutput.cell(row=global_.rowWrite, column=global_.valueColumnWrite,
                                                     value="add_PPAR_default_data")
                            listRes.clear()

                            if 8 * self.count == len(self.value[2:]):
                                startString = 2
                                stopString = 10
                                for iterator in range(0, self.count):
                                    valueAux = self.value[startString:stopString]
                                    valueAux = "0x" + valueAux
                                    listRes.append(valueAux)
                                    startString += 8
                                    stopString += 8
                            else:
                                print(self.name, "ERROR AT SIZE/COUNT RES")

                            global_.sheetOutput2.cell(row=1, column=global_.columnRes, value=self.TreeLevel2)
                            global_.sheetOutput2.cell(row=global_.rowRes, column=global_.columnRes, value=self.name)
                            rowRes2 = global_.rowRes + 1
                            for valueRes in listRes:
                                global_.sheetOutput2.cell(row=rowRes2, column=global_.columnRes, value=valueRes)
                                rowRes2 += 1
                            global_.columnRes += 1

                        elif self.count <= 32:
                            startString = 2
                            endString = 10
                            for iterator in range(0, self.count):
                                global_.sheetOutput.cell(row=global_.rowWrite,
                                                         column=global_.valueColumnWrite + (4 * iterator),
                                                         value="0x" + self.value[startString:endString])
                                startString = startString + 8
                                endString = endString + 8

                    elif (self.type == "uint8" or self.type == "sint8") and len(self.value) > 4 and self.count <= 32:
                        startString = 2
                        endString = 4
                        for iterator in range(0, self.count):
                            global_.sheetOutput.cell(row=global_.rowWrite, column=global_.valueColumnWrite + (4 * iterator),
                                                     value="0x" + self.value[startString:endString])
                            startString = startString + 2
                            endString = endString + 2

                    elif (self.type == "uint8" or self.type == "sint8") and len(self.value) > 4 and self.count > 32:
                        global_.sheetOutput.cell(row=global_.rowWrite, column=global_.valueColumnWrite, value="add_PPAR_default_data")
                        listRes.clear()
                        if 2 * self.count == len(self.value[2:]):
                            startString = 2
                            stopString = 4
                            for iterator in range(0, self.count):
                                valueAux = self.value[startString:stopString]
                                valueAux = "0x" + valueAux
                                listRes.append(valueAux)
                                startString += 2
                                stopString += 2
                        else:
                            print(self.name, "ERROR AT SIZE/COUNT")

                        global_.sheetOutput2.cell(row=1, column=global_.columnWriteSheet2, value=self.TreeLevel2)
                        global_.sheetOutput2.cell(row=global_.rowRes, column=global_.columnWriteSheet2, value=self.name)
                        rowRes2 = global_.rowRes + 1
                        for valueRes in listRes:
                            global_.sheetOutput2.cell(row=rowRes2, column=global_.columnWriteSheet2, value=valueRes)
                            rowRes2 += 1
                        global_.columnWriteSheet2 += 1

                    elif (self.type == "uint16" or self.type == "sint16") and (len(self.value) > 13 or len(self.value) == 10) and self.count <= 32:
                        startString = 2
                        endString = 6
                        for iterator in range(0, self.count):
                            global_.sheetOutput.cell(row=global_.rowWrite, column=global_.valueColumnWrite + (4 * iterator),
                                                     value="0x" + self.value[startString:endString])
                            startString = startString + 4
                            endString = endString + 4

                    elif (self.type == "uint16" or self.type == "sint16") and (len(self.value) > 13 or len(self.value) == 10) and self.count > 32:
                        global_.sheetOutput.cell(row=global_.rowWrite, column=global_.valueColumnWrite,
                                                 value="add_PPAR_default_data")
                        listRes.clear()
                        if 4 * self.count == len(self.value[2:]):
                            startString = 2
                            stopString = 6
                            for iterator in range(0, self.count):
                                valueAux = self.value[startString:stopString]
                                valueAux = "0x" + valueAux
                                listRes.append(valueAux)
                                startString += 4
                                stopString += 4
                        else:
                            print(self.name, "ERROR AT SIZE/COUNT RES")
                        global_.sheetOutput2.cell(row=1, column=global_.columnWriteSheet2, value=self.TreeLevel2)
                        global_.sheetOutput2.cell(row=global_.rowRes, column=global_.columnWriteSheet2, value=self.name)
                        rowRes2 = global_.rowRes + 1
                        for valueRes in listRes:
                            global_.sheetOutput2.cell(row=rowRes2, column=global_.columnWriteSheet2, value=valueRes)
                            rowRes2 += 1
                        global_.columnWriteSheet2 += 1

                    elif (self.type == "uint32" or self.type == "sint32") and (len(self.value) > 10) and self.count <= 32:
                        startString = 2
                        endString = 10
                        for iterator in range(0, self.count):
                            global_.sheetOutput.cell(row=global_.rowWrite, column=global_.valueColumnWrite + (4 * iterator),
                                                     value="0x" + self.value[startString:endString])
                            startString = startString + 8
                            endString = endString + 8

                    elif (self.type == "uint32" or self.type == "sint32") and (len(self.value) > 10) and self.count > 32:
                        global_.sheetOutput.cell(row=global_.rowWrite, column=global_.valueColumnWrite,
                                                 value="add_PPAR_default_data")
                        listRes.clear()

                        if 8 * self.count == len(self.value[2:]):
                            startString = 2
                            stopString = 10
                            for iterator in range(0, self.count):
                                valueAux = self.value[startString:stopString]
                                valueAux = "0x" + valueAux
                                listRes.append(valueAux)
                                startString += 8
                                stopString += 8
                        else:
                            print(self.name, "ERROR AT SIZE/COUNT RES")

                        global_.sheetOutput2.cell(row=1, column=global_.columnWriteSheet2, value=self.TreeLevel2)
                        global_.sheetOutput2.cell(row=global_.rowRes, column=global_.columnWriteSheet2, value=self.name)
                        rowRes2 = global_.rowRes + 1
                        for valueRes in listRes:
                            global_.sheetOutput2.cell(row=rowRes2, column=global_.columnWriteSheet2, value=valueRes)
                            rowRes2 += 1
                        global_.columnWriteSheet2 += 1

                    elif (self.type == "sint16" or self.type == "uint16") and len(self.value) == 13:
                        global_.sheetOutput.cell(row=global_.rowWrite, column=global_.valueColumnWrite, value=self.value[0:6])
                        global_.sheetOutput.cell(row=global_.rowWrite, column=global_.valueColumnWrite + 4, value=self.value[7:])
                    else:
                        global_.sheetOutput.cell(row=global_.rowWrite, column=global_.valueColumnWrite, value=self.value)

                else:

                    global_.sheetOutput.cell(row=global_.rowWrite, column=global_.valueColumnWrite, value=self.value)

            global_.rowWrite = global_.rowWrite + 1
        try:
            global_.excelWriteOutput.save(global_.pathOutput)
        except:  # pragma: no cover
            pass


if __name__ == "__main__":          # 1 min 10 secs vs 1 min 50 secs     # pragma: no cover
    progressMax = 13
    count = 0
    var = "Start"
    listOemSpecific = []

    global_.dialog = wx.ProgressDialog("Progress PPAR Tool", var, progressMax,
                                       style=wx.PD_CAN_ABORT | wx.PD_SMOOTH | wx.PD_ESTIMATED_TIME | wx.PD_ELAPSED_TIME | wx.PD_AUTO_HIDE)

    keepGoing = global_.dialog.Update(count, "Loading the Architect Excel...")

    global_.excelReadArchitect = load_workbook(filename=global_.pathArchitect, data_only=True)
    global_.sheetArchitect = global_.excelReadArchitect[global_.architectSheet]

    wx.Yield()  # GUI Refresh
    count += 1

    keepGoing = global_.dialog.Update(count, "Loading the Output Excel...")

    win32file.SetFileAttributes(global_.pathOutput, win32file.FILE_ATTRIBUTE_NORMAL)        # Set Read Only OFF

    global_.excelWriteOutput = load_workbook(filename=global_.pathOutput)
    global_.sheetOutput = global_.excelWriteOutput[global_.output1Sheet]
    global_.sheetOutput2 = global_.excelWriteOutput[global_.output2Sheet]

    count += 1
    wx.Yield()  # GUI Refresh

    if not global_.dialog.WasCancelled():

        keepGoing = global_.dialog.Update(count, "Reading OEM data from Architect excel")
        # breakpoint()

        try:
            objectForOem = OemSpecific("0", "0", "0", "0", "0", "0", "0")  # pragma: no cover
            if global_.checkOem:
                listOemSpecific = objectForOem.readOemSpecific()  # pragma: no cover
        except Exception as e:
            logging.error("Exception occurred: " + e.args[0], exc_info=True)
            global_.dialog.Destroy()
            global_.stateOfTheProgram = "An error occurred in OEM section\nFor more information check LogError! "
        else:
            if global_.checkOemName:
                count += 1
                keepGoing = global_.dialog.Update(count, "Reading OEM data from Architect excel")
                global_.dialog.Destroy()
            else:
                count += 1
                if not global_.dialog.WasCancelled():
                    keepGoing = global_.dialog.Update(count, "Reading data from Architect excel")
                    try:
                        objectForArhitectData = DataFromXlsArchitect("0", "0", "0", "0", "0", "0", "0")  # pragma: no cover
                        listForArhitectData = DataFromXlsArchitect.readXMLArchitect(objectForArhitectData, listOemSpecific)  # pragma: no cover

                    except Exception as e:
                        logging.error("Exception occurred: " + e.args[0], exc_info=True)
                        global_.dialog.Destroy()
                        global_.stateOfTheProgram = "An error occurred while reading data from Architect excel\n" \
                                                    "For more information check LogError!"
                    else:
                        count += 1
                        if not global_.dialog.WasCancelled():
                            keepGoing = global_.dialog.Update(count, "Parameter adaptation")
                            try:
                                parameterAdaptation(listForArhitectData)  # pragma: no cover
                            except RuntimeError:
                                global_.dialog.Destroy()
                            except Exception as e:
                                logging.error("Exception occurred: " + e.args[0], exc_info=True)
                                global_.dialog.Destroy()
                                global_.stateOfTheProgram = "Parameter adaptation error.\nFor more information check LogError! "

                            else:
                                count += 1
                                if not global_.dialog.WasCancelled():
                                    keepGoing = global_.dialog.Update(count, "Writing data from Architect excel")
                                    try:
                                        objectForArhitectData.writeInExcel(listForArhitectData)  # pragma: no cover
                                    except Exception as e:
                                        logging.error("Exception occurred: " + e.args[0], exc_info=True)
                                        global_.dialog.Destroy()
                                        global_.stateOfTheProgram = "An error occurred while writing data from Architect " \
                                                                    "excel\nFor more information check LogError!"

                                    else:
                                        count += 1

                                        keepGoing = global_.dialog.Update(count, "Loading the System Excel... This might take a while...")

                                        global_.excelReadSystem = load_workbook(filename=global_.pathSystem)
                                        global_.sheetSystem = global_.excelReadSystem[global_.systemSheet]

                                        wx.Yield()  # GUI Refresh
                                        count += 1
                                        if not global_.dialog.WasCancelled():
                                            keepGoing = global_.dialog.Update(count, "Reading and writing start address from Architect excel")

                                            try:
                                                takeFirstAddress()  # pragma: no cover
                                                objectForSystemData = DataFromXlsSystem2("0", "0", "0", "0", "0", "0", "0", "0")  # pragma: no cover

                                            except Exception as e:
                                                logging.error("Exception occurred: " + e.args[0], exc_info=True)
                                                global_.dialog.Destroy()
                                                global_.stateOfTheProgram = "An error occurred while reading and writing start address from Architect excel\nFor more information check LogError!"

                                            else:
                                                count += 1

                                                if not global_.dialog.WasCancelled():
                                                    keepGoing = global_.dialog.Update(count, "Reading data from System excel")
                                                    try:
                                                        listForSystemData = DataFromXlsSystem2.readXMLSystem2(
                                                            objectForSystemData)  # pragma: no cover
                                                    except Exception as e:
                                                        logging.error("Exception occurred: " + e.args[0], exc_info=True)
                                                        global_.dialog.Destroy()
                                                        global_.stateOfTheProgram = "An error occurred while reading data from System excel\nFor more information check LogError!"

                                                    else:
                                                        count += 1

                                                        if not global_.dialog.WasCancelled():
                                                            keepGoing = global_.dialog.Update(count, "Writing data from System excel")

                                                            try:
                                                                columnOem = objectForSystemData.writeInExcel2(listForSystemData)  # pragma: no cover
                                                            except Exception as e:
                                                                logging.error("Exception occurred: " + e.args[0],
                                                                              exc_info=True)
                                                                global_.dialog.Destroy()
                                                                global_.stateOfTheProgram = "An error occurred while writing data from System excel\nFor more information check LogError!"

                                                            else:

                                                                count += 1

                                                                if not global_.dialog.WasCancelled():
                                                                    keepGoing = global_.dialog.Update(count, "Reading A2L data")
                                                                    try:
                                                                        listA2L = a2lFunction()
                                                                    except Exception as e:
                                                                        logging.error("Exception occurred: " + e.args[0],
                                                                                      exc_info=True)
                                                                        global_.dialog.Destroy()
                                                                        global_.stateOfTheProgram = "An error occurred while reading A2L data\nFor more information check LogError!"

                                                                    else:

                                                                        count += 1

                                                                        if not global_.dialog.WasCancelled():
                                                                            keepGoing = global_.dialog.Update(count, "Writing A2L data")

                                                                            try:
                                                                                a2lWrite(listA2L)
                                                                            except Exception as e:
                                                                                logging.error("Exception occurred: " + e.args[0], exc_info=True)

                                                                                global_.dialog.Destroy()
                                                                                global_.stateOfTheProgram = "An error occurred while writing A2L data\nFor more information check LogError!"

                                                                            else:
                                                                                count += 1

                                                                                if not global_.dialog.WasCancelled():
                                                                                    keepGoing = global_.dialog.Update(count, "Parameter verification in progress")

                                                                                    try:
                                                                                        parameterVerificationOutput(listForArhitectData, listForSystemData)

                                                                                    except RuntimeError:
                                                                                        global_.dialog.Destroy()
                                                                                    except Exception as e:
                                                                                        logging.error("Exception occurred: " + e.args[0], exc_info=True)

                                                                                        global_.dialog.Destroy()
                                                                                        global_.stateOfTheProgram = "An error occurred while the verification of the Output\n" \
                                                                                                                    "For more information check LogError!"
                                                                                    else:

                                                                                        print("The program finished execution!")
                                                                                        count += 1
                                                                                        keepGoing = global_.dialog.Update(count, "The program finished execution!")
                                                                                        global_.stateOfTheProgram = "The program finished execution!"

                                                                                else:
                                                                                    global_.stateOfTheProgram = "The program was canceled!"

                                                                        else:
                                                                            global_.stateOfTheProgram = "The program was canceled!"
                                                                else:
                                                                    global_.stateOfTheProgram = "The program was canceled!"
                                                        else:
                                                            global_.stateOfTheProgram = "The program was canceled!"
                                                else:
                                                    global_.stateOfTheProgram = "The program was canceled!"
                                        else:
                                            global_.stateOfTheProgram = "The program was canceled!"
                                else:
                                    global_.stateOfTheProgram = "The program was canceled!"
                        else:
                            global_.stateOfTheProgram = "The program was canceled!"
                else:
                    global_.stateOfTheProgram = "The program was canceled!"
    else:
        global_.stateOfTheProgram = "The program was canceled!"

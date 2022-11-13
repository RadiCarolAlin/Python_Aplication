import wx

from SystemExcel import *
from FunctionForArhitectExcel import *
from ArchitectExcel import *
from GUI_Ppar import *
import global_


class OemSpecific:  # this class is created to take the Oem parameters for Architect excel
    def __init__(self, name, type, count, value, L2Architect, TreeLevel2, defaultValue):
        self.name = name
        self.type = type
        self.count = count
        self.value = value
        self.L2Architect = L2Architect
        self.TreeLevel2 = TreeLevel2
        self.defaultValue = defaultValue

    def __repr__(self):
        return "(" + str(self.name) + "," + str(self.type) + "," + str(self.count) + "," + str(
            self.value) + "," + str(self.L2Architect) + "," + str(self.defaultValue) + "," + str(self.TreeLevel2) + ")"

    def readOemSpecific(self):  # this function is used to read the Oem parameters from Architect excel
        nameOfTheProject = global_.sheetArchitect.cell(1, global_.valueColumn).value
        index = global_.projectOemName.rindex("_")
        nameOfTheOemSheet = global_.projectOemName[index + 1:]

        if nameOfTheOemSheet.upper() != nameOfTheProject.upper():
            global_.checkOemName = True
            global_.stateOfTheProgram = "The project OEM name doesn't correspond with the sheet OEM name!\n" + nameOfTheOemSheet + " != " + nameOfTheProject

        else:
            listSheets = global_.excelReadArchitect.sheetnames
            listOem = []
            for sheet in listSheets:
                wx.Yield()
                if global_.projectOemName == sheet:
                    sheetOem = global_.excelReadArchitect[sheet]
                    while global_.rowOem:
                        self.name = sheetOem.cell(global_.rowOem, global_.nameColumnOem).value
                        self.type = "uint8"
                        self.count = sheetOem.cell(global_.rowOem, global_.countColumnOem).value
                        self.value = sheetOem.cell(global_.rowOem, global_.valueColumnOem).value
                        self.value = conversionOem2(self.value, self.count)
                        self.L2Architect = "PPAR_SelfTest"
                        self.TreeLevel2 = "PPAR_OemSwBlock"
                        self.defaultValue = None
                        Object = OemSpecific(self.name, self.type, self.count, self.value, self.L2Architect, self.TreeLevel2, self.defaultValue)
                        listOem.append(Object)
                        global_.rowOem += 1
                        if self.name == global_.rowEndOem:
                            global_.rowOem = False
            return listOem
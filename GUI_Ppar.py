import os  # Import section.
import subprocess as sp
import wx
import sys
import traceback
import global_
import wx.lib.scrolledpanel as scrolled
import configparser
from openpyxl import load_workbook
import xlrd
import wx.lib.agw.floatspin as FS
import zipfile
import xmltodict
from SystemExcel import *
from FunctionForArhitectExcel import *
from OemSpecific import *
from A2LColumn import *
from ArchitectExcel import *
from Logging import *


Version = "1.1"


class MyApp(wx.App):  # Creating the main class.

    def __init__(self):
        super().__init__(clearSigInt=True)
        self.InitFrame()

    @staticmethod
    def InitFrame():  # Initialization of the frame.
        frame = MyFrame(parent=None, title="PPAR TOOL | Version: " + Version, pos=(100, 100), size=(1550, 720))
        frame.Centre()
        frame.Show(True)


class MyFrame(wx.Frame):  # Subclass of wx.Window; Frame is a top level window

    def __init__(self, parent, title, pos, size):  # Make the child class inherit all the methods and
        super().__init__(parent=parent, title=title, pos=pos, size=size)  # properties from its parent
        self.SetBackgroundColour('white')
        self.OnInit()
        MenuBar = wx.MenuBar()  # Creating the menu bar.
        MenuHelp = wx.Menu()
        MenuHelp.Append(21, '&Help me', 'Information about program')

        MenuBar.Append(MenuHelp, "&HELP")

        self.Bind(wx.EVT_MENU, handler=self.onHelp, id=21)
        self.SetMenuBar(MenuBar)
        self.CreateStatusBar()

    def OnInit(self):  # Initialization of the panel.
        MyPanel(parent=self)

    @staticmethod
    def onHelp(event):  # Opening a certain file when pressing the button HELP.  # pragma: no cover
        fileName = "PPAR Generator HelpMenu.pdf"
        sp.Popen([fileName], shell=True)

    @staticmethod
    def onExec(event):  # Opening a certain file when pressing the button CFG EXEC.  # pragma: no cover
        os.chdir('CFG_Cluster_Multiplier')
        os.startfile("CFG_Cluster_Multiplier.exe")
        os.chdir(os.path.dirname(os.getcwd()))


class MyPanel(scrolled.ScrolledPanel):  # A panel is a window on which controls are placed. (e.g. buttons and text boxes)
    def __init__(self, parent):  # This class is also inherited from wxWindow class.
        super().__init__(parent=parent)  # Make the child class inherit all the methods and properties from its parent.

        # Getting the right icons.

        bmpIconInformation = wx.ArtProvider.GetBitmap(id=wx.ART_INFORMATION, client=wx.ART_OTHER, size=(16, 16))
        font = wx.Font(-1, wx.DECORATIVE, wx.BOLD, wx.NORMAL, underline=True)
        self.img1 = wx.Image("Continental_Image.jpg", wx.BITMAP_TYPE_ANY)
        self.m_bitmap3 = wx.StaticBitmap(self, wx.ID_ANY, wx.Bitmap(self.img1), wx.DefaultPosition, wx.DefaultSize, 0)
        # Setting variables with icons.

        self.arhitectIcon = wx.StaticBitmap(self, wx.ID_ANY, bmpIconInformation)
        self.arhitectIcon.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.oemIcon = wx.StaticBitmap(self, wx.ID_ANY, bmpIconInformation)
        self.systemIcon = wx.StaticBitmap(self, wx.ID_ANY, bmpIconInformation)
        self.outputIcon = wx.StaticBitmap(self, wx.ID_ANY, bmpIconInformation)
        self.architectPathIcon = wx.StaticBitmap(self, wx.ID_ANY, bmpIconInformation)
        self.systemPathIcon = wx.StaticBitmap(self, wx.ID_ANY, bmpIconInformation)
        self.outputPathIcon = wx.StaticBitmap(self, wx.ID_ANY, bmpIconInformation)

        # Creating the informational elements and binding them, so we can mouse hover and get more information.

        self.architectText = wx.StaticText(self, wx.ID_ANY, 'Architect Settings:')
        self.architectText.Bind(wx.EVT_MOTION, self.onMouseOver)

        self.columnNameArchitect = wx.StaticText(self, wx.ID_ANY, 'Name column number:')
        self.columnNameArchitect.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.nameColumnArchitect = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.columnTypeArchitect = wx.StaticText(self, wx.ID_ANY, 'Type column number:')
        self.columnTypeArchitect.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.typeColumnArchitect = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.columnCountArchitect = wx.StaticText(self, wx.ID_ANY, 'Count column number:')
        self.columnCountArchitect.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.countColumnArchitect = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.columnValueArchitect = wx.StaticText(self, wx.ID_ANY, 'Value column number:')
        self.columnValueArchitect.SetFont(font)
        self.columnValueArchitect.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.valueColumnArchitect = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.spaceArchitectText = wx.StaticText(self, wx.ID_ANY, '                                         ')

        self.rowStartArchitect = wx.StaticText(self, wx.ID_ANY, 'Row start number:')
        self.rowStartArchitect.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.startRowArchitect = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.rowEndArchitect = wx.StaticText(self, wx.ID_ANY, 'Row end:')
        self.rowEndArchitect.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.endRowArchitect = wx.TextCtrl(self, wx.ID_ANY, value='', size=(150, -1))

        self.readRowStartAddressArchitect = wx.StaticText(self, wx.ID_ANY, 'Row address number')
        self.readRowStartAddressArchitect.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.rowStartReadAddressArchitect = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.readColumnStartAddressArchitect = wx.StaticText(self, wx.ID_ANY, 'Column address number:')
        self.readColumnStartAddressArchitect.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.columnStartReadAddressArchitect = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.resColumnArchitect = wx.StaticText(self, wx.ID_ANY, 'Reserved column number:')
        self.resColumnArchitect.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.columnResArchitect = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.space1ArchitectText = wx.StaticText(self, wx.ID_ANY, '                                         ')

        self.rowA2lLabel = wx.StaticText(self, wx.ID_ANY, 'Row A2L number:')
        self.rowA2lLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.rowA2l = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.columnA2lLabel = wx.StaticText(self, wx.ID_ANY, 'Column A2L number:')
        self.columnA2lLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.columnA2l = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.oemText = wx.StaticText(self, wx.ID_ANY, 'Oem Settings:    ')
        self.oemText.Bind(wx.EVT_MOTION, self.onMouseOver)

        self.oemInput = wx.CheckBox(parent=self, label="Oem Project")
        self.oemInput.Bind(event=wx.EVT_CHECKBOX, handler=self.onCheckBox)

        self.projectOemName = wx.StaticText(self, wx.ID_ANY, 'Project OEM name:')
        self.projectOemName.SetFont(font)
        self.projectOemName.Bind(wx.EVT_MOTION, self.onMouseOver)

        self.oemProjectName = wx.ComboBox(self, wx.ID_ANY, value='', size=(250, -1), style=wx.TE_READONLY)
        self.oemProjectName.Enable(False)
        self.oemProjectName.Append("")

        self.rowStartOem = wx.StaticText(self, wx.ID_ANY, 'Row start number:')
        self.rowStartOem.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.startRowOem = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))
        self.startRowOem.Enable(False)

        self.rowEndOem = wx.StaticText(self, wx.ID_ANY, 'Row end:')
        self.rowEndOem.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.endRowOem = wx.TextCtrl(self, wx.ID_ANY, value='', size=(150, -1))
        self.endRowOem.Enable(False)

        self.spaceOemText = wx.StaticText(self, wx.ID_ANY,
                                          '                                                                     ')

        self.columnNameOem = wx.StaticText(self, wx.ID_ANY, 'Name column number:')
        self.columnNameOem.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.nameColumnOem = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))
        self.nameColumnOem.Enable(False)

        self.columnCountOem = wx.StaticText(self, wx.ID_ANY, 'Count column number:')
        self.columnCountOem.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.countColumnOem = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))
        self.countColumnOem.Enable(False)

        self.columnValueOem = wx.StaticText(self, wx.ID_ANY, 'Value column number:')
        self.columnValueOem.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.valueColumnOem = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))
        self.valueColumnOem.Enable(False)

        self.rowOemA2lLabel = wx.StaticText(self, wx.ID_ANY, 'Row A2L:')
        self.rowOemA2lLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.rowOemA2l = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))
        self.rowOemA2l.Enable(False)

        self.columnOemA2lLabel = wx.StaticText(self, wx.ID_ANY, 'Column A2L:')
        self.columnOemA2lLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.columnOemA2l = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))
        self.columnOemA2l.Enable(False)

        self.systemText = wx.StaticText(self, wx.ID_ANY, 'System Settings:')
        self.systemText.Bind(wx.EVT_MOTION, self.onMouseOver)

        self.columnNameSystem = wx.StaticText(self, wx.ID_ANY, 'Name column number:')
        self.columnNameSystem.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.nameColumnSystem = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.columnTypeSystem = wx.StaticText(self, wx.ID_ANY, 'Type column number:')
        self.columnTypeSystem.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.typeColumnSystem = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.columnCountSystem = wx.StaticText(self, wx.ID_ANY, 'Count column number:')
        self.columnCountSystem.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.countColumnSystem = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.columnLimitValueSystem = wx.StaticText(self, wx.ID_ANY, 'Limit value column number:')
        self.columnLimitValueSystem.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.limitValueColumnSystem = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.columnValueSystem = wx.StaticText(self, wx.ID_ANY, 'Value column number:')
        self.columnValueSystem.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.valueColumnSystem = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.space1SystemText = wx.StaticText(self, wx.ID_ANY, '                                     ')

        self.columnLowLimitSystem = wx.StaticText(self, wx.ID_ANY, 'Low column number:')
        self.columnLowLimitSystem.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.lowLimitColumnSystem = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.columnMaxLimitSystem = wx.StaticText(self, wx.ID_ANY, 'Max column number:')
        self.columnMaxLimitSystem.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.maxLimitColumnSystem = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.rowStartSystemLabel = wx.StaticText(self, wx.ID_ANY, 'Row start number:')
        self.rowStartSystemLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.rowStartSystem = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.rowEndSystemLabel = wx.StaticText(self, wx.ID_ANY, 'Row end:')
        self.rowEndSystemLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.rowEndSystem = wx.TextCtrl(self, wx.ID_ANY, value='', size=(162, -1))

        self.projectSpecificNameLabel = wx.StaticText(self, wx.ID_ANY, 'Project specific name:')
        self.projectSpecificNameLabel.SetFont(font)
        self.projectSpecificNameLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.projectSpecificName = wx.TextCtrl(self, wx.ID_ANY, value='', size=(140, -1))

        self.projectNameTypeLabel = wx.StaticText(self, wx.ID_ANY, 'Project name type:')
        self.projectNameTypeLabel.SetFont(font)
        self.projectNameTypeLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.projectNameType = wx.TextCtrl(self, wx.ID_ANY, value='', size=(140, -1))

        self.columnWriteSheet2Label = wx.StaticText(self, wx.ID_ANY, 'Column write second sheet column:')
        self.columnWriteSheet2Label.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.columnWriteSheet2 = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.rowWriteSheet2Label = wx.StaticText(self, wx.ID_ANY, 'Row write second sheet column:')
        self.rowWriteSheet2Label.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.rowWriteSheet2 = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.space3SystemText = wx.StaticText(self, wx.ID_ANY, '                                     ')

        self.listWithNotUsedElementsLabel = wx.StaticText(self, wx.ID_ANY, 'Elements to be ignored:')
        self.listWithNotUsedElementsLabel.SetFont(font)
        self.listWithNotUsedElementsLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.listWithNotUsedElements = wx.TextCtrl(self, wx.ID_ANY, value='', size=(750, -1))

        self.outputText = wx.StaticText(self, wx.ID_ANY, 'Output Settings:')
        self.outputText.Bind(wx.EVT_MOTION, self.onMouseOver)

        self.nameColumnWriteLabel = wx.StaticText(self, wx.ID_ANY, 'Name column number:')
        self.nameColumnWriteLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.nameColumnWrite = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.typeColumnWriteLabel = wx.StaticText(self, wx.ID_ANY, 'Type column number:')
        self.typeColumnWriteLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.typeColumnWrite = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.countColumnWriteLabel = wx.StaticText(self, wx.ID_ANY, 'Count column number:')
        self.countColumnWriteLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.countColumnWrite = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.valueColumnWriteLabel = wx.StaticText(self, wx.ID_ANY, 'Value column number :')
        self.valueColumnWriteLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.valueColumnWrite = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.rowStartWriteLabel = wx.StaticText(self, wx.ID_ANY, 'Row start number:')
        self.rowStartWriteLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.rowStartWrite = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.writeRowAddressLabel = wx.StaticText(self, wx.ID_ANY, 'Row Address number:')
        self.writeRowAddressLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.writeRowAddress = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.space1OutputText = wx.StaticText(self, wx.ID_ANY, '                                     ')

        self.writeColumnAddressLabel = wx.StaticText(self, wx.ID_ANY, 'Column Address number:')
        self.writeColumnAddressLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.writeColumnAddress = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.writeL2ArchitectLabel = wx.StaticText(self, wx.ID_ANY, 'Column L2Architect number:')
        self.writeL2ArchitectLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.writeL2Architect = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.writeTreeLevel2Label = wx.StaticText(self, wx.ID_ANY, 'Column TreeLevel number:')
        self.writeTreeLevel2Label.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.writeTreeLevel2 = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.columnResLabel = wx.StaticText(self, wx.ID_ANY, 'Column reserved number:')
        self.columnResLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.columnRes = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.defaultColumnLabel = wx.StaticText(self, wx.ID_ANY, 'Column default number:')
        self.defaultColumnLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.defaultColumn = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.space2OutputText = wx.StaticText(self, wx.ID_ANY, '                                     ')

        self.rowResLabel = wx.StaticText(self, wx.ID_ANY, 'Row reserved number :')
        self.rowResLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.rowRes = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.rowWriteA2lLabel = wx.StaticText(self, wx.ID_ANY, 'Row A2L number:')
        self.rowWriteA2lLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.rowWriteA2l = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.columnWriteA2lLabel = wx.StaticText(self, wx.ID_ANY, 'Column A2L number:')
        self.columnWriteA2lLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.columnWriteA2l = wx.SpinCtrl(self, wx.ID_ANY, value="0", min=0, max=10000, size=(70, -1))

        self.inputArchitectPathLabel = wx.StaticText(self, wx.ID_ANY, 'Architect Path:      ')
        self.inputArchitectPathLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.inputArchitectPath = wx.TextCtrl(self, wx.ID_ANY, value='Path', style=wx.TE_READONLY, size=(780, -1))
        buttonInputArchitect = wx.Button(parent=self, label="Browse")
        buttonInputArchitect.Bind(event=wx.EVT_BUTTON, handler=self.onButtonReadPathArchitect)

        self.sheet1Text = wx.StaticText(self, wx.ID_ANY, 'Sheet Architect:')
        self.sheet1Text.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.sheet1Input = wx.ComboBox(self, wx.ID_ANY, value='PPAR Definition', size=(160, -1), style=wx.TE_READONLY)

        self.inputSystemPathLabel = wx.StaticText(self, wx.ID_ANY, 'System Path:         ')
        self.inputSystemPathLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.inputSystemPath = wx.TextCtrl(self, wx.ID_ANY, value='Path', style=wx.TE_READONLY, size=(780, -1))
        buttonInputSystem = wx.Button(parent=self, label="Browse")
        buttonInputSystem.Bind(event=wx.EVT_BUTTON, handler=self.onButtonReadPathSystem)

        self.sheet2Text = wx.StaticText(self, wx.ID_ANY, 'Sheet System:   ')
        self.sheet2Text.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.sheet2Input = wx.ComboBox(self, wx.ID_ANY, value='PPAR Definition', size=(160, -1), style=wx.TE_READONLY)

        self.OutputPathLabel = wx.StaticText(self, wx.ID_ANY, 'Output Path:         ')
        self.OutputPathLabel.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.OutputPath = wx.TextCtrl(self, wx.ID_ANY, value='Path', style=wx.TE_READONLY, size=(600, -1))
        buttonOutputPath = wx.Button(parent=self, label="Browse")
        buttonOutputPath.Bind(event=wx.EVT_BUTTON, handler=self.onButtonReadPathOutCompare)

        self.sheet3Text = wx.StaticText(self, wx.ID_ANY, 'Sheet 1:')
        self.sheet3Text.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.sheet3Input = wx.ComboBox(self, wx.ID_ANY, value='PPAR_definition', size=(160, -1), style=wx.TE_READONLY)

        self.sheet4Text = wx.StaticText(self, wx.ID_ANY, 'Sheet 2:')
        self.sheet4Text.Bind(wx.EVT_MOTION, self.onMouseOver)
        self.sheet4Input = wx.ComboBox(self, wx.ID_ANY, value='add_PPAR_default_data', size=(160, -1), style=wx.TE_READONLY)

        self.buttonRun = wx.Button(parent=self, label="Run", style=wx.BORDER_NONE)
        self.buttonRun.Bind(event=wx.EVT_BUTTON, handler=self.onButtonRun)

        buttonLoad = wx.Button(parent=self, label="Load Settings", size=(160, 30))
        buttonLoad.Bind(event=wx.EVT_BUTTON, handler=self.onLoadSetting)

        buttonSave = wx.Button(parent=self, label="Save Settings", size=(160, 30))
        buttonSave.Bind(event=wx.EVT_BUTTON, handler=self.onSaveSetting)

        cancelBtn = wx.Button(self, wx.ID_ANY, 'Cancel')
        self.Bind(wx.EVT_BUTTON, self.onCancel, cancelBtn)

        # Creating sizer, so we can organize our interface in lines. This will help the user to modify the size of the
        # interface, so it looks correct on all platforms and on all resolutions.

        mainSizer = wx.BoxSizer(wx.VERTICAL)
        architectSizer = wx.BoxSizer(wx.VERTICAL)
        firstArhitectSizer = wx.BoxSizer(wx.HORIZONTAL)
        secondArhitectSizer = wx.BoxSizer(wx.HORIZONTAL)
        thirdArchitectSizer = wx.BoxSizer(wx.HORIZONTAL)
        oemSizer = wx.BoxSizer(wx.VERTICAL)
        firstOemSizer = wx.BoxSizer(wx.HORIZONTAL)
        secondOemSizer = wx.BoxSizer(wx.HORIZONTAL)
        systemSizer = wx.BoxSizer(wx.VERTICAL)
        firstSystemSizer = wx.BoxSizer(wx.HORIZONTAL)
        secondSystemSizer = wx.BoxSizer(wx.HORIZONTAL)
        fourthSystemSizer = wx.BoxSizer(wx.HORIZONTAL)
        outputSizer = wx.BoxSizer(wx.VERTICAL)
        firstOutputSizer = wx.BoxSizer(wx.HORIZONTAL)
        secondOutputSizer = wx.BoxSizer(wx.HORIZONTAL)
        thirdOutputSizer = wx.BoxSizer(wx.HORIZONTAL)
        pathSizer = wx.BoxSizer(wx.VERTICAL)
        firstPathSizer = wx.BoxSizer(wx.HORIZONTAL)
        secondPathSizer = wx.BoxSizer(wx.HORIZONTAL)
        thirdPathSizer = wx.BoxSizer(wx.HORIZONTAL)
        buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        loadAndSaveSizer = wx.BoxSizer(wx.HORIZONTAL)

        firstArhitectSizer.Add(self.arhitectIcon, 0, wx.ALL, 5)
        firstArhitectSizer.Add(self.architectText, 0, wx.ALL, 5)

        firstArhitectSizer.Add(self.columnNameArchitect, 0, wx.ALL, 5)
        firstArhitectSizer.Add(self.nameColumnArchitect, 0, wx.ALL, 5)

        firstArhitectSizer.Add(self.columnTypeArchitect, 0, wx.ALL, 5)
        firstArhitectSizer.Add(self.typeColumnArchitect, 0, wx.ALL, 5)

        firstArhitectSizer.Add(self.columnCountArchitect, 0, wx.ALL, 5)
        firstArhitectSizer.Add(self.countColumnArchitect, 0, wx.ALL, 5)

        firstArhitectSizer.Add(self.columnValueArchitect, 0, wx.ALL, 5)
        firstArhitectSizer.Add(self.valueColumnArchitect, 0, wx.ALL, 5)

        secondArhitectSizer.Add(self.spaceArchitectText, 0, wx.ALL, 5)

        secondArhitectSizer.Add(self.rowStartArchitect, 0, wx.ALL, 5)
        secondArhitectSizer.Add(self.startRowArchitect, 0, wx.ALL, 5)

        secondArhitectSizer.Add(self.rowEndArchitect, 0, wx.ALL, 5)
        secondArhitectSizer.Add(self.endRowArchitect, 0, wx.EXPAND | wx.ALL, 5)

        secondArhitectSizer.Add(self.readRowStartAddressArchitect, 0, wx.ALL, 5)
        secondArhitectSizer.Add(self.rowStartReadAddressArchitect, 0, wx.ALL, 5)

        secondArhitectSizer.Add(self.readColumnStartAddressArchitect, 0, wx.ALL, 5)
        secondArhitectSizer.Add(self.columnStartReadAddressArchitect, 0, wx.ALL, 5)

        secondArhitectSizer.Add(self.resColumnArchitect, 0, wx.ALL, 5)
        secondArhitectSizer.Add(self.columnResArchitect, 0, wx.ALL, 5)

        thirdArchitectSizer.Add(self.space1ArchitectText, 0, wx.ALL, 5)
        thirdArchitectSizer.Add(self.rowA2lLabel, 0, wx.ALL, 5)
        thirdArchitectSizer.Add(self.rowA2l, 0, wx.ALL, 5)
        thirdArchitectSizer.Add(self.columnA2lLabel, 0, wx.ALL, 5)
        thirdArchitectSizer.Add(self.columnA2l, 0, wx.ALL, 5)
        thirdArchitectSizer.Add(self.defaultColumnLabel, 0, wx.ALL, 5)
        thirdArchitectSizer.Add(self.defaultColumn, 0, wx.ALL, 5)

        firstOemSizer.Add(self.oemIcon, 0, wx.ALL, 5)
        firstOemSizer.Add(self.oemText, 0, wx.ALL, 5)
        firstOemSizer.Add(self.oemInput, 0, wx.ALL, 5)

        firstOemSizer.Add(self.projectOemName, 0, wx.ALL, 5)
        firstOemSizer.Add(self.oemProjectName, 0, wx.ALL, 5)

        firstOemSizer.Add(self.rowStartOem, 0, wx.ALL, 5)
        firstOemSizer.Add(self.startRowOem, 0, wx.ALL, 5)

        firstOemSizer.Add(self.rowEndOem, 0, wx.ALL, 5)
        firstOemSizer.Add(self.endRowOem, 0, wx.ALL, 5)

        secondOemSizer.Add(self.spaceOemText, 0, wx.ALL, 5)

        secondOemSizer.Add(self.columnNameOem, 0, wx.ALL, 5)
        secondOemSizer.Add(self.nameColumnOem, 0, wx.ALL, 5)

        secondOemSizer.Add(self.columnCountOem, 0, wx.ALL, 5)
        secondOemSizer.Add(self.countColumnOem, 0, wx.ALL, 5)

        secondOemSizer.Add(self.columnValueOem, 0, wx.ALL, 5)
        secondOemSizer.Add(self.valueColumnOem, 0, wx.ALL, 5)

        secondOemSizer.Add(self.rowOemA2lLabel, 0, wx.ALL, 5)
        secondOemSizer.Add(self.rowOemA2l, 0, wx.ALL, 5)
        secondOemSizer.Add(self.columnOemA2lLabel, 0, wx.ALL, 5)
        secondOemSizer.Add(self.columnOemA2l, 0, wx.ALL, 5)

        firstSystemSizer.Add(self.systemIcon, 0, wx.ALL, 5)
        firstSystemSizer.Add(self.systemText, 0, wx.ALL, 5)

        firstSystemSizer.Add(self.columnNameSystem, 0, wx.ALL, 5)
        firstSystemSizer.Add(self.nameColumnSystem, 0, wx.ALL, 5)

        firstSystemSizer.Add(self.columnTypeSystem, 0, wx.ALL, 5)
        firstSystemSizer.Add(self.typeColumnSystem, 0, wx.ALL, 5)

        firstSystemSizer.Add(self.columnCountSystem, 0, wx.ALL, 5)
        firstSystemSizer.Add(self.countColumnSystem, 0, wx.ALL, 5)

        firstSystemSizer.Add(self.columnValueSystem, 0, wx.ALL, 5)
        firstSystemSizer.Add(self.valueColumnSystem, 0, wx.ALL, 5)

        firstSystemSizer.Add(self.columnLimitValueSystem, 0, wx.ALL, 5)
        firstSystemSizer.Add(self.limitValueColumnSystem, 0, wx.ALL, 5)

        secondSystemSizer.Add(self.space1SystemText, 0, wx.ALL, 5)
        secondSystemSizer.Add(self.rowStartSystemLabel, 0, wx.ALL, 5)
        secondSystemSizer.Add(self.rowStartSystem, 0, wx.ALL, 5)
        secondSystemSizer.Add(self.rowEndSystemLabel, 0, wx.ALL, 5)
        secondSystemSizer.Add(self.rowEndSystem, 0, wx.ALL, 5)
        secondSystemSizer.Add(self.projectSpecificNameLabel, 0, wx.ALL, 5)
        secondSystemSizer.Add(self.projectSpecificName, 0, wx.ALL, 5)

        secondSystemSizer.Add(self.projectNameTypeLabel, 0, wx.ALL, 5)
        secondSystemSizer.Add(self.projectNameType, 0, wx.ALL, 5)

        fourthSystemSizer.Add(self.space3SystemText, 0, wx.ALL, 5)
        fourthSystemSizer.Add(self.listWithNotUsedElementsLabel, 0, wx.ALL, 5)
        fourthSystemSizer.Add(self.listWithNotUsedElements, 0, wx.ALL, 5)

        firstOutputSizer.Add(self.outputIcon, 0, wx.ALL, 5)
        firstOutputSizer.Add(self.outputText, 0, wx.ALL, 5)
        firstOutputSizer.Add(self.nameColumnWriteLabel, 0, wx.ALL, 5)
        firstOutputSizer.Add(self.nameColumnWrite, 0, wx.ALL, 5)
        firstOutputSizer.Add(self.typeColumnWriteLabel, 0, wx.ALL, 5)
        firstOutputSizer.Add(self.typeColumnWrite, 0, wx.ALL, 5)
        firstOutputSizer.Add(self.countColumnWriteLabel, 0, wx.ALL, 5)
        firstOutputSizer.Add(self.countColumnWrite, 0, wx.ALL, 5)
        firstOutputSizer.Add(self.valueColumnWriteLabel, 0, wx.ALL, 5)
        firstOutputSizer.Add(self.valueColumnWrite, 0, wx.ALL, 5)
        firstOutputSizer.Add(self.rowStartWriteLabel, 0, wx.ALL, 5)
        firstOutputSizer.Add(self.rowStartWrite, 0, wx.ALL, 5)

        secondOutputSizer.Add(self.space1OutputText, 0, wx.ALL, 5)
        secondOutputSizer.Add(self.writeColumnAddressLabel, 0, wx.ALL, 5)
        secondOutputSizer.Add(self.writeColumnAddress, 0, wx.ALL, 5)
        secondOutputSizer.Add(self.writeRowAddressLabel, 0, wx.ALL, 5)
        secondOutputSizer.Add(self.writeRowAddress, 0, wx.ALL, 5)
        secondOutputSizer.Add(self.writeL2ArchitectLabel, 0, wx.ALL, 5)
        secondOutputSizer.Add(self.writeL2Architect, 0, wx.ALL, 5)
        secondOutputSizer.Add(self.writeTreeLevel2Label, 0, wx.ALL, 5)
        secondOutputSizer.Add(self.writeTreeLevel2, 0, wx.ALL, 5)
        secondOutputSizer.Add(self.columnLowLimitSystem, 0, wx.ALL, 5)
        secondOutputSizer.Add(self.lowLimitColumnSystem, 0, wx.ALL, 5)
        secondOutputSizer.Add(self.columnMaxLimitSystem, 0, wx.ALL, 5)
        secondOutputSizer.Add(self.maxLimitColumnSystem, 0, wx.ALL, 5)

        thirdOutputSizer.Add(self.space2OutputText, 0, wx.ALL, 5)
        thirdOutputSizer.Add(self.columnResLabel, 0, wx.ALL, 5)
        thirdOutputSizer.Add(self.columnRes, 0, wx.ALL, 5)
        thirdOutputSizer.Add(self.rowResLabel, 0, wx.ALL, 5)
        thirdOutputSizer.Add(self.rowRes, 0, wx.ALL, 5)
        thirdOutputSizer.Add(self.rowWriteA2lLabel, 0, wx.ALL, 5)
        thirdOutputSizer.Add(self.rowWriteA2l, 0, wx.ALL, 5)
        thirdOutputSizer.Add(self.columnWriteA2lLabel, 0, wx.ALL, 5)
        thirdOutputSizer.Add(self.columnWriteA2l, 0, wx.ALL, 5)
        thirdOutputSizer.Add(self.columnWriteSheet2Label, 0, wx.ALL, 5)
        thirdOutputSizer.Add(self.columnWriteSheet2, 0, wx.ALL, 5)
        thirdOutputSizer.Add(self.rowWriteSheet2Label, 0, wx.ALL, 5)
        thirdOutputSizer.Add(self.rowWriteSheet2, 0, wx.ALL, 5)

        firstPathSizer.Add(self.architectPathIcon, 0, wx.ALL, 5)
        firstPathSizer.Add(self.inputArchitectPathLabel, 0, wx.ALL, 5)
        firstPathSizer.Add(self.inputArchitectPath, 0, wx.ALL, 5)
        firstPathSizer.Add(buttonInputArchitect, 0, wx.ALL, 5)
        firstPathSizer.Add(self.sheet1Text, 0, wx.ALL, 5)
        firstPathSizer.Add(self.sheet1Input, 0, wx.ALL, 5)

        secondPathSizer.Add(self.systemPathIcon, 0, wx.ALL, 5)
        secondPathSizer.Add(self.inputSystemPathLabel, 0, wx.ALL, 5)
        secondPathSizer.Add(self.inputSystemPath, 0, wx.ALL, 5)
        secondPathSizer.Add(buttonInputSystem, 0, wx.ALL, 5)
        secondPathSizer.Add(self.sheet2Text, 0, wx.ALL, 5)
        secondPathSizer.Add(self.sheet2Input, 0, wx.ALL, 5)

        thirdPathSizer.Add(self.outputPathIcon, 0, wx.ALL, 5)
        thirdPathSizer.Add(self.OutputPathLabel, 0, wx.ALL, 5)
        thirdPathSizer.Add(self.OutputPath, 0, wx.ALL, 5)
        thirdPathSizer.Add(buttonOutputPath, 0, wx.ALL, 5)
        thirdPathSizer.Add(self.sheet3Text, 0, wx.ALL, 5)
        thirdPathSizer.Add(self.sheet3Input, 0, wx.ALL, 5)
        thirdPathSizer.Add(self.sheet4Text, 0, wx.ALL, 5)
        thirdPathSizer.Add(self.sheet4Input, 0, wx.ALL, 5)

        buttonSizer.Add(self.buttonRun, 0, wx.ALL, 5)
        buttonSizer.Add(cancelBtn, 0, wx.ALL, 5)

        loadAndSaveSizer.Add(buttonLoad, 0, wx.ALL, 5)
        loadAndSaveSizer.Add(buttonSave, 0, wx.ALL, 5)

        architectSizer.Add(firstArhitectSizer, 0, wx.ALL | wx.EXPAND, 5)
        architectSizer.Add(secondArhitectSizer, 0, wx.ALL | wx.EXPAND, 5)
        architectSizer.Add(thirdArchitectSizer, 0, wx.ALL | wx.EXPAND, 5)
        oemSizer.Add(firstOemSizer, 0, wx.ALL | wx.EXPAND, 5)
        oemSizer.Add(secondOemSizer, 0, wx.ALL | wx.EXPAND, 5)
        systemSizer.Add(firstSystemSizer, 0, wx.ALL | wx.EXPAND, 5)
        systemSizer.Add(secondSystemSizer, 0, wx.ALL | wx.EXPAND, 5)
        systemSizer.Add(fourthSystemSizer, 0, wx.ALL | wx.EXPAND, 5)
        outputSizer.Add(firstOutputSizer, 0, wx.ALL | wx.EXPAND, 5)
        outputSizer.Add(secondOutputSizer, 0, wx.ALL | wx.EXPAND, 5)
        outputSizer.Add(thirdOutputSizer, 0, wx.ALL | wx.EXPAND, 5)

        pathSizer.Add(firstPathSizer, 0, wx.ALL | wx.EXPAND, 5)
        pathSizer.Add(secondPathSizer, 0, wx.ALL | wx.EXPAND, 5)
        pathSizer.Add(thirdPathSizer, 0, wx.ALL | wx.EXPAND, 5)

        mainSizer.Add(self.m_bitmap3, 1, wx.EXPAND, 0)
        mainSizer.Add(wx.StaticLine(self, ), 0, wx.ALL | wx.EXPAND, 5)
        mainSizer.Add(pathSizer, 0, wx.ALL | wx.EXPAND, 5)
        mainSizer.Add(wx.StaticLine(self, ), 0, wx.ALL | wx.EXPAND, 5)
        mainSizer.Add(architectSizer, 0, wx.ALL | wx.EXPAND, 5)
        mainSizer.Add(wx.StaticLine(self, ), 0, wx.ALL | wx.EXPAND, 5)
        mainSizer.Add(oemSizer, 0, wx.ALL | wx.EXPAND, 5)
        mainSizer.Add(wx.StaticLine(self, ), 0, wx.ALL | wx.EXPAND, 5)
        mainSizer.Add(systemSizer, 0, wx.ALL | wx.EXPAND, 5)
        mainSizer.Add(wx.StaticLine(self, ), 0, wx.ALL | wx.EXPAND, 5)
        mainSizer.Add(outputSizer, 0, wx.ALL | wx.EXPAND, 5)
        mainSizer.Add(wx.StaticLine(self, ), 0, wx.ALL | wx.EXPAND, 5)

        mainSizer.Add(loadAndSaveSizer, 0, wx.CENTER)
        mainSizer.Add(wx.StaticLine(self, ), 0, wx.ALL | wx.EXPAND, 5)
        mainSizer.Add(buttonSizer, 0, wx.CENTER)

        self.SetupScrolling()
        self.SetSizer(mainSizer)
        mainSizer.Fit(self)
        self.Layout()

    def onMouseOver(self, event):    # pragma: no cover

        # Architect Hover Section

        self.architectText.SetToolTip("The section where are the settings for the Architect excel.")
        self.arhitectIcon.SetToolTip("The section where are the settings for the Architect excel.")
        self.columnNameArchitect.SetToolTip("Set the column to read names of the parameters from Architect excel.")
        self.columnTypeArchitect.SetToolTip("Set the column to read types of the parameters from Architect excel.")
        self.columnCountArchitect.SetToolTip("Set the column to read count number of "
                                             "the parameters from Architect excel.")
        self.columnValueArchitect.SetToolTip("Set the column to read values of the parameters from Architect excel.")
        self.rowStartArchitect.SetToolTip("Set the row number where the program\nwill start reading data "
                                          "from Architect excel.")
        self.rowEndArchitect.SetToolTip("Set the row parameter to stop reading lines from Architect excel."
                                        "\nThe program will read the Architect excel until it reaches this parameter.")

        self.readRowStartAddressArchitect.SetToolTip("Set the row from where the program will read\nthe first address in decimal from Architect excel.")
        self.readColumnStartAddressArchitect.SetToolTip("Set the column from where the program will read\nthe first address in decimal from Architect excel.")
        self.resColumnArchitect.SetToolTip("This is a special column used only for the reserved values "
                                           "because sometimes the reserved value"
                                           " doesn't have name,type,count or value.\n"
                                           "When a number is in this column the program is creating automatically the "
                                           "reserved parameter.\nIn generally this column is the right side of the "
                                           "total size from Architect excel.")

        self.rowA2lLabel.SetToolTip("Set the starting row number where the program\nwill start reading A2L names "
                                    "from Architect excel.")
        self.columnA2lLabel.SetToolTip("Set the column number where the program\nwill start reading A2L names "
                                       "from Architect excel.")

        self.defaultColumnLabel.SetToolTip("Set the column number where the program\nwill read the default numbers\nIf the value from the selected project is None, "
                                           "the program will take this value instead.")

        # OEM Hover Section

        self.oemText.SetToolTip("The section where are the settings for OEM.")
        self.oemIcon.SetToolTip("The section where are the settings for OEM.")
        self.oemInput.SetToolTip("Check the box if the project has OEM.")
        self.projectOemName.SetToolTip("Set the name of the project.")
        self.rowStartOem.SetToolTip("Set the row number where the program\nwill start reading data "
                                    "from the selected OEM sheet.")
        self.rowEndOem.SetToolTip("Set the row parameter to stop reading lines from the selected OEM sheet."
                                  "\nThe program will read the selected OEM sheet until it reaches this parameter.")
        self.columnNameOem.SetToolTip("Set the column to read names of the parameters from the selected OEM sheet.")
        self.columnCountOem.SetToolTip("Set the column to read count number of the parameters"
                                       " from the selected OEM sheet.")
        self.columnValueOem.SetToolTip("Set the column to read values of the parameters from the selected OEM sheet.")

        self.rowOemA2lLabel.SetToolTip("Set the starting row number where the program\nwill start reading A2L names "
                                       "from the selected OEM sheet.")
        self.columnOemA2lLabel.SetToolTip("Set the column number where the program\nwill start reading A2L names "
                                          "from the selected OEM sheet.")

        # System Hover Section

        self.systemText.SetToolTip("The section where are the settings for the System excel.")
        self.systemIcon.SetToolTip("The section where are the settings for the System excel.")
        self.columnNameSystem.SetToolTip("Set the column to read names of the parameters from System excel.")
        self.columnCountSystem.SetToolTip("Set the column to read count number of the parameters from System excel.")
        self.columnLimitValueSystem.SetToolTip("Set the column to read limits of the parameters from System excel.")
        self.columnValueSystem.SetToolTip("Set the column to read values of the parameters from System excel.")
        self.rowStartSystemLabel.SetToolTip("Set the row number where the program\nwill start reading data "
                                            "from System excel.")
        self.rowEndSystemLabel.SetToolTip("Set the row parameter to stop reading lines from the selected OEM sheet."
                                          "\nThe program will read the System excel until it reaches this parameter.")
        self.projectSpecificNameLabel.SetToolTip("Set the complete name of the selected project.\nThis name is used to find specific values from the System excel.")
        self.projectNameTypeLabel.SetToolTip("Set type of the project from the System excel.\nAt this moment you can choose between default ARS or default SRR")
        self.listWithNotUsedElementsLabel.SetToolTip("Set the names of the parameters that doesn't need update from the System excel."
                                                     " These parameters will have the values from the Architect excel!")

        # Output Hover Section

        self.outputText.SetToolTip("The section where are the settings for the Output excel.")
        self.outputIcon.SetToolTip("The section where are the settings for the Output excel.")
        self.nameColumnWriteLabel.SetToolTip("Set the column to write names of the parameters from Output excel.")
        self.typeColumnWriteLabel.SetToolTip("Set the column to write type of the parameters from Output excel.")
        self.countColumnWriteLabel.SetToolTip("Set the column to write count of the parameters from Output excel.")
        self.valueColumnWriteLabel.SetToolTip("Set the column to write values of the parameters from Output excel.")
        self.rowStartWriteLabel.SetToolTip("Set the row number where the program will start writing data.")
        self.writeRowAddressLabel.SetToolTip("Set the row number where the program will start writing addresses.")
        self.writeColumnAddressLabel.SetToolTip("Set the column number where the program will write addresses.")
        self.writeL2ArchitectLabel.SetToolTip("Set the L2 Architect column number.")
        self.writeTreeLevel2Label.SetToolTip("Set the Tree Level 2 column number.")
        self.columnResLabel.SetToolTip("Set the column number where the reserved parameters will be writen.\n"
                                       "These parameters will be writen on the second sheet named \"add_PPAR_default_data\".")
        self.rowResLabel.SetToolTip("Set the starting row number where the reserved parameters will be writen.\n"
                                    "These parameters will be writen on the second sheet named \"add_PPAR_default_data\".")
        self.rowWriteA2lLabel.SetToolTip("Set the starting row number where the A2L parameters will be writen.")
        self.columnWriteA2lLabel.SetToolTip("Set the column number where the A2L parameters will be writen.")

        self.columnWriteSheet2Label.SetToolTip("Set the column number where the parameters that doesn't have enough space in the first sheet will be writen separately.\n"
                                               "These parameters will be writen on the second sheet named \"add_PPAR_default_data\".")
        self.rowWriteSheet2Label.SetToolTip("Set the starting row number where the parameters that doesn't have enough space in the first sheet will be writen separately.\n"
                                            "These parameters will be writen on the second sheet named \"add_PPAR_default_data\".")
        self.columnLowLimitSystem.SetToolTip("Set the column number where the parameter that have limits from System excel will be writen. On this column will be writen the low limit.")
        self.columnMaxLimitSystem.SetToolTip("Set the column number where the parameter that have limits from System excel will be writen. On this column will be writen the max limit.")

        # Path Hover Section

        self.inputArchitectPathLabel.SetToolTip("Press the \"Browse\" button and select the Architect Excel")
        self.architectPathIcon.SetToolTip("Press the \"Browse\" button and select the Architect Excel")
        self.inputSystemPathLabel.SetToolTip("Press the \"Browse\" button and select the System Excel")
        self.systemPathIcon.SetToolTip("Press the \"Browse\" button and select the System Excel")
        self.OutputPathLabel.SetToolTip("Press the \"Browse\" button and select the Output Excel")
        self.outputPathIcon.SetToolTip("Press the \"Browse\" button and select the Output Excel")

        # Sheet Hover Section

        self.sheet1Text.SetToolTip("The program will read from this sheet the information from the Architect Excel")
        self.sheet2Text.SetToolTip("The program will read from this sheet the information from the System Excel")
        self.sheet3Text.SetToolTip("The program will write in this sheet the information from the Excels")
        self.sheet4Text.SetToolTip("The program will write in this sheet the additional information from the Excels")

    def onLoadSetting(self, event):  # pragma: no cover
        pathTxt = self.ShowDialogLoad()
        if pathTxt == "Path":
            dlg = wx.RichMessageDialog(parent=None, message="No txt file selected!")
            dlg.ShowModal()
        else:

            parser = configparser.ConfigParser()
            parser.read(pathTxt)

            try:
                sectionConfigFileArchitect = "Architect"
                self.nameColumnArchitect.SetValue(parser.getint(sectionConfigFileArchitect, "nameColumn"))
                self.typeColumnArchitect.SetValue(parser.getint(sectionConfigFileArchitect, "typeColumn"))
                self.countColumnArchitect.SetValue(parser.getint(sectionConfigFileArchitect, "countColumn"))
                self.valueColumnArchitect.SetValue(parser.getint(sectionConfigFileArchitect, "valueColumn"))
                self.startRowArchitect.SetValue(parser.getint(sectionConfigFileArchitect, "rowStart"))
                self.endRowArchitect.SetValue(parser.get(sectionConfigFileArchitect, "rowEnd"))
                self.rowStartReadAddressArchitect.SetValue(parser.getint(sectionConfigFileArchitect, "firstAdressRow"))
                self.columnStartReadAddressArchitect.SetValue(parser.getint(sectionConfigFileArchitect, "firstAdressColumn"))
                self.columnResArchitect.SetValue(parser.getint(sectionConfigFileArchitect, "resColumn"))
                self.rowA2l.SetValue(parser.getint(sectionConfigFileArchitect, "rowa2l"))
                self.columnA2l.SetValue(parser.getint(sectionConfigFileArchitect, "columna2l"))
                self.defaultColumn.SetValue(parser.getint(sectionConfigFileArchitect, "defaultColumn"))

            except Exception as e:
                dlg = wx.RichMessageDialog(parent=None, message="Error at loading from Architect section.\nCheck if the variable " + e.args[0] + " exists in the loaded file.\n")
                dlg.ShowModal()
                logging.warning(e)
            else:
                try:
                    sectionConfigOem = "ArchitectOEM"

                    valid = True
                    global_.projectOemName = parser.get(sectionConfigOem, "projectOemName")
                    for dropBoxItem in self.oemProjectName.GetItems():
                        if global_.projectOemName in dropBoxItem:
                            valid = False

                    if valid is True:
                        self.oemProjectName.Append(parser.get(sectionConfigOem, "projectOemName"))

                    self.oemProjectName.SetStringSelection(global_.projectOemName)
                    self.startRowOem.SetValue(parser.getint(sectionConfigOem, "rowStartOem"))
                    self.endRowOem.SetValue(parser.get(sectionConfigOem, "rowEndOem"))
                    self.nameColumnOem.SetValue(parser.getint(sectionConfigOem, "nameColumnOem"))
                    self.countColumnOem.SetValue(parser.getint(sectionConfigOem, "countColumnOem"))
                    self.valueColumnOem.SetValue(parser.getint(sectionConfigOem, "valueColumnOem"))
                    self.rowOemA2l.SetValue(parser.getint(sectionConfigOem, "rowOema2l"))
                    self.columnOemA2l.SetValue(parser.getint(sectionConfigOem, "columnOema2l"))

                except Exception as e:
                    dlg = wx.RichMessageDialog(parent=None, message="Error at loading from Oem section.\nCheck if the variable " + e.args[0] + " exists in the loaded file.\n")
                    dlg.ShowModal()
                    logging.error(e.args[0], exc_info=True)
                else:
                    try:
                        sectionConfigFileSystem = "SystemExcel"
                        self.nameColumnSystem.SetValue(parser.getint(sectionConfigFileSystem, "nameColumn2"))
                        self.typeColumnSystem.SetValue(parser.getint(sectionConfigFileSystem, "typeColumn2"))
                        self.countColumnSystem.SetValue(parser.getint(sectionConfigFileSystem, "countColumn2"))
                        self.limitValueColumnSystem.SetValue(parser.getint(sectionConfigFileSystem, "typeOfLimitColumn"))
                        self.valueColumnSystem.SetValue(parser.getint(sectionConfigFileSystem, "valueColumn2"))
                        self.rowStartSystem.SetValue(parser.getint(sectionConfigFileSystem, "rowStart2"))
                        self.rowEndSystem.SetValue(parser.get(sectionConfigFileSystem, "rowEnd2"))
                        self.projectSpecificName.SetValue(parser.get(sectionConfigFileSystem, "projectSpecificName"))
                        self.projectNameType.SetValue(parser.get(sectionConfigFileSystem, "projectNameType"))
                        self.columnWriteSheet2.SetValue(parser.getint(sectionConfigFileSystem,
                                                                      "columnForOutputAdd_PPAR_default_data"))
                        self.rowWriteSheet2.SetValue(parser.getint(sectionConfigFileSystem, "rowForOutputAdd_PPAR_default_data"))
                        self.listWithNotUsedElements.SetValue(parser.get(sectionConfigFileSystem, "listWithNotUsedElements"))
                    except Exception as e:
                        dlg = wx.RichMessageDialog(parent=None, message="Error at loading from System section.\nCheck if the variable " + e.args[0] + " exists in the loaded file.\n")
                        dlg.ShowModal()
                        logging.error(e.args[0], exc_info=True)
                    else:
                        try:
                            sectionConfigFileWrite = "OutputExcel"
                            self.rowWriteA2l.SetValue(parser.getint(sectionConfigFileWrite, "rowWritea2l"))
                            self.columnWriteA2l.SetValue(parser.getint(sectionConfigFileWrite, "columnWritea2l"))
                            self.nameColumnWrite.SetValue(parser.getint(sectionConfigFileWrite, "nameColumnWrite"))
                            self.typeColumnWrite.SetValue(parser.getint(sectionConfigFileWrite, "typeColumnWrite"))
                            self.countColumnWrite.SetValue(parser.getint(sectionConfigFileWrite, "countColumnWrite"))
                            self.valueColumnWrite.SetValue(parser.getint(sectionConfigFileWrite, "valueColumnWrite"))
                            self.rowStartWrite.SetValue(parser.getint(sectionConfigFileWrite, "rowStartWrite"))
                            self.writeRowAddress.SetValue(parser.getint(sectionConfigFileWrite, "writeRowAdress"))
                            self.writeColumnAddress.SetValue(parser.getint(sectionConfigFileWrite, "writeColumnAdress"))
                            self.writeL2Architect.SetValue(parser.getint(sectionConfigFileWrite, "writeL2Architect"))
                            self.writeTreeLevel2.SetValue(parser.getint(sectionConfigFileWrite, "writeTreeLevel2"))
                            self.columnRes.SetValue(parser.getint(sectionConfigFileWrite, "columnRes"))
                            self.rowRes.SetValue(parser.getint(sectionConfigFileWrite, "rowRes"))
                            self.lowLimitColumnSystem.SetValue(parser.getint(sectionConfigFileWrite, "lowLimitColumnFromOutputExcel"))
                            self.maxLimitColumnSystem.SetValue(parser.getint(sectionConfigFileWrite, "maxLimitColumnFromOutputExcel"))
                        except Exception as e:
                            dlg = wx.RichMessageDialog(parent=None, message="Error at loading from Output section.\nCheck if the variable " + e.args[0] + " exists in the loaded file.\n")
                            dlg.ShowModal()
                            logging.error(e.args[0], exc_info=True)

    def onSaveSetting(self, event):  # pragma: no cover
        pathTxt = self.ShowDialogSave()
        if pathTxt == "Path":
            dlg = wx.RichMessageDialog(parent=None, message="No txt file selected!")
            dlg.ShowModal()
        else:
            with open(pathTxt, "w") as file:
                try:
                    file.write("[Architect]\n")
                    file.write("nameColumn=" + str(self.nameColumnArchitect.GetValue()) + "\n")
                    file.write("typeColumn=" + str(self.typeColumnArchitect.GetValue()) + "\n")
                    file.write("countColumn=" + str(self.countColumnArchitect.GetValue()) + "\n")
                    file.write("valueColumn=" + str(self.valueColumnArchitect.GetValue()) + "\n")
                    file.write("resColumn=" + str(self.columnResArchitect.GetValue()) + "\n")
                    file.write("rowStart=" + str(self.startRowArchitect.GetValue()) + "\n")
                    file.write("rowEnd=" + str(self.endRowArchitect.GetValue()) + "\n")
                    file.write("firstAdressRow=" + str(self.rowStartReadAddressArchitect.GetValue()) + "\n")
                    file.write("firstAdressColumn=" + str(self.columnStartReadAddressArchitect.GetValue()) + "\n")
                    file.write("rowa2l=" + str(self.rowA2l.GetValue()) + "\n")
                    file.write("columna2l=" + str(self.columnA2l.GetValue()) + "\n")
                    file.write("defaultColumn=" + str(self.defaultColumn.GetValue()) + "\n")
                except Exception as e:
                    dlg = wx.RichMessageDialog(parent=None, message="Error at saving from Architect section.\nCheck the variable " + e.args[0] + ".")
                    dlg.ShowModal()
                    logging.error(e.args[0], exc_info=True)
                else:
                    try:
                        file.write("\n\n")
                        file.write("[ArchitectOEM]\n")
                        if global_.checkOem is True:
                            file.write("projectOemName=" + str(self.oemProjectName.GetValue()) + "\n\n")
                            file.write("nameColumnOem=" + str(self.nameColumnOem.GetValue()) + "\n")
                            file.write("countColumnOem=" + str(self.countColumnOem.GetValue()) + "\n")
                            file.write("valueColumnOem=" + str(self.valueColumnOem.GetValue()) + "\n")
                            file.write("rowStartOem=" + str(self.startRowOem.GetValue()) + "\n")
                            file.write("rowEndOem=" + str(self.endRowOem.GetValue()) + "\n")
                            file.write("rowOema2l=" + str(self.rowOemA2l.GetValue()) + "\n")
                            file.write("columnOema2l=" + str(self.columnOemA2l.GetValue()) + "\n")
                        else:
                            file.write("projectOemName=" + "\n\n")
                            file.write("nameColumnOem=0" + "\n")
                            file.write("countColumnOem=0" + "\n")
                            file.write("valueColumnOem=0" + "\n")
                            file.write("rowStartOem=0" + "\n")
                            file.write("rowEndOem=" + "\n")
                            file.write("rowOema2l=0" + "\n")
                            file.write("columnOema2l=0" + "\n")

                    except Exception as e:
                        dlg = wx.RichMessageDialog(parent=None, message="Error at saving from Oem section.\nCheck the variable " + e.args[0] + ".")
                        dlg.ShowModal()
                        logging.error(e.args[0], exc_info=True)
                    else:
                        try:
                            file.write("\n\n")
                            file.write("[SystemExcel]\n")
                            file.write("projectSpecificName=" + str(self.projectSpecificName.GetValue()) + "\n")
                            file.write("projectNameType=" + str(self.projectNameType.GetValue()) + "\n")
                            file.write("listWithNotUsedElements=" + str(self.listWithNotUsedElements.GetValue()) + "\n\n")

                            file.write("nameColumn2=" + str(self.nameColumnSystem.GetValue()) + "\n")
                            file.write("typeColumn2=" + str(self.typeColumnSystem.GetValue()) + "\n")
                            file.write("countColumn2=" + str(self.countColumnSystem.GetValue()) + "\n")
                            file.write("valueColumn2=" + str(self.valueColumnSystem.GetValue()) + "\n")
                            file.write("typeOfLimitColumn=" + str(self.limitValueColumnSystem.GetValue()) + "\n")
                            file.write("rowStart2=" + str(self.rowStartSystem.GetValue()) + "\n")
                            file.write("rowEnd2=" + str(self.rowEndSystem.GetValue()) + "\n")
                            file.write("columnForOutputAdd_PPAR_default_data=" + str(self.columnWriteSheet2.GetValue()) + "\n")
                            file.write("rowForOutputAdd_PPAR_default_data=" + str(self.rowWriteSheet2.GetValue()) + "\n")
                        except Exception as e:
                            dlg = wx.RichMessageDialog(parent=None, message="Error at saving from System section.\nCheck the variable " + e.args[0] + ".")
                            dlg.ShowModal()
                            logging.error(e.args[0], exc_info=True)
                        else:
                            try:
                                file.write("\n\n")
                                file.write("[OutputExcel]\n")
                                file.write("nameColumnWrite=" + str(self.nameColumnWrite.GetValue()) + "\n")
                                file.write("typeColumnWrite=" + str(self.typeColumnWrite.GetValue()) + "\n")
                                file.write("countColumnWrite=" + str(self.countColumnWrite.GetValue()) + "\n")
                                file.write("valueColumnWrite=" + str(self.valueColumnWrite.GetValue()) + "\n")
                                file.write("rowStartWrite=" + str(self.rowStartWrite.GetValue()) + "\n")
                                file.write("writeRowAdress=" + str(self.writeRowAddress.GetValue()) + "\n")
                                file.write("writeColumnAdress=" + str(self.writeColumnAddress.GetValue()) + "\n")
                                file.write("writeL2Architect=" + str(self.writeL2Architect.GetValue()) + "\n")
                                file.write("writeTreeLevel2=" + str(self.writeTreeLevel2.GetValue()) + "\n")
                                file.write("columnRes=" + str(self.columnRes.GetValue()) + "\n")
                                file.write("rowRes=" + str(self.rowRes.GetValue()) + "\n")
                                file.write("rowWritea2l=" + str(self.rowWriteA2l.GetValue()) + "\n")
                                file.write("columnWritea2l=" + str(self.columnWriteA2l.GetValue()) + "\n")
                                file.write("lowLimitColumnFromOutputExcel=" + str(self.lowLimitColumnSystem.GetValue()) + "\n")
                                file.write("maxLimitColumnFromOutputExcel=" + str(self.maxLimitColumnSystem.GetValue()) + "\n\n")
                            except Exception as e:
                                dlg = wx.RichMessageDialog(parent=None, message="Error at saving from Output section.\nCheck the variable " + e.args[0] + ".")
                                dlg.ShowModal()
                                logging.error(e.args[0], exc_info=True)

    def ShowDialogSave(self):    # pragma: no cover
        ExcelPath = "Path"
        dlg = wx.FileDialog(
            self, message="Save file as ...",
            defaultDir="",
            defaultFile="", wildcard="*.txt", style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
        )
        if dlg.ShowModal() == wx.ID_OK:
            ExcelPath = dlg.GetPath()
            with open(ExcelPath, 'w'):
                pass
        dlg.Destroy()
        return ExcelPath

    def ShowDialogLoad(self):  # Dialog box when we press the Load button.   # pragma: no cover
        ExcelPath = "Path"
        dlg = wx.FileDialog(self, "Choose a txt file",
                            defaultDir="",
                            defaultFile="",
                            wildcard="*.txt")
        if dlg.ShowModal() == wx.ID_OK:
            ExcelPath = dlg.GetPath()
        dlg.Destroy()
        return ExcelPath

    def onCheckBox(self, event):  # Getting information when the box is checked or not.  # pragma: no cover
        if not self.oemInput.GetValue():
            self.oemProjectName.Enable(False)
            self.startRowOem.Enable(False)
            self.endRowOem.Enable(False)
            self.nameColumnOem.Enable(False)
            self.countColumnOem.Enable(False)
            self.valueColumnOem.Enable(False)
            self.rowOemA2l.Enable(False)
            self.columnOemA2l.Enable(False)

        else:
            self.oemProjectName.Enable(True)
            self.startRowOem.Enable(True)
            self.endRowOem.Enable(True)
            self.nameColumnOem.Enable(True)
            self.countColumnOem.Enable(True)
            self.valueColumnOem.Enable(True)
            self.rowOemA2l.Enable(True)
            self.columnOemA2l.Enable(True)

    def ShowDialogExcel(self):  # Dialog box when we press the Browse button for an Excel.   # pragma: no cover
        ExcelPath = "Path"
        dlg = wx.FileDialog(self, "Choose an Excel file",
                            defaultDir="",
                            defaultFile="",
                            wildcard="*.xlsx;*.xlsm;*.xls")

        if dlg.ShowModal() == wx.ID_OK:
            ExcelPath = dlg.GetPath()
        dlg.Destroy()
        return ExcelPath

    @staticmethod
    def get_sheet_ids(file_path):    # pragma: no cover
        sheet_names = []
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            xml = zip_ref.open(r'xl/workbook.xml').read()
            dictionary = xmltodict.parse(xml)

            if not isinstance(dictionary['workbook']['sheets']['sheet'], list):
                sheet_names.append(dictionary['workbook']['sheets']['sheet']['@name'])
            else:
                for sheet in dictionary['workbook']['sheets']['sheet']:
                    sheet_names.append(sheet['@name'])
        return sheet_names

    def onButtonReadPathArchitect(self, event):  # Getting the Limits Excel path     # pragma: no cover
        global_.pathArchitect = self.ShowDialogExcel()
        self.inputArchitectPath.SetValue(global_.pathArchitect)
        try:
            wx.ComboBox.SetItems(self.oemProjectName, self.get_sheet_ids(global_.pathArchitect))
            wx.ComboBox.SetItems(self.sheet1Input, self.get_sheet_ids(global_.pathArchitect))
            self.oemProjectName.SetStringSelection(global_.projectOemName)
            self.sheet1Input.SetStringSelection(global_.architectSheet)
            self.oemProjectName.Append("")
        except:
            pass

    def onButtonReadPathSystem(self, event):  # Getting the CSV path     # pragma: no cover
        global_.pathSystem = self.ShowDialogExcel()
        self.inputSystemPath.SetValue(global_.pathSystem)
        try:
            wx.ComboBox.SetItems(self.sheet2Input, self.get_sheet_ids(global_.pathSystem))
            self.sheet2Input.SetStringSelection(global_.systemSheet)
        except:
            pass

    def onButtonReadPathOutCompare(self, event):  # Getting the Output Excel path    # pragma: no cover
        global_.pathOutput = self.ShowDialogExcel()
        self.OutputPath.SetValue(global_.pathOutput)
        try:
            wx.ComboBox.SetItems(self.sheet3Input, self.get_sheet_ids(global_.pathOutput))
            wx.ComboBox.SetItems(self.sheet4Input, self.get_sheet_ids(global_.pathOutput))
            self.sheet3Input.SetStringSelection(global_.output1Sheet)
            self.sheet4Input.SetStringSelection(global_.output2Sheet)
        except:
            pass

    def my_message(self, exception_value, exception_traceback):  # pragma: no cover
        msg = "Oh no! An error has occurred.\n\n"  # Function used to PopUp a message when an error occurred.
        tb = traceback.format_exception(self, exception_value, exception_traceback)
        for iterator in tb:
            msg += iterator
        dlg = wx.MessageDialog(None, msg, str(self), wx.OK | wx.ICON_INFORMATION)
        dlg.ShowModal()
        dlg.Destroy()

    sys.excepthook = my_message

    @staticmethod
    def OutputDialog():  # Dialog box when the program finished execution.   # pragma: no cover
        dlg = wx.MessageDialog(None, global_.stateOfTheProgram + "\nDo you want to open the output file?", 'Output',
                               wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal()
        if result == wx.ID_YES:
            os.startfile(global_.pathOutput)

    def configVariables(self):       # pragma: no cover
        # ARCHITECT
        global_.nameColumn = (self.nameColumnArchitect.GetValue())
        global_.typeColumn = (self.typeColumnArchitect.GetValue())
        global_.countColumn = (self.countColumnArchitect.GetValue())
        global_.valueColumn = (self.valueColumnArchitect.GetValue())
        global_.resColumn = (self.columnResArchitect.GetValue())
        global_.row = (self.startRowArchitect.GetValue())
        global_.rowEnd = (self.endRowArchitect.GetValue())
        global_.readRowAddress = (self.rowStartReadAddressArchitect.GetValue())
        global_.readColumnAddress = (self.columnStartReadAddressArchitect.GetValue())
        global_.rowA2L = (self.rowA2l.GetValue())
        global_.columnA2L = (self.columnA2l.GetValue())
        global_.defaultColumn = self.defaultColumn.GetValue()

        # SYSTEM
        global_.nameColumn2 = (self.nameColumnSystem.GetValue())
        global_.typeColumn2 = (self.typeColumnSystem.GetValue())
        global_.countColumn2 = (self.countColumnSystem.GetValue())
        global_.valueColumn2 = (self.valueColumnSystem.GetValue())
        global_.row2 = (self.rowStartSystem.GetValue())
        global_.rowEnd2 = (self.rowEndSystem.GetValue())
        global_.limitValueColumn = (self.limitValueColumnSystem.GetValue())
        global_.projectSpecificName = (self.projectSpecificName.GetValue())
        global_.projectNameType = (self.projectNameType.GetValue())
        global_.listWithNotUsedElements = (self.listWithNotUsedElements.GetValue())
        global_.columnWriteSheet2 = (self.columnWriteSheet2.GetValue())
        global_.rowWriteSheet2 = (self.rowWriteSheet2.GetValue())

        # OUTPUT
        global_.nameColumnWrite = (self.nameColumnWrite.GetValue())
        global_.typeColumnWrite = (self.typeColumnWrite.GetValue())
        global_.countColumnWrite = (self.countColumnWrite.GetValue())
        global_.valueColumnWrite = (self.valueColumnWrite.GetValue())
        global_.rowWrite = (self.rowStartWrite.GetValue())
        global_.writeL2Architect = (self.writeL2Architect.GetValue())
        global_.writeTreeLevel2 = (self.writeTreeLevel2.GetValue())
        global_.columnRes = (self.columnRes.GetValue())
        global_.rowRes = (self.rowRes.GetValue())
        global_.writeRowAddress = (self.writeRowAddress.GetValue())
        global_.writeColumnAddress = (self.writeColumnAddress.GetValue())
        global_.rowWriteA2L = (self.rowWriteA2l.GetValue())
        global_.columnWriteA2L = (self.columnWriteA2l.GetValue())
        global_.lowLimitColumn = (self.lowLimitColumnSystem.GetValue())
        global_.maxLimitColumn = (self.maxLimitColumnSystem.GetValue())

        # OEM
        if self.oemInput.GetValue():
            global_.projectOemName = (self.oemProjectName.GetValue())
            global_.nameColumnOem = (self.nameColumnOem.GetValue())
            global_.countColumnOem = (self.countColumnOem.GetValue())
            global_.valueColumnOem = (self.valueColumnOem.GetValue())
            global_.rowOem = (self.startRowOem.GetValue())
            global_.rowEndOem = (self.endRowOem.GetValue())
            global_.rowOEMa2l = (self.rowOemA2l.GetValue())
            global_.columnOemA2L = (self.columnOemA2l.GetValue())
        else:
            global_.projectOemName = ""
            global_.nameColumnOem = 0
            global_.countColumnOem = 0
            global_.valueColumnOem = 0
            global_.rowOem = 0
            global_.rowEndOem = ""
            global_.rowOEMa2l = 0
            global_.columnOemA2L = 0

        # SHEETS
        global_.systemSheet = (self.sheet2Input.GetValue())
        global_.architectSheet = (self.sheet1Input.GetValue())
        global_.output1Sheet = (self.sheet3Input.GetValue())
        global_.output2Sheet = (self.sheet4Input.GetValue())

    def onButtonRun(self, event):    # pragma: no cover
        logging.FileHandler("LogError.log", mode='w')
        global_.checkOem = self.oemInput.GetValue()
        self.configVariables()
        checkState = True

        if checkState and not self.checkArchitectValuesFromGUI():
            global_.stateOfTheProgram = "Error Architect section \nCheck if the boxes are 0 or empty!"
            checkState = False

        if checkState and not self.checkOemValuesFromGUI():
            global_.stateOfTheProgram = "Error Oem section \nCheck if the boxes are 0 or empty!"
            checkState = False

        if checkState and not self.checkSystemValuesFromGUI():
            global_.stateOfTheProgram = "Error System section \nCheck if the boxes are 0 or empty!"
            checkState = False

        if checkState and not self.checkOutputValuesFromGUI():
            global_.stateOfTheProgram = "Error Output section \nCheck if the boxes are 0 or empty!"
            checkState = False

        if checkState and not self.checkSheetValuesFromGUI():
            global_.stateOfTheProgram = "Error Sheet section \nCheck if the boxes are 0 or empty!"
            checkState = False

        if checkState and not self.checkPathGUI():
            checkState = False

        self.checkOpenedFile()
        if checkState and not global_.outputPathOk:
            global_.stateOfTheProgram = "Output Excel is opened! The program can't write inside an opened Excel!"
            checkState = False

        if checkState and not global_.checkOem:
            listWithSheets = self.get_sheet_ids(global_.pathArchitect)
            for sheetName in listWithSheets:
                if global_.projectSpecificName.strip().upper() in sheetName.strip().upper():
                    dlg = wx.MessageDialog(None, "The checkbox for oem is disabled and the Project OEM Name is filled!\nDo you want to run the project with the selected OEM?", 'OEM check',
                                           wx.YES_NO | wx.ICON_QUESTION)
                    result = dlg.ShowModal()
                    if result == wx.ID_YES:
                        global_.checkOem = True
                        self.oemInput.SetValue(True)
                        self.configVariables()
                        self.oemProjectName.SetStringSelection(sheetName)
                        global_.projectOemName = sheetName
                        self.oemProjectName.Enable(True)
                        self.startRowOem.Enable(True)
                        self.endRowOem.Enable(True)
                        self.nameColumnOem.Enable(True)
                        self.countColumnOem.Enable(True)
                        self.valueColumnOem.Enable(True)
                        self.rowOemA2l.Enable(True)
                        self.columnOemA2l.Enable(True)
                        if not self.checkOemValuesFromGUI():
                            global_.stateOfTheProgram = "Error Oem section \nCheck if the boxes are 0 or empty!"
                            checkState = False

        if checkState:
            logging.warning("VERSION: " + str(Version))
            logging.warning("GUI PARAMETERS:\n")
            try:
                logArchitect()
            except:
                dlg = wx.RichMessageDialog(parent=None, message="Error at Architect config section")
                dlg.ShowModal()
            else:
                try:
                    logSystem()
                except:
                    dlg = wx.RichMessageDialog(parent=None, message="Error at System config section")
                    dlg.ShowModal()
                else:
                    try:
                        logOem()
                    except:
                        dlg = wx.RichMessageDialog(parent=None, message="Error at Oem config section")
                        dlg.ShowModal()
                    else:
                        try:
                            logOutput()
                        except:
                            dlg = wx.RichMessageDialog(parent=None, message="Error at Output config section")
                            dlg.ShowModal()
                        else:
                            exec(open("ArchitectExcel.py").read(), globals())

        if global_.stateOfTheProgram == "The program finished execution!":
            self.OutputDialog()
        else:
            try:
                global_.dialog.Destroy()
            except:
                pass

            dlg = wx.RichMessageDialog(parent=None, message=global_.stateOfTheProgram)
            dlg.ShowModal()

    def onCancel(self, event):  # Event when the Cancel button is pressed.   # pragma: no cover
        self.closeProgram()

    def closeProgram(self):  # pragma: no cover
        self.GetParent().Close()

    @staticmethod
    def checkArchitectValuesFromGUI():   # pragma: no cover
        if global_.nameColumn == 0 or global_.typeColumn == 0 or global_.countColumn == 0 or global_.valueColumn == 0 or \
                global_.resColumn == 0 or global_.row == 0 or global_.rowEnd == "" or global_.readRowAddress == 0 \
                or global_.readColumnAddress == 0 or global_.rowA2L == 0 or global_.columnA2L == 0 or global_.defaultColumn == "":
            return False
        return True

    def checkOemValuesFromGUI(self):     # pragma: no cover
        if self.oemInput.GetValue():
            if global_.projectOemName == "" or global_.nameColumnOem == 0 or \
                    global_.countColumnOem == 0 or global_.valueColumnOem == 0 or global_.rowOem == 0 or \
                    global_.rowEndOem == "" or global_.rowOEMa2l == 0 or global_.columnOemA2L == 0:
                return False
        return True

    @staticmethod
    def checkOpenedFile():   # pragma: no cover
        try:
            os.rename(global_.pathOutput, 'TestingTheNameOutput.xls')
            os.rename('TestingTheNameOutput.xls', global_.pathOutput)
            global_.outputPathOk = True

        except OSError:
            global_.outputPathOk = False

    @staticmethod
    def checkSystemValuesFromGUI():  # pragma: no cover
        if global_.nameColumn2 == 0 or global_.typeColumn2 == 0 or global_.countColumn2 == 0 or global_.valueColumn2 == 0 \
                or global_.row2 == 0 or global_.rowEnd2 == "" or global_.limitValueColumn == 0 or global_.projectNameType == "":
            return False
        return True

    @staticmethod
    def checkOutputValuesFromGUI():  # pragma: no cover
        if global_.nameColumnWrite == 0 or global_.typeColumnWrite == 0 or global_.countColumnWrite == 0 or \
                global_.valueColumnWrite == 0 or global_.rowWrite == 0 or global_.writeL2Architect == 0 or \
                global_.writeTreeLevel2 == 0 or global_.columnRes == 0 or global_.rowRes == 0 or \
                global_.writeRowAddress == 0 or global_.writeColumnAddress == 0 or \
                global_.rowWriteA2L == 0 or global_.columnWriteA2L == 0 or global_.columnWriteSheet2 == 0 \
                or global_.rowWriteSheet2 == 0 or global_.lowLimitColumn == 0 or global_.maxLimitColumn == 0:
            return False
        return True

    @staticmethod
    def checkSheetValuesFromGUI():   # pragma: no cover
        if global_.architectSheet == "" or global_.systemSheet == "" or global_.output1Sheet == "" \
                or global_.output2Sheet == "":
            return False
        return True

    @staticmethod
    def checkPathGUI():  # pragma: no cover
        try:
            if global_.pathArchitect == "" or global_.pathArchitect == "Path":
                raise NameError
        except NameError:
            global_.stateOfTheProgram = "Architect PATH Error\n" + "Select the Architect excel!"
        else:
            try:
                if global_.pathSystem == "" or global_.pathSystem == "Path":
                    raise NameError
            except NameError:
                global_.stateOfTheProgram = "System PATH Error\n" + "Select the System excel"

            else:
                try:
                    if global_.pathOutput == "" or global_.pathOutput == "Path":
                        raise NameError
                except NameError:
                    global_.stateOfTheProgram = "Output PATH Error\n" + "Select the Output excel"
                else:
                    return True
        return False


if __name__ == "__main__":  # Activates the application.     # pragma: no cover
    app = MyApp()
    # --------- FOR DEBUGGING GUI ---------

    # import wx.lib.inspection
    # wx.lib.inspection.InspectionTool().Show()

    # --------- FOR DEBUGGING GUI ---------
    app.MainLoop()

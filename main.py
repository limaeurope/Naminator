import shutil
from functools import reduce

import PySimpleGUI as sg
import os, sys
import ctypes
import platform

import conventioner
import constandard

conStandard = None

#---------------- pyinstaller compatibility#----------------
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
#----------------/pyinstaller compatibility#----------------

class DestFile:
    def __init__(self, p_source, p_values, p_conStandard: constandard.Convention, ):
        self.sSourceFileName = p_source
        if p_conStandard:
            self.isNotRenamed = False
            self.sDest = p_conStandard.getFileName(p_source, p_values)
        else:
            self.isNotRenamed = True
            self.sDest = p_source

    def __str__(self):
        return self.sDest

    def __eq__(self, other):
        return self.sSourceFileName == other.sSourceFileName


def make_dpi_aware():
    if int(platform.release()) >= 8:
        ctypes.windll.shcore.SetProcessDpiAwareness(True)

structKeys = { "disabled_readonly_background_color":   "light grey",
                "enable_events":                        True,
           }

bottomKeys = {
                "button_background_color":               "light grey",
                "enable_events":                        True,
           }

DISABLED_FOR_NOW = {"disabled": True}

def ComboBoxT(CBox_name, readonly=True):
    return sg.InputCombo(["Nincs Nevezék"], default_value="Nincs Nevezék",key=CBox_name, readonly=readonly, **bottomKeys, **DISABLED_FOR_NOW)

menu_def = [['File', ['Exit',]],]

make_dpi_aware()

DocNum = 0
iSheet = 0
bCreateNewFolder = True

CBox_Project    = "Projekt kód"
CBox_Phase      = "Tervfázis"
CBox_Building   = "Épület, -rész jele"
CBox_Storey     = "Szint"
CBox_Role       = "Szerepkör"
CBox_DocType    = "Dokumentumtípus"
CBox_Status     = "Státusz"
CBox_Rev        = "Revízió"
CBox_name9 = "post1"

CBox_List = [
    ComboBoxT(CBox_Project, readonly=False),
    ComboBoxT(CBox_Phase),
    ComboBoxT(CBox_Building),
    ComboBoxT(CBox_Storey),
    ComboBoxT(CBox_Role),
    ComboBoxT(CBox_DocType),
    ComboBoxT(CBox_Status),
    ComboBoxT(CBox_Rev),
    ComboBoxT(CBox_name9),
]

FileListBox = sg.Listbox(values=[], size=(70, 30), key="FILELIST", enable_events=True, select_mode=sg.LISTBOX_SELECT_MODE_EXTENDED)
ModListBox = sg.Listbox(values=[], size=(70, 30), key="MODLIST", enable_events=True, select_mode=sg.LISTBOX_SELECT_MODE_EXTENDED)
ConventionTemplate = ""

SelectedItems = []
SelectedItem = None
DocTypeDict = {}
errorList = []

WestCoast = [
    [
        sg.Checkbox('Statikus', default=True, key="STRUCTURAL", enable_events=True, **DISABLED_FOR_NOW),

        sg.Text("Nevezék Template: ", font=('Helvetica', 8, 'normal')),
        sg.In(size=(25, 1), key="CONVENTION_TEMPLATE_FILE", enable_events=True),
        sg.FileBrowse(),

        sg.Text("Munkamappa: "),
        sg.In(size=(25, 1), key="FOLDER", enable_events=True),
        sg.FolderBrowse(),

        sg.Text("Filenév előnézet "),
        sg.In(size=(50, 1), key="PREV", enable_events=True),

        sg.Checkbox('Új könytár', default=True, key="NEW_FOLDER_CREATE", enable_events=True, **DISABLED_FOR_NOW),
        sg.In(size=(25, 1), key="NEW_FOLDER_NAME", default_text='Result', **structKeys),
    ],

    [
        sg.Text("Munkalap neve: ", font=('Helvetica', 8, 'normal')),
        sg.InputCombo([], key="SHEET_NAME", enable_events=True, size=(25, 1), ),
        sg.InputCombo([], key="START_COL", size=(25, 1),
                      tooltip='A "Tervfázis" oszlop fejlécben használt neve, ehhez képest indul a számozás, pl. "Tervfázis", "Phase"',
                      enable_events=True,
                      ),
        sg.InputCombo([], key="NAME_COL", size=(25, 1),
                      tooltip='Az oszlop fejlécben használt neve, amelyben a tervlap elnevezése van, pl. "Terv elnevezése"',
                      enable_events=True,
                      ),
    ],

    [
        FileListBox,
        sg.Column([[sg.Button(">", key="ADD"), ], [sg.Button(">>", key="ADDALL"),]], justification="top"),
        ModListBox,
        sg.Column([[sg.Button("Lista törlése", key="FLUSHLIST")], [sg.Button("Kijelöltek törlése", key="FLUSHSELECTED")]]),
        sg.Image(resource_path(r'lima.2016_Feher.png'))
    ],

    [
        sg.Column([[sg.Text(CBox_Project,   size=(15, 1), key="Drop1H")], [CBox_List[0]]]), sg.Text("", size=(1, 0)),
        sg.Column([[sg.Text(CBox_Phase,     size=(15, 1), key="Drop2H")], [CBox_List[1]]]), sg.Text("", size=(1, 0)),
        sg.Column([[sg.Text(CBox_Building,  size=(15, 1), key="Drop3H")], [CBox_List[2]]]), sg.Text("", size=(1, 0)),
        sg.Column([[sg.Text(CBox_Storey,    size=(15, 1), key="Drop4H")], [CBox_List[3]]]), sg.Text("", size=(1, 0)),
        sg.Column([[sg.Text(CBox_Role,      size=(15, 1), key="Drop6H")], [CBox_List[5]]]), sg.Text("", size=(1, 0)),
        sg.Column([[sg.Text(CBox_DocType,   size=(15, 1), key="Drop5H")], [CBox_List[4]]]), sg.Text("", size=(1, 0)),

        sg.Column([[sg.Text("Dok. száma",   size=(15, 1), key="Drop7H")], [sg.Spin([i for i in range(1,1000)], size=(10,1), initial_value=1,key="CustNUM"),]]),sg.Text("", size=(1,0)),
        sg.Column([[sg.Text("Megnevezés",   size=(25, 1), key="Drop9H")], [sg.In(size=(25, 1), key="CustName")]]), sg.Text("", size=(1, 0)),

        sg.Column([[sg.Text(CBox_Status,   size=(15, 1), key="Drop8H")], [CBox_List[7]]]), sg.Text("", size=(1, 0)),
        sg.Column([[sg.Text(CBox_Rev,      size=(25, 1), key="Drop10H")], [CBox_List[8]]]),
        sg.Column([[sg.Button("Átnevez",   size=(10,1), key="RENAME")],  [sg.Checkbox('Elnevezés megtartása', default=True, key="KeepName")]])
    ]
]

layout = [[sg.Menu(menu_def)],
    [   
        sg.Column(WestCoast, justification="right")
    ]
]

sg.theme("Dark Blue 3")
window = sg.Window(title="RENAM-R", layout=layout, margins=(0, 8), font=('Helvetica', 8, 'normal'))

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------

while True:
    event, values = window.read()

    if event == "Exit" or event == sg.WIN_CLOSED:
        break

    if event == "FOLDER":
        folder = values["FOLDER"]
        if folder:
            file_list = os.listdir(folder)

            filenames = [
                fext
                for fext in file_list
                if os.path.isfile(os.path.join(folder, fext))
                ]
            window["FILELIST"].update(filenames)

    if event == "MODLIST":
        SelectedItem = window["MODLIST"].get()[0]
        window["PREV"].update(SelectedItem.sDest)

    if event == "CONVENTION_TEMPLATE_FILE":
        if os.path.isfile(values["CONVENTION_TEMPLATE_FILE"]):
            try:
                if conStandard:
                    conStandard.reset()
                    window["SHEET_NAME"].update(values=[])
                    window["START_COL"].update(values=[])
                    window["NAME_COL"].update(values=[])

                if not window["STRUCTURAL"].get():
                    ConventionTemplate = conventioner.GetConvention(values["CONVENTION_TEMPLATE_FILE"])
                    ComboNameList = ConventionTemplate[0]
                    ComboList = ConventionTemplate[1]

                    window[CBox_Project].update(values=ComboList[0], size=(10, 5), visible=True, set_to_index=0)
                    window[CBox_Phase].update(values=ComboList[1], size=(16, 5), visible=True, set_to_index=0)
                    window[CBox_Building].update(values=ComboList[2], size=(16, 5), visible=True, set_to_index=0)
                    window[CBox_Storey].update(values=ComboList[3], size=(16, 5), visible=True, set_to_index=0)
                    window[CBox_Role].update(values=ComboList[4], size=(16, 5), visible=True, set_to_index=0)
                    window[CBox_DocType].update(values=ComboList[5], size=(16, 5), visible=True, set_to_index=0)
                    # window[CBox_name7].update(values=ComboList[6],size=(15, 5), visible=True)
                    window[CBox_Rev].update(values=ComboList[7], size=(16, 5), visible=True, set_to_index=0)
                    window[CBox_name9].update(values=ComboList[8], size=(16, 5), visible=True, set_to_index=0)

                    conStandard = constandard.GeneralConvention(values["CONVENTION_TEMPLATE_FILE"], iSheet)
                else:
                    conStandard = constandard.StructuralDesignerConvention(values["CONVENTION_TEMPLATE_FILE"],
                                                                           values["SHEET_NAME"],
                                                                           values["START_COL"],
                                                                           values["NAME_COL"])
                window["SHEET_NAME"].update(values=conStandard.sheets, set_to_index=0)
                #FIXME update() method

                # conStandard = constandard.StructuralDesignerConvention(values["CONVENTION_TEMPLATE_FILE"],
                #                                                        values["SHEET_NAME"],
                #                                                        values["START_COL"],
                #                                                        values["NAME_COL"])
                # window["START_COL"].update(values=constandard.StructuralDesignerConvention.headersList)
                # window["NAME_COL"].update(values=constandard.StructuralDesignerConvention.headersList)
                # _phaseKey = constandard.StructuralDesignerConvention.getIndexByKey("Phase")
                # window["START_COL"].update(values=constandard.StructuralDesignerConvention.headersList, set_to_index=_phaseKey)
                # _nameKey = constandard.StructuralDesignerConvention.getIndexByKey("Document name")
                # window["NAME_COL"].update(values=constandard.StructuralDesignerConvention.headersList, set_to_index=_nameKey)

            except (TypeError, KeyboardInterrupt, PermissionError, IndexError) as e:
                continue

    if event == "RENAME":
        folder = values["FOLDER"]
        file_list = os.listdir(folder)

        if bCreateNewFolder:
            _sNewFolder = os.path.join(folder, window["NEW_FOLDER_NAME"].get())
            if not os.path.exists(_sNewFolder):
                os.mkdir(_sNewFolder)

            for item in SelectedItems:
                shutil.copyfile(os.path.join(folder, item.sSourceFileName), os.path.join(folder, _sNewFolder, item.sDest) )
        else:
            for item in SelectedItems:
                os.rename(item.sSourceFileName, item.sDest)

        filenames = [
            fext
            for fext in file_list
            if os.path.isfile(os.path.join(folder, fext))
            ]

        window["FILELIST"].update(filenames)
        SelectedItems = []
        window["MODLIST"].update(SelectedItems)

    if event == "FLUSHLIST":
        SelectedItems = []
        window["MODLIST"].update(SelectedItems)

    if event == "pre6":
        window[CBox_Role].update(values=(conStandard.getAvailableDocID(values[CBox_DocType])[0]), size=(15, 5), visible=True)
        DocTypeDict = conStandard.getAvailableDocID(values[CBox_DocType])[1]

    if event == "PREV":
        SelectedItem.sDest = window["PREV"].get()
        window["MODLIST"].update(SelectedItems)

    if event == "ADDALL":
        try:
            for fileName in window["FILELIST"].get_list_values():
                if (fileToRename :=  DestFile(fileName, values, conStandard)) not in SelectedItems:
                    SelectedItems.append(fileToRename)
                    SelectedItem = fileToRename

                    window["MODLIST"].update(SelectedItems)
                    window["PREV"].update(fileToRename.sDest)
        except IndexError:
            pass

    if event == "ADD":
        try:
            for fileName in window["FILELIST"].get():
                if (fileToRename :=  DestFile(fileName, values, conStandard)) not in SelectedItems:
                    if fileToRename.isNotRenamed:
                        errorList.append(fileToRename)
                    SelectedItems.append(fileToRename)
                    SelectedItem = fileToRename

                    window["MODLIST"].update(SelectedItems)
                    window["PREV"].update(fileToRename.sDest)
            if len(errorList) == 1:
                sg.popup(f"Ez a fájl nem lett átnevezve: {errorList[0]}")
            elif len(errorList) >1:
                _sFiles = reduce(lambda f, f1: f"{f}\n{f1}", errorList)
                sg.popup(f"Ezez a fájlok nem lettek átnevezve: {_sFiles}")
            errorList = []

        except IndexError:
            pass

    if event == "FLUSHSELECTED":
        for fileName in window["MODLIST"].get():
            SelectedItems.remove(fileName)
        window["MODLIST"].update(SelectedItems)

    if event == "NEW_FOLDER_CREATE":
        bCreateNewFolder = window["NEW_FOLDER_CREATE"].get()
        window["NEW_FOLDER_NAME"].update(disabled=not bCreateNewFolder)

    if event == "SHEET_NAME":
        window["START_COL"].update(values=[])
        window["NAME_COL"].update(values=[])
        conStandard = constandard.StructuralDesignerConvention(values["CONVENTION_TEMPLATE_FILE"],
                                                               values["SHEET_NAME"],
                                                               values["START_COL"],
                                                               values["NAME_COL"])
        window["START_COL"].update(values=constandard.StructuralDesignerConvention.headersList)
        window["NAME_COL"].update(values=constandard.StructuralDesignerConvention.headersList)

    if event == "START_COL":
        conStandard = constandard.StructuralDesignerConvention(values["CONVENTION_TEMPLATE_FILE"],
                                                               values["SHEET_NAME"],
                                                               values["START_COL"],
                                                               values["NAME_COL"])

    if event == "NAME_COL":
        conStandard = constandard.StructuralDesignerConvention(values["CONVENTION_TEMPLATE_FILE"],
                                                               values["SHEET_NAME"],
                                                               values["START_COL"],
                                                               values["NAME_COL"])

    if event == "STRUCTURAL":
        window["START_COL"].update(disabled=values["STRUCTURAL"])
        window["NAME_COL"].update(disabled=values["STRUCTURAL"])

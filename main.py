import PySimpleGUI as sg
import os

import conventioner
import ctypes
import platform
import constandard

conStandard = None

class DestFile:
    def __init__(self, p_values, p_conStandard: constandard.Convention, ):
        source = p_values["-FILELIST-"][0]
        self.sSourceFileName = source
        if p_conStandard:
            self.sDest = p_conStandard.getFileName(p_values)
        else:
            self.sDest = source

    def __str__(self):
        return self.sDest

    def __eq__(self, other):
        return self.sSourceFileName == other.sSourceFileName


def make_dpi_aware():
    if int(platform.release()) >= 8:
        ctypes.windll.shcore.SetProcessDpiAwareness(True)

def ComboBoxT(CBox_name, readonly=True):
    return sg.InputCombo(["Nincs Nevezék"], default_value="Nincs Nevezék",key=CBox_name, enable_events=True, readonly=readonly)

menu_def = [['File', ['Open', 'Save', 'Exit',]],
                ['Edit', ['Paste', ['Special', 'Normal',], 'Undo'],],
                ['Help', 'About...'],]

make_dpi_aware()

DocNum = 0

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

FileListBox = sg.Listbox(values=[], size=(70, 30), key="-FILELIST-", enable_events=True)
ModListBox = sg.Listbox(values=[], size=(70, 30), key="-MODLIST-", enable_events=True, select_mode=sg.LISTBOX_SELECT_MODE_SINGLE)
ConventionTemplate = ""

SelectedItems = []
SelectedItem = None
DocTypeDict = {}

WestCoast = [
    [
        #FIXME default=True
        sg.Checkbox('Statikus', default=True, key="-STRUCTURAL-"),

        sg.Text("Nevezék Template: ", font=('Helvetica', 8, 'normal')),
        sg.In(size=(25, 1), key="-CONVENTION_TEMPLATE_FILE-", enable_events=True),
        sg.FileBrowse(),

        sg.Text("Munkamappa: "),
        sg.In(size=(25, 1), key="-FOLDER-", enable_events=True),
        sg.FolderBrowse(),

        sg.Text("Filenév: "),
        sg.In(size=(25, 1), key="-PREV-", enable_events=True),
    ],

    [
        FileListBox,ModListBox,
        sg.Button("Lista törlése", key="-FLUSHLIST-"),
        sg.Image(r'lima.2016_Feher.png')
    ],

    [
        sg.Column([[sg.Text(CBox_Project,   size=(15, 1), key="-Drop1H-")], [CBox_List[0]]]), sg.Text("-", size=(1, 0)),
        sg.Column([[sg.Text(CBox_Phase,     size=(15, 1), key="-Drop2H-")], [CBox_List[1]]]), sg.Text("-", size=(1, 0)),
        sg.Column([[sg.Text(CBox_Building,  size=(15, 1), key="-Drop3H-")], [CBox_List[2]]]), sg.Text("-", size=(1, 0)),
        sg.Column([[sg.Text(CBox_Storey,    size=(15, 1), key="-Drop4H-")], [CBox_List[3]]]), sg.Text("-", size=(1, 0)),
        sg.Column([[sg.Text(CBox_Role,      size=(15, 1), key="-Drop6H-")], [CBox_List[5]]]), sg.Text("-", size=(1, 0)),
        sg.Column([[sg.Text(CBox_DocType,   size=(15, 1), key="-Drop5H-")], [CBox_List[4]]]), sg.Text("-", size=(1, 0)),

        sg.Column([[sg.Text("Dok. száma",   size=(15, 1), key="-Drop7H-")], [sg.Spin([i for i in range(1,1000)], size=(10,1), initial_value=1,key="-CustNUM-"),]]),sg.Text("-", size=(1,0)),
        sg.Column([[sg.Text("Megnevezés",   size=(25, 1), key="-Drop9H-")], [sg.In(size=(25, 1), key="-CustName-")]]), sg.Text("-", size=(1, 0)),

        sg.Column([[sg.Text(CBox_Status,   size=(15, 1), key="-Drop8H-")], [CBox_List[7]]]), sg.Text("-", size=(1, 0)),
        sg.Column([[sg.Text(CBox_Rev,      size=(25, 1), key="-Drop10H-")], [CBox_List[8]]]),
        sg.Column([[sg.Button("Átnevez",   size=(10,1), key="-RENAME-")],  [sg.Checkbox('Elnevezés megtartása', default=True, key="-KeepName-")]])
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

    if event == "-FOLDER-":
        folder = values["-FOLDER-"]
        if folder:
            file_list = os.listdir(folder)

            filenames = [
                fext
                for fext in file_list
                if os.path.isfile(os.path.join(folder, fext))
                ]

            window["-FILELIST-"].update(filenames)

    if event == "-FILELIST-":
        try:
            if (fileToRename :=  DestFile(values, conStandard)) not in SelectedItems:
                SelectedItems.append(fileToRename)
                SelectedItem = fileToRename

                window["-MODLIST-"].update(SelectedItems)
                window["-PREV-"].update(fileToRename.sDest)

        except IndexError:
            pass

    if event == "-MODLIST-":
        SelectedItem = window["-MODLIST-"].get()[0]
        window["-PREV-"].update(SelectedItem.sDest)

    if event == "-CONVENTION_TEMPLATE_FILE-":
        if os.path.isfile(values["-CONVENTION_TEMPLATE_FILE-"]):
            try:
                if not window["-STRUCTURAL-"].get():
                    ConventionTemplate = conventioner.GetConvention(values["-CONVENTION_TEMPLATE_FILE-"])
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

                    conStandard = constandard.GeneralConvention(values["-CONVENTION_TEMPLATE_FILE-"])
                else:
                    conStandard = constandard.StructuralDesignerConvention(values["-CONVENTION_TEMPLATE_FILE-"])
            except (TypeError, KeyboardInterrupt, PermissionError) as e:
                continue

    if event == "-RENAME-":
        for item in SelectedItems:
            os.rename(item.sSourceFileName, item.sDest)

        folder = values["-FOLDER-"]
        file_list = os.listdir(folder)

        filenames = [
            fext
            for fext in file_list
            if os.path.isfile(os.path.join(folder, fext))
            ]

        window["-FILELIST-"].update(filenames)
        SelectedItems = []
        window["-MODLIST-"].update(SelectedItems)

    if event == "-FLUSHLIST-":
        SelectedItems = []
        window["-MODLIST-"].update(SelectedItems)

    if event == "pre6":
        window[CBox_Role].update(values=(conStandard.getAvailableDocID(values[CBox_DocType])[0]), size=(15, 5), visible=True)
        DocTypeDict = conStandard.getAvailableDocID(values[CBox_DocType])[1]

    if event == "-PREV-":
        SelectedItem.sDest = window["-PREV-"].get()
        window["-MODLIST-"].update(SelectedItems)


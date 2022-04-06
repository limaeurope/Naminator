import PySimpleGUI as sg
import os
import conventioner
import ctypes
import platform
import constandard

def make_dpi_aware():
    if int(platform.release()) >= 8:
        ctypes.windll.shcore.SetProcessDpiAwareness(True)

def ComboBoxT(CBox_name, CBox_list):
    return sg.InputCombo(CBox_list, default_value=CBox_list[0],key=CBox_name, enable_events=True)

menu_def = [['File', ['Open', 'Save', 'Exit',]],
                ['Edit', ['Paste', ['Special', 'Normal',], 'Undo'],],
                ['Help', 'About...'],]

#menu_def = []

make_dpi_aware()

DocNum = 0


CBox_List = []

CBox_name1 = "pre1"
CBox_name2 = "pre2"
CBox_name3 = "pre3"
CBox_name4 = "pre4"
CBox_name5 = "pre5"
CBox_name6 = "pre6"
CBox_name7 = "pre7"
CBox_name8 = "pre8"
CBox_name9 = "post1"


Choice1 = ComboBoxT(CBox_name1,["Nincs Nevezék"])
Choice2 = ComboBoxT(CBox_name2,["Nincs Nevezék"])
Choice3 = ComboBoxT(CBox_name3,["Nincs Nevezék"])
Choice4 = ComboBoxT(CBox_name4,["Nincs Nevezék"])
Choice5 = ComboBoxT(CBox_name5,["Nincs Nevezék"])
Choice6 = ComboBoxT(CBox_name6,["Nincs Nevezék"])
Choice7 = ComboBoxT(CBox_name7,["Nincs Nevezék"])
Choice8 = ComboBoxT(CBox_name8,["Nincs Nevezék"])
Choice9 = ComboBoxT(CBox_name9,["Nincs Nevezék"])

CBox_List = [Choice1,
Choice2,
Choice3,
Choice4,
Choice5,
Choice6,
Choice7,
Choice8,
Choice9
]

FileListBox = sg.Listbox(values=[], size=(70, 30), key="-FILELIST-", enable_events=True)
ModListBox = sg.Listbox(values=[], size=(70, 30), key="-MODLIST-", enable_events=True)
ConventionTemplate = ""


SelectedItems = []


DocTypeDict = {}






WestCoast = [
    [
        sg.Text("Nevezék Template: ", font=('Helvetica', 8, 'normal')),
        sg.In(size=(25, 1), key="-CTFILE-", enable_events=True),
        sg.FileBrowse(),
        sg.Text("Munkamappa: "),
        sg.In(size=(25, 1), key="-FOLDER-", enable_events=True),
        sg.FolderBrowse()
    ],
    [
        FileListBox,ModListBox,
        sg.Button("Lista törlése", key="-FLUSHLIST-"),
        # sg.Image(r'C:\Users\matepeter\Documents\Homestead\PythonLearn\1-0_BlackBoxDev\1-0_DEV\1-1_Projects\Naminator\lima.2016_Feher.png')
    ],

    [
        sg.Column([[sg.Text("Projekt kód",size=(15, 1), key="-Drop1H-")],[CBox_List[0]]]),
        sg.Text("-", size=(1,0)),
        sg.Column([[sg.Text("Tervfázis",size=(15, 1), key="-Drop2H-")],[CBox_List[1]]]),
        sg.Text("-", size=(1,0)),
        sg.Column([[sg.Text("Épület, -rész jele",size=(15, 1), key="-Drop3H-")],[CBox_List[2]]]),
        sg.Text("-", size=(1,0)),
        sg.Column([[sg.Text("Szint",size=(15, 1), key="-Drop4H-")],[CBox_List[3]]]),
        sg.Text("-", size=(1,0)),
        sg.Column([[sg.Text("Szerepkör",size=(15, 1), key="-Drop6H-")],[CBox_List[5]]]),
        sg.Text("-", size=(1,0)),
        sg.Column([[sg.Text("Dokumentumtípus",size=(15, 1), key="-Drop5H-")],[CBox_List[4]]]),
        sg.Text("-", size=(1,0)),
        sg.Column([[sg.Text("Dok. száma",size=(15, 1), key="-Drop7H-")],[sg.Spin([i for i in range(1,1000)], size=(10,1), initial_value=1,key="-CustNUM-"),]]),
        sg.Text("-", size=(1,0)),
        sg.Column([[sg.Text("Megnevezés",size=(25, 1), key="-Drop9H-")],[sg.In(size=(25, 1), key="-CustName-")]]),
        sg.Text("-", size=(1,0)),
        sg.Column([[sg.Text("Státusz",size=(15, 1), key="-Drop8H-")],[CBox_List[7]]]),
        sg.Text("-", size=(1,0)),
        sg.Column([[sg.Text("Revízió",size=(25, 1), key="-Drop10H-")],[CBox_List[8]]]),
        sg.Column([[sg.Button("Átnevez", size=(10,1), key="-RENAME-")],[sg.Checkbox('Elnevezés megtartása', default=True, key="-KeepName-")]])
    ]

]


layout = [[sg.Menu(menu_def)],
    [   

        sg.Column(WestCoast, justification="right")


    ]
]




sg.theme("Dark Blue 3")
window = sg.Window(title="RENAM-R", layout=layout, margins=(0, 8), font=('Helvetica', 8, 'normal'))



while True:
    event, values = window.read()



    if event == "Exit" or event == sg.WIN_CLOSED:
        
        break



    if event == "-FOLDER-":
        folder = values["-FOLDER-"]
        file_list = os.listdir(folder)
        

        
        filenames = [
            fext
            for fext in file_list
            if os.path.isfile(os.path.join(folder, fext))
            ]

        
        


        window["-FILELIST-"].update(filenames)



    if event == "-FILELIST-":
        if SelectedItems.count(values["-FILELIST-"]) == 0:
            
            SelFileTemp = (values["-FILELIST-"])
            SelectedItems.append(SelFileTemp)

            window["-MODLIST-"].update(SelectedItems)
    
    if event == "-MODLIST-":
        #sg.Popup('Selected ', values["-MODLIST-"])
        pass
        


    if event == "-CTFILE-":
        #sg.Popup("Nevezék Kiválasztva: ", values["-CTFILE-"])

        if os.path.isfile(values["-CTFILE-"]):
            templatePath = values["-CTFILE-"]
            ConventionTemplate = conventioner.GetConvention(values["-CTFILE-"])
            ComboNameList = ConventionTemplate[0]
            ComboList = ConventionTemplate[1]
            # print(ComboNameList)
            # print(type(ComboList))

            window[CBox_name1].update(values=ComboList[0],size=(10, 5), visible=True, set_to_index=0)
            window[CBox_name2].update(values=ComboList[1],size=(16, 5), visible=True, set_to_index=0)
            window[CBox_name3].update(values=ComboList[2],size=(16, 5), visible=True, set_to_index=0)
            window[CBox_name4].update(values=ComboList[3],size=(16, 5), visible=True, set_to_index=0)
            window[CBox_name5].update(values=ComboList[4],size=(16, 5), visible=True, set_to_index=0)
            window[CBox_name6].update(values=ComboList[5],size=(16, 5), visible=True, set_to_index=0)
            # window[CBox_name7].update(values=ComboList[6],size=(15, 5), visible=True)
            window[CBox_name8].update(values=ComboList[7],size=(16, 5), visible=True, set_to_index=0)
            window[CBox_name9].update(values=ComboList[8],size=(16, 5), visible=True, set_to_index=0)




    if event == "-RENAME-":
        
        IndexStart = int(values["-CustNUM-"])






        for item in SelectedItems:

            
            IndexStartSTR = str(IndexStart).zfill(3)
            # print(IndexStartSTR)
            constandard.ConTemplatePath = values["-CTFILE-"]
            prefixes = values[CBox_name1]+"-"+constandard.getPhase(values[CBox_name2], templatePath)+"-"+constandard.getBuilding(values[CBox_name3], templatePath)+"-"+constandard.getStoreyDict(values[CBox_name4], templatePath)+"-"+DocTypeDict.get(values[CBox_name5])+"-"+constandard.getRolesDict(values[CBox_name6], templatePath) + "-" + IndexStartSTR + "-"
            suffixes = "-" + constandard.getStatus(values[CBox_name8], templatePath) + "-" + values[CBox_name9]
            workfile = folder + "/" + item[0]
            if values["-KeepName-"] == False:
                pass                
                moddedfile = folder + "/" + prefixes + values["-CustName-"] + suffixes + os.path.splitext(item[0])[1]
            if values["-KeepName-"] == True:
                pass
                moddedfile = folder + "/" + prefixes + os.path.splitext(item[0])[0] + suffixes + os.path.splitext(item[0])[1]

            # print(moddedfile)
            os.rename(workfile, moddedfile)
            IndexStart = IndexStart+1
        
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
        # print(values[CBox_name6])
        
        window[CBox_name5].update(values=(constandard.getAvailableDocID(values[CBox_name6], templatePath)[0]),size=(15, 5), visible=True)
        DocTypeDict = constandard.getAvailableDocID(values[CBox_name6], templatePath)[1]
        # print(DocTypeDict)
        # print(constandard.getAvailableDocID(values[CBox_name6], templatePath))
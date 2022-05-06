import xlrd
import os

_A_ =  0;   _B_ =  1;   _C_ =  2;   _D_ =  3;   _E_ =  4
_F_ =  5;   _G_ =  6;   _H_ =  7;   _I_ =  8;   _J_ =  9
_K_ = 10;   _L_ = 11;   _M_ = 12;   _N_ = 13;   _O_ = 14
_P_ = 15;   _Q_ = 16;   _R_ = 17;   _S_ = 18;   _T_ = 19
_U_ = 20;   _V_ = 21;   _W_ = 22;   _X_ = 23;   _Y_ = 24
_Z_ = 25

ConTemplatePath = r"Q:\BIM_management\10_BIMgoodies\DATA\LIMA_nevezektan.xls"

#FIXME
CBox_Project    = "Projekt kód"
CBox_Phase      = "Tervfázis"
CBox_Building   = "Épület, -rész jele"
CBox_Storey     = "Szint"
CBox_Role       = "Szerepkör"
CBox_DocType    = "Dokumentumtípus"
CBox_Status     = "Státusz"
CBox_Rev        = "Revízió"
CBox_name9 = "post1"

class Convention:
    singleton = None
    template = None
    sheetAll = None
    sheetRoleID = None
    sheetRole = None
    _sSheet = 0
    sheets = None

    def __new__(cls, templateFile, p_sSheet, *args, **kwargs):
        if not cls.singleton:
            if not templateFile:
                return None
            cls.singleton = super().__new__(cls)
            try:
                cls.template = xlrd.open_workbook(templateFile, formatting_info=True)
                cls._sSheet = p_sSheet
                cls.sheets = [s.name for s in cls.template.sheets()]
            except (FileNotFoundError, PermissionError, IndexError):
                return None
        return cls.singleton

    def getConvention(self, selectedValue, keyColIndex, valueColIndex):
        try:
            return next((z[1] for z in zip(Convention.sheetAll.col_values(keyColIndex)[1:],
                                           Convention.sheetAll.col_values(valueColIndex)[1:]) if z[0] == selectedValue))
        except Exception:
            return None

    def getFileName(self, p_source, p_values):
        """To be overridden"""
        pass

    @classmethod
    def reset(cls):
        cls.singleton = None
        cls.template = None
        cls.sheetAll = None
        cls.sheetRoleID = None
        cls.sheetRole = None
        cls._sSheet = 0
        cls.sheets = None



class GeneralConvention(Convention):
    def __new__(cls, templateFile, p_sSheet, *args, **kwargs):
        Convention.singleton = super().__new__(cls, templateFile, p_sSheet)
        Convention.sheetRole = Convention.template.sheet_by_index(6)
        Convention.sheetRoleID = Convention.template.sheet_by_index(7)
        Convention.sheetAll = Convention.template.sheet_by_index(11)

        return Convention.singleton

    # in 1-2 Tervfázis mapping
    def getPhase(self, chosenPhase):
        return self.getConvention(chosenPhase, _B_, _C_)

    # in 3-4 Épületrész mapping
    def getBuilding(self, chosenBuilding):
        return self.getConvention(chosenBuilding, _D_, _E_)

    # in 5-6 Szint mapping
    def getStoreyDict(self, chosenStorey):
        return self.getConvention(chosenStorey, _F_, _G_)

    # in 9-10 Szerepkör mapping
    def getRolesDict(self, selectedRole):
        return self.getConvention(selectedRole, _J_, _K_)

    # in 12-13 Státusz mapping
    def getStatus(self, chosenStatus):
        return self.getConvention(chosenStatus, _M_, _N_)

    # in 7-8 Dokumentumtípus mapping (CSERÉLVE)
    def getDocTypeNUM(self, DocTypeSTR):
        return self.getConvention(DocTypeSTR, _I_, _H_)

    # Szerepkör - elérhető DoctypeID mapping
    @staticmethod
    def getAvailableDocID(RoleCode):
        D_RRoleID = {z[0]: z[1] for z in
                     zip(Convention.sheetRoleID.col_values(1)[2:], Convention.sheetRoleID.col_values(0)[2:])}

        D_RoleID = {z[0]: z[1] for z in
                    zip(Convention.sheetRole.col_values(1)[2:], Convention.sheetRole.col_values(2)[2:])}

        RoleID = D_RoleID.get(D_RRoleID.get(RoleCode))

        DocTypeList = filter(lambda x: x != "", Convention.sheetAll.col_values(7)[1:])

        D_DocType = {z[0]: z[1] for z in
                     zip(Convention.sheetAll.col_values(7)[1:], Convention.sheetAll.col_values(8)[1:])}

        OutListNUM = [checkItem for checkItem in DocTypeList if checkItem[:2] == RoleID]

        D_FlipDocType = {v: k for k, v in D_DocType.items()}

        OutListSTR = [D_DocType.get(DocIDNum) for DocIDNum in OutListNUM]

        return OutListSTR, D_FlipDocType

    def getFileName(self, p_source, p_values):
        IndexStart = int(p_values["-CustNUM-"])

        _sIndexStart = str(IndexStart).zfill(3)
        prefixes = p_values[CBox_Project]

        prefixes += "-" + self.getPhase(p_values[CBox_Phase])
        prefixes += "-" + self.getBuilding(p_values[CBox_Building])
        prefixes += "-" + self.getStoreyDict(p_values[CBox_Storey])
        prefixes += "-" + self.getDocTypeNUM(p_values[CBox_Role])
        prefixes += "-" + self.getRolesDict(p_values[CBox_DocType])

        prefixes += "-" + _sIndexStart + "-"

        suffixes = "-" + self.getStatus(p_values[CBox_Rev])
        suffixes += "-" + p_values[CBox_name9]

        _ext = os.path.splitext(p_source)[1]
        _origName = os.path.splitext(p_source)[0]
        if p_values["-KeepName-"] == False:
            return prefixes + p_values["-CustName-"] + suffixes + _ext
        if p_values["-KeepName-"] == True:
            return prefixes + _origName + suffixes + _ext


class NameIntPair:
    def __init__(self, p_value, p_number):
        self.value = p_value
        self.number = p_number

    def __str__(self):
        return self.value


def cellVal(p_row, p_sStartCol:NameIntPair, p_inc:int)->str:
    if isinstance(p_row[p_sStartCol.number + p_inc].value, str):
        return p_row[p_sStartCol.number + p_inc].value
    else:
        return "{:.0f}".format(p_row[p_sStartCol.number + p_inc].value)


class StructuralDesignerConvention(Convention):
    _dict = {}
    iCodeStartCol = _C_
    iNameCol = _Q_
    iHeaderRow = 1              #2nd row of sheet
    headersList = []

    def __new__(cls, templateFile, p_sSheet, *args, **kwargs):
        if not cls.singleton:
            if not templateFile:
                return None
            cls.singleton = super().__new__(cls, templateFile, p_sSheet)
        return cls.singleton

    def __init__(self, templateFile, p_sSheet, p_sCodeStartCol, p_sNameCol):
        try:
            if p_sSheet:
                StructuralDesignerConvention.sheetAll = StructuralDesignerConvention.template.sheet_by_name(p_sSheet)
            else:
                StructuralDesignerConvention.sheetAll = StructuralDesignerConvention.template.sheet_by_index(0)
            StructuralDesignerConvention.iCodeStartCol = p_sCodeStartCol
            StructuralDesignerConvention.iNameCol = p_sNameCol
            _headers = zip(StructuralDesignerConvention.sheetAll.row(StructuralDesignerConvention.iHeaderRow), range(StructuralDesignerConvention.sheetAll.ncols))
            StructuralDesignerConvention.headersList = [NameIntPair(h[0].value, h[1]) for h in _headers if h[0].ctype == xlrd.XL_CELL_TEXT]

            for row in StructuralDesignerConvention.sheetAll.get_rows():
                try:
                    self.pre = ""
                    self.pre += cellVal(row, p_sCodeStartCol, -2) + cellVal(row, p_sCodeStartCol, -1)
                    self.pre += cellVal(row, p_sCodeStartCol, 0) + cellVal(row, p_sCodeStartCol, 1)
                    self.pre += cellVal(row, p_sCodeStartCol, 2) + cellVal(row, p_sCodeStartCol, 3)
                    self.pre += cellVal(row, p_sCodeStartCol, 4) + cellVal(row, p_sCodeStartCol, 5)
                    self.pre += cellVal(row, p_sCodeStartCol, 6) + cellVal(row, p_sCodeStartCol, 7)
                    self.pre += cellVal(row, p_sCodeStartCol, 8) + cellVal(row, p_sCodeStartCol, 9)
                    self.pre += cellVal(row, p_sCodeStartCol, 10) + " "                                  #row[_P_].value

                    if self.pre != 'PhaseDisciplineGroupBuildingNumberRevision':
                        StructuralDesignerConvention._dict[self.pre] = row[p_sNameCol.number].value
                except:
                    pass
        except:
            pass

    def getFileName(self, p_source, p_values):
        _fileName = os.path.splitext(p_source)[0]
        _fileExt = os.path.splitext(p_source)[1]

        for pre in StructuralDesignerConvention._dict.keys():
            if _fileName.startswith(pre):
                return pre + StructuralDesignerConvention._dict[pre] + _fileExt
        return p_source

    @staticmethod
    def getIndexByKey(p_key:str)->int:
        return next((p for p in StructuralDesignerConvention.headersList if p.contains(p_key)), 0)
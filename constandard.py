import xlrd

ConTemplatePath = ""
WB_Convention = ""
Sh_Convention = ""

ConTemplatePath = r"Q:\BIM_management\10_BIMgoodies\DATA\LIMA_nevezektan.xls"


#in 1-2 Tervfázis mapping



def UpdateConvention(ConTemplatePath):
    WB_Convention = xlrd.open_workbook(ConTemplatePath, formatting_info=True)
    Sh_Convention = WB_Convention.sheet_by_index(11)



def getPhase(chosenPhase, ConTemplate): 
    WB_Convention = xlrd.open_workbook(ConTemplate, formatting_info=True)
    Sh_Convention = WB_Convention.sheet_by_index(11)
    PhaseKeys = list(filter(lambda x: x != "", Sh_Convention.col_values(1)[1:]))
    PhaseValues = list(filter(lambda x: x != "", Sh_Convention.col_values(2)[1:]))

    D_Phase = {}

    for key in PhaseKeys:
        keyindex = PhaseKeys.index(key)
        value = PhaseValues[keyindex]
        D_Phase[key] = value

    return(D_Phase.get(chosenPhase))




#in 3-4 Épületrész mapping




def getBuilding(chosenBuilding, ConTemplate):
    WB_Convention = xlrd.open_workbook(ConTemplate, formatting_info=True)
    Sh_Convention = WB_Convention.sheet_by_index(11)
    BuildingKeys = list(filter(lambda x: x != "", Sh_Convention.col_values(3)[1:]))
    BuildingValues = list(filter(lambda x: x != "", Sh_Convention.col_values(4)[1:]))

    D_Building = {}

    for key in BuildingKeys:
        keyindex = BuildingKeys.index(key)
        value = BuildingValues[keyindex]
        D_Building[key] = value

    return(D_Building.get(chosenBuilding))



#in 5-6 Szint mapping





def getStoreyDict(chosenStorey, ConTemplate):
    WB_Convention = xlrd.open_workbook(ConTemplate, formatting_info=True)
    Sh_Convention = WB_Convention.sheet_by_index(11)
    StoreyKeys = list(filter(lambda x: x != "", Sh_Convention.col_values(5)[1:]))
    StoreyValues = list(filter(lambda x: x != "", Sh_Convention.col_values(6)[1:]))
    D_Storey = {}

    for key in StoreyKeys:
        keyindex = StoreyKeys.index(key)
        value = StoreyValues[keyindex]
        D_Storey[key] = value

    return(D_Storey.get(chosenStorey))



#in 9-10 Szerepkör mapping

# def getRolesDict(selectedRole):

#     RolesKeys = list(filter(lambda x: x != "", Sh_Convention.col_values(9)[1:]))
#     RolesValues = list(filter(lambda x: x != "", Sh_Convention.col_values(10)[1:]))


#     D_Roles = {}

#     for key in RolesKeys:
#         keyindex = RolesKeys.index(key)
#         value = RolesValues[keyindex]
#         D_Roles[key] = value

def getRolesDict(selectedRole, ConTemplate):
    WB_Convention = xlrd.open_workbook(ConTemplate, formatting_info=True)
    #Sh_Convention = WB_Convention.sheet_by_index(11)
    
    Sh_RoleID = WB_Convention.sheet_by_index(7)

    RealRoleIDKeys = list(filter(lambda x: x != "", Sh_RoleID.col_values(1)[2:]))
    RealRoleIDValues = list(filter(lambda x: x != "", Sh_RoleID.col_values(0)[2:]))

    D_RealRoleID = {}

    for key in RealRoleIDKeys:
        keyindex = RealRoleIDKeys.index(key)
        value = RealRoleIDValues[keyindex]
        D_RealRoleID[key] = value
    return(D_RealRoleID.get(selectedRole))


#in 12-13 Státusz mapping

def getStatus(chosenStatus, ConTemplate):

    StatusKeys = list(filter(lambda x: x != "", Sh_Convention.col_values(12)[1:]))
    StatusValues = list(filter(lambda x: x != "", Sh_Convention.col_values(13)[1:]))


    D_Status = {}

    for key in StatusKeys:
        keyindex = StatusKeys.index(key)
        value = StatusValues[keyindex]
        D_Status[key] = value

    return(D_Status.get(chosenStatus))




#in 7-8 Dokumentumtípus mapping (CSERÉLVE)

def getDocTypeNUM(DocTypeSTR, ConTemplate):

    DocTypeKeys = list(filter(lambda x: x != "", Sh_Convention.col_values(8)[1:]))
    DocTypeValues = list(filter(lambda x: x != "", Sh_Convention.col_values(7)[1:]))


    D_DocType = {}

    for key in DocTypeKeys:
        keyindex = DocTypeKeys.index(key)
        value = DocTypeValues[keyindex]
        D_DocType[key] = value
    return(D_DocType)


#Szerepkör - elérhető DoctypeID mapping

def getAvailableDocID(RoleCode, ConTemplate):


    Sh_RoleID = WB_Convention.sheet_by_index(7)

    RRoleIDKeys = list(filter(lambda x: x != "", Sh_RoleID.col_values(1)[2:]))
    RRoleIDValues = list(filter(lambda x: x != "", Sh_RoleID.col_values(0)[2:]))

    D_RRoleID = {}

    for key in RRoleIDKeys:
        keyindex = RRoleIDKeys.index(key)
        value = RRoleIDValues[keyindex]
        D_RRoleID[key] = value

    # print(D_RRoleID.get(RoleCode))



    Sh_Role = WB_Convention.sheet_by_index(6)

    RoleIDKeys = list(filter(lambda x: x != "", Sh_Role.col_values(1)[2:]))
    RoleIDValues = list(filter(lambda x: x != "", Sh_Role.col_values(2)[2:]))

    D_RoleID = {}

    for key in RoleIDKeys:
        keyindex = RoleIDKeys.index(key)
        value = RoleIDValues[keyindex]
        D_RoleID[key] = value

    RoleID = D_RoleID.get(D_RRoleID.get(RoleCode))

    # print(RoleID)
   

    DocTypeList = list(filter(lambda x: x != "", Sh_Convention.col_values(7)[1:]))

    OutListNUM = []

    DocTypeKeys = list(filter(lambda x: x != "", Sh_Convention.col_values(7)[1:]))
    DocTypeValues = list(filter(lambda x: x != "", Sh_Convention.col_values(8)[1:]))


    D_DocType = {}

    for key in DocTypeKeys:
        keyindex = DocTypeKeys.index(key)
        value = DocTypeValues[keyindex]
        D_DocType[key] = value

    for checkitem in DocTypeList:
        
        if checkitem[:2] == RoleID:
            OutListNUM.append(checkitem)

    D_FlipDocType ={}
    for k, v in D_DocType.items():
        D_FlipDocType[v] = k

    
    OutListSTR = []

    for DocIDNum in OutListNUM:
        D_DocType.get(DocIDNum)
        OutListSTR.append(D_DocType.get(DocIDNum))

        
    return(OutListSTR, D_FlipDocType)




# print(getPhase("Felmérési terv", ConTemplatePath)) 
# print(getRolesDict("BIM tervezés", ConTemplatePath))
# print(getBuilding("A épület", ConTemplatePath))
# print(getStoreyDict("1. emelet", ConTemplatePath))
# print(getAvailableDocID("BIM tervezés", ConTemplatePath))
# print(getStatus("Munkaközi (belsős)", ConTemplatePath))
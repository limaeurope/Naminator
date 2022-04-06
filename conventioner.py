import xlrd
import PySimpleGUI as sg


def ComboBoxT(CBox_name, CBox_list):
    return sg.InputCombo(CBox_list, default_value=CBox_list[0], key=CBox_name, enable_events=True)

def GetConvention(CFile):
        ConventionTemplate = xlrd.open_workbook(CFile)
        WorkSheet = ConventionTemplate.sheet_by_index(11)


        DropVal1 = WorkSheet.col_values(0)
        DropVal2 = WorkSheet.col_values(1)
        DropVal3 = WorkSheet.col_values(3)
        DropVal4 = WorkSheet.col_values(5)
        DropVal5 = WorkSheet.col_values(8)
        DropVal6 = WorkSheet.col_values(9)
        DropVal7 = WorkSheet.col_values(7)
        DropVal8 = WorkSheet.col_values(12)
        DropVal9 = WorkSheet.col_values(15)
        DropVal10 = WorkSheet.col_values(9)

        CBox_List = [WorkSheet.col_values(0)[0], WorkSheet.col_values(1)[0], WorkSheet.col_values(3)[0], WorkSheet.col_values(5)[0], WorkSheet.col_values(7)[0], WorkSheet.col_values(9)[0], WorkSheet.col_values(11)[0], WorkSheet.col_values(12)[0], WorkSheet.col_values(14)[0], WorkSheet.col_values(15)[0]]

        CBox_name1 = "DropVal1"
        CBox_name2 = "DropVal2"
        CBox_name3 = "DropVal3"
        CBox_name4 = "DropVal4"
        CBox_name5 = "DropVal5"
        CBox_name6 = "DropVal6"
        CBox_name7 = "DropVal7"
        CBox_name8 = "DropVal8"
        CBox_name9 = "DropVal9"
        CBox_name10 = "DropVal10"

        CBox_values1 = DropVal1[1:]
        CBox_values2 = DropVal2[1:]
        CBox_values3 = DropVal3[1:]
        CBox_values4 = DropVal4[1:]
        CBox_values5 = DropVal5[1:]
        CBox_values6 = DropVal6[1:]
        CBox_values7 = DropVal7[1:]
        CBox_values8 = DropVal8[1:]
        CBox_values9 = DropVal9[1:]
        CBox_values10 = DropVal10[1:]

        CBox_Master = [[CBox_List],[CBox_values1, CBox_values2, CBox_values3, CBox_values4, CBox_values5, CBox_values6, CBox_values7, CBox_values8, CBox_values9, CBox_values10]]

        #print(CBox_name1)


        Choice1 = ComboBoxT(CBox_name1, CBox_values1)
        Choice2 = ComboBoxT(CBox_name2, CBox_values2)
        Choice3 = ComboBoxT(CBox_name3, CBox_values3)
        Choice4 = ComboBoxT(CBox_name4, CBox_values4)
        Choice5 = ComboBoxT(CBox_name5, CBox_values5)
        Choice6 = ComboBoxT(CBox_name6, CBox_values6)
        Choice7 = ComboBoxT(CBox_name7, CBox_values7)
        Choice8 = ComboBoxT(CBox_name8, CBox_values8)
        Choice9 = ComboBoxT(CBox_name9, CBox_values9)
        Choice10 = ComboBoxT(CBox_name10, CBox_values10)

        CBox_List = [Choice1,
        Choice2,
        Choice3,
        Choice4,
        Choice5,
        Choice6,
        Choice7,
        Choice8,
        Choice9,
        Choice10
        ]

        return(CBox_Master)
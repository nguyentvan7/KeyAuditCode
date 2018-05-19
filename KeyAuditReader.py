from openpyxl import load_workbook

# Return amount of rooms based on building. Rows is rooms + 1.
def building_check(ltr):
    if ltr == "B" or ltr == "C" or ltr == "H":
        return 159
    elif ltr == "D" or ltr == "G":
        return 232
    elif ltr == "E":
        return 231
    elif ltr == "F":
        return 92
    elif ltr == "I":
        return 239
    elif ltr == "J":
        return 213
    else:
        return 149

# Load full key audit.
auditBook = load_workbook(filename = "VDSKeyAudit.xlsx")
auditSheetB = auditBook['B']
auditSheetC = auditBook['C']
auditSheetD = auditBook['D']
auditSheetE = auditBook['E']
auditSheetF = auditBook['F']
auditSheetG = auditBook['G']
auditSheetH = auditBook['H']
auditSheetI = auditBook['I']
auditSheetJ = auditBook['J']
auditSheetK = auditBook['K']

# Setup for reading audit.
columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
roomAmount = building_check('B')

# Reading audit back to user.
for row in range(2, roomAmount):
    for column in columns:
        print(auditSheetB[column + "1"].value + ": " + str(auditSheetB[column + str(row)].value))
    print()
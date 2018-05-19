import openpyxl as xl

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
auditBook = xl.load_workbook(filename = "VDSKeyAudit.xlsx")
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

# Load sorted keycodes (1-1866).
# Rooms are stored in column A, keycodes are stored in column B.
sortedBook = xl.load_workbook(filename = "SortedDecoded.xlsx")
sortedSheet = sortedBook["Vista"]

# Setup for reading audit.
columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']

# Reading audit back to user.
for currentKey in range(1, 1867):
    # Decode building from currentKey.
    currentRoom = sortedSheet["A" + str(currentKey + 1)].value
    currentLetter = currentRoom[3]
    roomAmount = building_check(currentLetter)
    auditSheet = auditBook[currentLetter]
    print("Keycode: " + str(currentKey))
    for row in range(2, roomAmount):
        for column in columns:
            if currentRoom[5:] == auditSheet["A" + str(row)].value:
                print(auditSheet[column + "1"].value + ": " + str(auditSheet[column + str(row)].value))
                selectedRow = row
    verify = input("Do any changes need to be made?\n")
    # Changes need to be made.
    if verify != "":
        choice = input("What field needs to be changed?\n")
        while choice != "done":
            # Changing sparky key amount.
            if choice == "s":
                auditSheet["D" + str(selectedRow)].value = input("How many sparky keys?\n")
            # Changing room key amount.
            elif choice == "r":
                auditSheet["E" + str(selectedRow)].value = input("How many room keys?\n")
            # Changing mail key amount.
            elif choice == "m":
                auditSheet["F" + str(selectedRow)].value = input("How many mail keys?\n")
            # Changing fob amount.
            elif choice == "f":
                auditSheet["G" + str(selectedRow)].value = input("How many fobs?\n")
            # Changing clear/discrepant.
            elif choice == "c":
                change = input("Clear or discrepant?\n")
                if change == "c":
                    auditSheet["H" + str(selectedRow)].value = "Clear"
                else:
                    auditSheet["H" + str(selectedRow)].value = "Discrepant"
            # Changing comments.
            else:
                auditSheet["I" + str(selectedRow)].value = input("Input comment:\n")
            choice = input("What field needs to be changed?\n")
        auditBook.save("VDSKeyAudit.xlsx")
        break
    print()
#auditBook.save("VDSKeyAudit.xlsx")
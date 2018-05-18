from openpyxl import Workbook

# Returns array of room numbers to be used for comparisons later.
def building_check(bldg):
    if bldg == "VDSB":
        result = [1001, 4014, 1, 14]
    elif bldg == "VDSC":
        result = [1015, 4028, 15, 28]
    elif bldg == "VDSD":
        result = [1029, 4046, 29, 46]
    elif bldg == "VDSE":
        result = [1047, 4064, 47, 64]
    elif bldg == "VDSF":
        result = [1065, 4073, 65, 73]
    elif bldg == "VDSG":
        result = [1074, 4091, 74, 91]
    elif bldg == "VDSH":
        result = [1092, 4105, 92, 105]
    elif bldg == "VDSI":
        result = [1106, 4123, 106, 123]
    elif bldg == "VDSJ":
        result = [1124, 7138, 124, 138]
    elif bldg == "VDSK":
        result = [1139, 7149, 139, 149]
    elif bldg == "VAV":
        result = [1201, 4222, 201, 238]
    return result

# Create spreadsheet.
book = Workbook()
sheet = book.active
# Create header.
sheet['A1'] = "Room"
sheet['B1'] = "Keycode"
# Used for roomspaces.
alph = ["A", "B", "C", "D", "C1", "C2"]
building = input("What building?\n")
ranges = building_check(building)
room = ranges[0]
# For VAV.
keyCode = 1
unit = input("What is the unit type for room " + building + "-" + str(room) + "?\n")
# Continuously decode until specified.
while unit != "DONE":
    # Changing buildings when specified.
    if unit == "NEW":
        building = input("What building?\n")
        ranges = building_check(building)
        room = ranges[0]
        unit = input("What is the unit type for room " + building + "-" + str(room) + "?\n")
    if unit != "N" and unit != "" and unit != "4" and unit != "3":
        keyCode = int(input("What is the key code for room " + building + "-" + str(room) + "A?\n"))

    #Checking what unit type and decoding key code appropriately.
    if unit == "A" or unit == "E":
        print(building + "-" + str(room) + " corresponds to key code " + str(keyCode))
        sheet.append([building + "-" + str(room) + "A", keyCode])
    elif unit == "B":
        for letter in alph:
            if letter == "C":
                break
            print(building + "-" + str(room) + letter + " corresponds to key code " + str(keyCode))
            sheet.append([building + "-" + str(room) + letter, keyCode])
            keyCode += 1
    elif unit == "C":
        for letter in alph:
            if letter != "C" and letter != "D":
                print(building + "-" + str(room) + letter + " corresponds to key code " + str(keyCode))
                sheet.append([building + "-" + str(room) + letter, keyCode])
                keyCode += 1
    elif unit == "D" or unit == "4":
        for letter in alph:
            if letter == "C1":
                break
            print(building + "-" + str(room) + letter + " corresponds to key code " + str(keyCode))
            sheet.append([building + "-" + str(room) + letter, keyCode])
            keyCode += 1
    # For VAV.
    elif unit == "3":
        for letter in alph:
            if letter == "D":
                break
            print(building + "-" + str(room) + letter + " corresponds to key code " + str(keyCode))
            sheet.append([building + "-" + str(room) + letter, keyCode])
            keyCode += 1
    room += 1
    # Changing buildings when reaching the end of room numbers for the building.
    if room > ranges[1]:
        building = input("What building?\n")
        ranges = building_check(building)
        room = ranges[0]
    # Incrementing the floor number when reaching the end of the building range, if the floor is still within the range.
    elif int(str(room)[1:]) > ranges[3]:
        room = room + 1000 - (ranges[3] - ranges[2] + 1)
    unit = input("What is the unit type for room " + building + "-" + str(room) + "?\n")
book.save("Decoded.xlsx")
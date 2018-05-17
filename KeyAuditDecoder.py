# Returns array of room numbers to be used for comparisons later.
def building_check(bldg):
    if building == "VDSB":
        result = [1001, 4014, 1, 14]
    elif building == "VDSC":
        result = [1015, 4028, 15, 28]
    elif building == "VDSD":
        result = [1029, 4046, 29, 46]
    elif building == "VDSE":
        result = [1047, 4064, 47, 64]
    elif building == "VDSF":
        result = [1065, 4073, 65, 73]
    elif building == "VDSG":
        result = [1074, 4091, 74, 91]
    elif building == "VDSH":
        result = [1092, 4105, 92, 105]
    elif building == "VDSI":
        result = [1106, 4124, 106, 124]
    elif building == "VDSJ":
        result = [1125, 7138, 125, 138]
    else:
        result = [1139, 7149, 139, 149]
    return result

# Used for roomspaces.
alph = ["A", "B", "C", "D", "C1", "C2"]
building = input("What building?\n")
ranges = building_check(building)
room = ranges[0]
unit = input("What is the unit type for room " + str(room) + "?\n")
# Continuously decode until specified.
while unit != "DONE":
    # Changing buildings when specified.
    if unit == "NEW":
        building = input("What building?\n")
        ranges = building_check(building)
        room = ranges[0]
        unit = input("What is the unit type for room " + building + "-" + str(room) + "?\n")
    keyCode = int(input("What is the key code for room " + building + "-" + str(room) + "?\n"))

    #Checking what unit type and decoding key code appropriately.
    if unit == "A" or unit == "E":
        print(building + "-" + str(room) + " corresponds to key code " + str(keyCode))
    elif unit == "B":
        for letter in alph:
            if letter == "C":
                break
            print(building + "-" + str(room) + letter + " corresponds to key code " + str(keyCode))
            keyCode += 1
    elif unit == "C":
        for letter in alph:
            if letter != "C" and letter != "D":
                print(building + "-" + str(room) + letter + " corresponds to key code " + str(keyCode))
                keyCode += 1
    elif unit == "D":
        for letter in alph:
            if letter == "C1":
                break
            print(building + "-" + str(room) + letter + " corresponds to key code " + str(keyCode))
            keyCode += 1
    room += 1
    # Changing buildings when reaching the end of room numbers for the building.
    if room > ranges[1]:
        building = input("What building?\n")
        ranges = building_check(building)
        room = ranges[0]
    # Incrementing the floor number when reaching the end of the building range, if the floor is still within the range.
    elif int(str(room)[2:]) > ranges[3]:
        room = room + 1000 - (ranges[3] - ranges[2])
    unit = input("What is the unit type for room " + building + "-" + str(room) + "?\n")
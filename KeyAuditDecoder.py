alph = ["A", "B", "C", "D", "C1", "C2"]
building = input("What building?\n")
room = int(input("What is the first room number for " + building + "?\n"))
unit = input("What is the unit type for room " + str(room) + "?\n")
while unit != "DONE":
    if unit == "NEW":
        building = input("What building?\n")
        room = int(input("What is the first room number for " + building + "?\n"))
        unit = input("What is the unit type for room " + building + "-" + str(room) + "?\n")
    keyCode = int(input("What is the first key code for room " + building + "-" + str(room) + "?\n"))
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
    else:
        for letter in alph:
            if letter == "C1":
                break
            print(building + "-" + str(room) + letter + " corresponds to key code " + str(keyCode))
            keyCode += 1
    room += 1
    unit = input("What is the unit type for room " + building + "-" + str(room) + "?\n")
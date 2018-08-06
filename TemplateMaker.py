import openpyxl as xl

# Load decoded keycodes.
codeBook = xl.load_workbook(filename = "SortedDecoded.xlsx")

# Load Vista keycodes.
codeSheet = codeBook["Vista"]

# Create 12-3am template sheet.
# 873 rooms in B-F.
book = xl.Workbook()
sheet = book.active

# Create header.
sheet["A1"] = "Keycode"
sheet["B1"] = "Room"
sheet["C1"] = "# Sparky"
sheet["D1"] = "# Room"
sheet["E1"] = "# Mail"
sheet["F1"] = "# Fob"
currentRow = 2

# Fill in keycode and room.
# Iterate through all keycodes and add B-F codes into sheet.
# 1866 rooms in VDS.
for codeNum in range(1, 1867):
    room = codeSheet["A" + str(codeNum)]
    if room[3] == "B" or room[3] == "C" or room[3] == "D" or room[3] == "E" or room[3] == "F":
        sheet["B" + str(currentRow)] = room
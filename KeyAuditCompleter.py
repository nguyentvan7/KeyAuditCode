import openpyxl as xl
import datetime

# Return building letter and row index in Excel sheet for specified keycode.
def room_find(keycode, cChoice):
	if cChoice == '1' or cChoice == '3':
		if keycode in bCodes:
			return 'B' + str(bCodes.index(keycode))
		if keycode in cCodes:
			return 'C' + str(cCodes.index(keycode))
		if keycode in dCodes:
			return 'D' + str(dCodes.index(keycode))
		if keycode in eCodes:
			return 'E' + str(eCodes.index(keycode))
		if keycode in fCodes:
			return 'F' + str(fCodes.index(keycode))
		if keycode in gCodes:
			return 'G' + str(gCodes.index(keycode))
		if keycode in hCodes:
			return 'H' + str(hCodes.index(keycode))
		if keycode in iCodes:
			return 'I' + str(iCodes.index(keycode))
		if keycode in jCodes:
			return 'J' + str(jCodes.index(keycode))
		if keycode in kCodes:
			return 'K' + str(kCodes.index(keycode))
	else:
		return 'L' + str(lCodes.index(keycode))

# Keycodes of each building.
bCodes = list(range(1549, 1708))
cCodes = list(range(1708, 1867))
dCodes = list(range(1158, 1390))
eCodes = list(range(804, 809)) + list(range(926, 1128)) + list(range(1133, 1158))
fCodes = list(range(834, 926))
gCodes = list(range(602, 804)) + list(range(809, 834)) + list(range(1128, 1133))
hCodes = list(range(1390, 1549))
iCodes = list(range(363, 602))
jCodes = list(range(1, 5)) + list(range(150, 174)) + list(range(176, 184)) + list(range(186, 363))
kCodes = list(range(5, 150)) + [174, 175, 184, 185]
lCodes = list(range(1, 401))
		
# Load audit book.
fillBook = xl.load_workbook(filename = '../Fill.xlsx')

# Setup for reading audit.
columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']

# Determine what buildings to anticipate.
complexChoice = input('\nPlease choose a complex:\n1. Vista\n2. Villas\n3. Both\n\n')
while complexChoice != '1' and complexChoice != '2' and complexChoice != '3':
	print('\nInvalid choice, please input 1, 2 or 3\n')
	complexChoice = input('\n1. Vista\n2. Villas\n3. All?\n\n')

if (complexChoice == '1' or complexChoice == '3'):
	print('\nUsing Vista keycodes.')
else:
	print('\nUsing Villas Keycodes.')

inputKey = input(
	'\nPlease begin entering keycodes.\n'
	'Enter h for help. Enter x to exit.\n\n')
			
# Continuously read for input.
while inputKey != 'x' or complexChoice == 3:
	# Print help menu.
	if inputKey == 'h':
		print(
			'\nFormat of keycodes:\n'
			'INPUT = KEYCODE. .COMMENT\n'
			'KEYCODE = NUM.NUM.NUM | NUM.NUM.NUM.NUM\n'
			'COMMENT = r | m | f | u | o | c | n | COMMENT.COMMENT | EPSILON\n'
			'\n'
			'Comment definitions:\n'
			'r = room\n'
			'm = mail\n'
			'f = fob\n'
			'u = urgent\n'
			'o = lock out\n'
			'c = lock change\n'
			'n = reset normal keys\n'
			's = reset studio keys\n'
			'none = full set\n')

	# inputKey is only numbers, so a full set.
	elif inputKey.isdigit():
		# room contains:
		# room[0] -> building letter.
		# int(room[1:]) -> row number in fillSheet for room.
		# Must add 2 to int(room[1:]) to compensate for 0-indexing and header in Excel.
		try:
			room = room_find(int(inputKey), complexChoice)
			fillSheet = fillBook[room[0]]

			# Determine if room is a studio.
			if fillSheet['G' + str(int(room[1:]) + 2)].value == 'Studio':
				# Fill in full set for studio.
				fillSheet['E' + str(int(room[1:]) + 2)].value = 1
				fillSheet['F' + str(int(room[1:]) + 2)].value = 1
			else:
				# Fill in full set for regular room.
				fillSheet['D' + str(int(room[1:]) + 2)].value = 1
				fillSheet['E' + str(int(room[1:]) + 2)].value = 1
				fillSheet['F' + str(int(room[1:]) + 2)].value = 1
		except TypeError:
			print('Please enter a valid keycode/comment.\n')

		# inputKey has comments attached.
	elif inputKey != 'x':
		try:
			room = room_find(int(inputKey[0:inputKey.index(' ')]), complexChoice)
			fillSheet = fillBook[str(room[0])]
			comments = inputKey[inputKey.index(' ') + 1:]
			# Iterate through comments.
			for comment in comments:
				# Set room key.
				if comment == 'r':
					fillSheet['D' + str(int(room[1:]) + 2)].value = 1
				# Set mail key.
				elif comment == 'm':
					fillSheet['E' + str(int(room[1:]) + 2)].value = 1
				# Set fob.
				elif comment == 'f':
					fillSheet['F' + str(int(room[1:]) + 2)].value = 1
				# Set urgent.
				elif comment == 'u':
					fillSheet['C' + str(int(room[1:]) + 2)].value = 0
					fillSheet['G' + str(int(room[1:]) + 2)].value = "Urgent"
				# Set lock out.
				elif comment == 'o':
					fillSheet['C' + str(int(room[1:]) + 2)].value = 0
					fillSheet['G' + str(int(room[1:]) + 2)].value = "Lock out"
				# Set lock change.
				elif comment == 'c':
					fillSheet['C' + str(int(room[1:]) + 2)].value = 0
					fillSheet['G' + str(int(room[1:]) + 2)].value = "Lock change"
				# Set normal keys.
				elif comment == 'n':
					fillSheet['C' + str(int(room[1:]) + 2)].value = 1
					fillSheet['D' + str(int(room[1:]) + 2)].value = 0
					fillSheet['E' + str(int(room[1:]) + 2)].value = 0
					fillSheet['F' + str(int(room[1:]) + 2)].value = 0
					fillSheet['G' + str(int(room[1:]) + 2)].value = ""
				# Set studio keys.
				elif comment == 's':
					fillSheet['C' + str(int(room[1:]) + 2)].value = 0
					fillSheet['D' + str(int(room[1:]) + 2)].value = 0
					fillSheet['E' + str(int(room[1:]) + 2)].value = 0
					fillSheet['F' + str(int(room[1:]) + 2)].value = 0
					fillSheet['G' + str(int(room[1:]) + 2)].value = "Studio"
				# Set full set for accidental whitespace.
				elif comment.isspace():
					if fillSheet['G' + str(int(room[1:]) + 2)].value == 'Studio':
						# Fill in full set for studio.
						fillSheet['E' + str(int(room[1:]) + 2)].value = 1
						fillSheet['F' + str(int(room[1:]) + 2)].value = 1
					else:
						# Fill in full set for regular room.
						fillSheet['D' + str(int(room[1:]) + 2)].value = 1
						fillSheet['E' + str(int(room[1:]) + 2)].value = 1
						fillSheet['F' + str(int(room[1:]) + 2)].value = 1
				else:
					print(
						'One or more of your comments was not recognized.\n'
						'Please reset this room and re-enter your comments.\n')
		except ValueError:
			print('Please enter a valid keycode/comment.\n')
		except TypeError:
			print('Please enter a valid keycode/comment.\n')
	
	# Switch from Vista to Villas for doing all keys.
	if complexChoice == '3' and inputKey == 'x':
		complexChoice = '2'
		print("\n Switching to Villas keycodes.\n\n")

	inputKey = input()

fillBook.save("../Completed/" + datetime.datetime.today().strftime('%y%m%d') + "KeyAudit.xlsx")
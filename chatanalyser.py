with open("../_chat.txt", 'r') as file:
	Mayankwords=0
	Otherwords=0
	for line in file:
		if 'end-to-end' not in line.split():
			if 'Mayank:' in line.split():
				Mayankwords+=1
			else:
				Otherwords+=1

	print(f"Mayank wrote {Mayankwords} lines")
	print(f"Other Person wrote {Otherwords} lines")

with open('prob16-1-in.txt', 'r') as f:
	die = [f.readline() for x in range(int(f.readline()))]
	for x in f.readlines()[1:]:
		x = x.strip()
		temp = die[:]
		fin = 'can'
		for y in x:
			for z in temp:
				if y in z:
					temp.remove(z)
					break
			else:
				fin = 'CANNOT'
				break
		print(f"{x} {fin} be formed")

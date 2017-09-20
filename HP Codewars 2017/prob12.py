with open('prob12-1-in.txt', 'r') as f:
	a = f.read().split('\n')[1:]
	for x in range(0, len(a)-1, 2):
		tutor = list(map(int, a[x].split()[1:]))
		tutee = list(map(int, a[x+1].split()[1:]))
		if tutor[-1] <= tutee[0]:
			print('0')
			continue
		elif tutor[0] > tutee[-1]:
			print(len(tutee)*len(tutor))
			continue
		else:
			s = 0
			end = -1
			for y in range(len(tutor)):
				for z in range(end+1, len(tutee), 1):
					if tutor[y] <= tutee[z]:
						s += z
						end = z-1
						break
			print(s)

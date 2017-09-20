with open('prob13-2-in.txt', 'r') as f:
	a = f.read()[:-1]
	d = ['u', 'd', 'l', 'r']
	x, y, e = 0, 0, 'r'
	word = []
	for z in a:
		word.append([z, x, y])
		char = z.lower()
		if char in d:
			e = char
			if char == 'r':
				x += 1
			if char == 'l':
				x -= 1
			if char == 'u':
				y += 1
			if char == 'd':
				y -= 1
		else:
			if e == 'r':
				x += 1
			if e == 'l':
				x -= 1
			if e == 'u':
				y += 1
			if e == 'd':
				y -= 1
	del x, y, e, d
	left, down = min([z[1] for z in word]), min([z[2] for z in word])
	for z in word:
		z[1] -= left
		z[2] -= down
	word = sorted(word, key=lambda q: (q[2], -q[1]), reverse=True)
	x = [[z for z in word if z[2] == x] for x in range(word[0][2], -1, -1)]
	s = ''
	for z in x:
		for y in range(z[-1][1]+1):
			for g in z:
				if g[1] == y:
					s += g[0]
					break
			else:
				s += ' '
		s += '\n'
	print(s)

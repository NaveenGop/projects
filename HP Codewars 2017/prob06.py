with open('prob06-1-in.txt', 'r') as f:
	for x in f.readlines()[1:]:
		a = list(map(int, x.split()[1:]))
		b = a[0]
		a = [a[y+1]-a[y] for y in range(len(a)-1)]
		a = list(map(lambda q: 0-q, a))
		c = [b]
		for y in range(len(a)):
			c.append(c[y] + a[y])
		print(' '.join(map(str, c)))

with open('prob05-1-in.txt', 'r') as f:
	a = []
	for x in f.readlines()[1:]:
		b = x.split()
		a.append((b[0], float(b[1])/float(b[2])))
	c = sorted(a, key=lambda q: q[1])
	print(c[0][0], c[0][1])

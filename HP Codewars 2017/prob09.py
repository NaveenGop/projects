with open('prob09-2-in.txt', 'r') as f:
	a = [tuple(map(int, x.split())) for x in iter(f.readline, '0 0\n')]
	height = max(a, key=lambda q: q[1])[1]
	width = max(a, key=lambda q: q[0])[0]
	width = width + (10-width%10)
	pic = [list(' '*width) for x in range(height)]
	for x in a[::-1]:
		start = x[0]-x[1]
		end = x[0]+x[1]
		for y in range(x[1], 0, -1):
			row = y-x[1]-1
			if start+x[1]-y >= 0:
				pic[row][start+x[1]-y] = '/'
			pic[row][end-(x[1]-y)-1] = '\\'
			if start + (x[1] - y + 1) >= 0:
				pic[row][start+(x[1]-y+1):end-(x[1]-y)-1] = ' ' * (end-2 - start - 2*(x[1]-y))
			else:
				pic[row][0:end-(x[1]-y)-1] = ' ' * (end - (x[1]-y)-1)

	for x in pic:
		print(''.join(x))
	print('1234567890'*(width//10+1))

with open('prob11-1-in.txt', 'r') as f:
	noise = int(f.readline().split()[2])
	length = list(map(int, f.readline().split()))[1:-1]
	char = ''.join([x.replace(' ', '').rstrip() for x in f.readlines()])
	a, b = set(list(char)), {}
	for x in a:
		b[x] = char.count(x)
	for x in b:
		if b[x] >= noise:
			char = char.replace(x, '')
	char = list(char)
	for x in range(1, len(length)):
		length[x] = length[x-1]+length[x]+1
	for x in length:
		char.insert(x, ' ')
	print(''.join(char))

import re
with open('prob10-2-in.txt', 'r') as f:
	for x in f.readlines()[1:]:
		a = [(y, x[y]) for y in range(len(x))]
		regex = re.compile('[\w]')
		for y in range(len(a)):
			if not regex.search(a[y][1]):
				a[y] = None
		a = [y for y in a if y is not None]
		c = ''.join((y[1].lower() for y in a))
		final = []
		for y in range(len(a)):
			for z in range(len(a), y+1, -1):
				if c[y:z] == ''.join(reversed(c[y:z])):
					final.append(a[y:z])
		if len(final) == 0:
			print('NO PALINDROME')
			continue
		else:
			final = sorted(final, key=lambda q: len(q))[-1]
			print(x[final[0][0]:final[-1][0]+1])

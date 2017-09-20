with open('prob07-1-in.txt') as f:
	for x in f.readlines()[1:]:
		a = list(map(int, x.split()[:-1]))
		a.append(float(x.split()[-1]))
		d = a[0]
		for y in range(100):
			e = a[3] + (a[1]*d + a[4])/(a[2]*d + a[5])
			if abs(d-e) < a[-1]:
				print(e)
				break
			d = e
		else:
			print('DIVERGES')

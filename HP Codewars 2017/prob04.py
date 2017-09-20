with open('prob04-1-in.txt', 'r') as f:
	for x in f.readlines()[1:]:
		y = list(map(float, x.split()))
		print(y[0]/(y[1]/60))
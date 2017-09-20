with open('prob02-1-in.txt', 'r') as f:
	a = list(map(int, f.readline().split()))
	print(a[0]*a[1])
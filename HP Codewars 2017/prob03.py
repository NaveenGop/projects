with open('prob03-1-in.txt', 'r') as f:
	for x in f.readlines():
		print(int(x)*0.299792)
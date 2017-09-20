for x in range(1, 24):
	with open(f'prob0{x}.py' if x < 10 else f'prob{x}.py', 'x') as f:
		f.write(f"with open('prob{x if x >= 10 else '0'+str(x)}-1-in.txt', 'r') as f:\n\t")

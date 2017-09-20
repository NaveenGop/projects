with open('prob15-2-in.txt', 'r') as f:
	die = {	'L': [1, 2, 3, 4, 5, 6],
			'F': [0, 1, 1, 2, 3, 5],
			'T': [0, 0, 1, 1, 2, 4],
			'N': [1, 2, 2, 3, 3, 4],
			'Z': [1, 1, 1, 2, 2, 3],
			'P': [0, 1, 2, 3, 5, 7],
			'G': [1, 2, 3, 3, 4, 5]}

	for x in f.readlines()[:-1]:
		dice, total = x[:5].split(), list(map(int, x[6:].split()))
		total_chart = [die[y] for y in dice if y in die.keys()]
		sums = []
		first = total_chart[0]
		second = total_chart[1]
		if len(total_chart) == 3:
			third = total_chart[2]
			sums = [a + b + c for a in first for b in second for c in third]
		else:
			sums = [a + b for a in first for b in second]
		sums = [(y, sums.count(y)) for y in set(sums)]
		unique = sorted(sums, key=lambda q: q[0] in total, reverse=True)[:3]
		count = sum([y[1] for y in sums])
		print(' '.join(dice))
		for y in unique:
			print(f"{y[0]}  {y[1]}    {y[1]*100/count}%")

with open('prob18-2-in.txt', 'r') as f:
	p = int(f.readline())
	poke = []
	for x in range(p):
		a = f.readline().split()
		poke.append({'name': a[0], 'type': a[1],
					 'weakness': a[2], 'resistance': a[3],
					 'hp': int(a[4]), 'attack1': {'name': a[5], 'charge': int(a[6]), 'damage': int(a[7])},
					 'attack2': {'name': a[8], 'charge': int(a[9]), 'damage': int(a[10])}})
	p = int(f.readline())
	for x in range(p):
		a = f.readline().split()
		first = [y for y in poke if y['name'] == a[0]][0]
		second = [y for y in poke if y['name'] == a[1]][0]
		o1, o2 = first['hp'], second['hp']
		turn, first['power'], second['power'] = 1, 0, 0
		w1, w2 = 2 if first['weakness'] == second['type'] else 1, 2 if second['weakness'] == first['type'] else 1
		r1, r2 = 0.5 if first['resistance'] == second['type'] else 1, 0.5 if second['resistance'] == first['type'] else 1
		while first['hp'] * second['hp'] > 0:
			first['power'] += 1
			second['power'] += 1
			if turn % 2:

				if first['power'] >= first['attack2']['charge']:
					first['power'] -= first['attack2']['charge']
					second['hp'] -= first['attack2']['damage'] * w2 * r2
				elif first['power'] >= first['attack1']['charge']:
					first['power'] -= first['attack1']['charge']
					second['hp'] -= first['attack1']['damage'] * w2 * r2
			else:

				if second['power'] >= second['attack2']['charge']:
					second['power'] -= second['attack2']['charge']
					first['hp'] -= second['attack2']['damage'] * w1 * r1
				elif second['power'] >= second['attack1']['charge']:
					second['power'] -= second['attack1']['charge']
					first['hp'] -= second['attack1']['damage'] * w1 * r1

			print(turn, first['hp'], second['hp'])
			turn += 1
		if first['hp'] > 0:
			winner = first
			loser = second['name']
		else:
			winner = second
			loser = first['name']
		print(f"{winner['name']} defeats {loser} with {winner['hp']} HP remaining.")
		first['hp'], second['hp'] = o1,o2
	y = 2 + 2

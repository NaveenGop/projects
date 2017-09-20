import re


def roman(z: str) -> int:
	s = 0.5 if z.endswith('S') else 0

	s += 4 * z.count('IV')
	z = z.replace('IV', '')

	s += 9 * z.count('IX')
	z = z.replace('IX', '')

	s += 40 * z.count('XL')
	z = z.replace('XL', '')

	s += 90 * z.count('XC')
	z = z.replace('XC', '')

	s += 400 * z.count('CD')
	z = z.replace('CD', '')

	s += 900 * z.count('CM')
	z = z.replace('CM', '')

	s += z.count('I') + z.count('V') * 5 + z.count('X') * 10 + z.count('L') * 50 + z.count('C') * 100 + z.count(
		'D') * 500 + z.count('M') * 1000
	return s

# TODO
with open('prob17-2-in.txt', 'r') as f:
	regex = re.compile('M{0,2}C?M?C?D?C{0,2}X?C?X?L?X{0,2}I?X?I?V?I{0,3}S?')
	strip = re.compile('[^MDCLXVIS]')
	for x in f.readlines()[:-1]:
		x = strip.sub('', x.upper())
		s = sorted(regex.findall(x), key=roman)[-1]
		print(s, roman(s))
from urllib.request import urlopen
from bs4 import BeautifulSoup, element
from multiprocessing import Pool, freeze_support
from datetime import datetime
from csv import DictWriter, DictReader
import re, os


def extract(url: str, car_list: set) -> tuple:
	temp, info = [], []
	page_soup = BeautifulSoup(urlopen(url).read(), 'html.parser')
	containers = page_soup.findAll("li", {"class": "result-row"})
	first_regex = re.compile('\\d{4}(?!\\d)|\\d{2}(?=\\s+)(?![kK])')
	regex = re.compile('\\d{4}(?!\\d)|\\d{2}(?=\\s+)(?![kK])')
	for x in containers:
		z = ([a for a in y.contents if type(a) is element.Tag] for y in (b for b in x.contents if type(b) is element.Tag)
			if list(y.attrs.values()) == [['result-info']])
		for l in z:
			data = {}
			for m in l:
				if 'datetime' in m.attrs.keys():
					data['date'] = m.attrs['datetime']
				elif 'href' in m.attrs.keys():
					data['link'] = 'https://sacramento.craigslist.org' + m.attrs['href']
					ps = BeautifulSoup(urlopen(data['link']).read(), 'html.parser')
					try:
						for dt in (u.text for u in filter(lambda q: type(q) is element.Tag, ps.findAll("p", {"class": "attrgroup"})[-1].contents)):
							try:
								a = dt.split(': ')
								data[a[0]] = a[1]
							except IndexError:
								continue
					except IndexError:
						pass
					data['title'] = m.text.encode(encoding='cp1252', errors='replace').decode(errors='ignore')
					match = first_regex.match(m.text)
					if match is None:
						match = regex.search(m.text)
					if match:
						b = match.group(0)
						data['year'] = int(b) if len(b) == 4 else int(('20' if b.startswith(('1', '0')) else '19') + b)
					if 'mustang' in m.text.lower():
						data['make'] = 'Ford'
					else:
						a = car_list.intersection(m.text.lower().split())
						if len(a) == 1:
							data['make'] = list(a)[0]
				elif m.attrs['class'] == ['result-meta']:
					t = (x for x in m.contents if type(x) is element.Tag)
					t = (x for x in t if x.attrs['class'] == ['result-price'])
					try:
						data['price'] = next(t).text
					except StopIteration:
						data['price'] = 'NOT INDICATED'
			info.append(data)
		temp.extend(z)
	return temp, info


def get_cars(url: str ='http://www.listingallcars.com') -> set:
	p_soup = BeautifulSoup(urlopen(url).read(), 'html.parser')
	ponts = p_soup.findAll("table", {"id": "allmakeListPhone"})[0].contents
	car_list = {x.text[:x.text.find('(')].strip().lower() for x in ponts[1:-1]}
	car_list.update(["chevy", "vw"])
	return car_list


def get_total_sale(url: str) -> int:
	return int(BeautifulSoup(urlopen(url).read(), 'html.parser').find("span", {"class": "totalcount"}).text)


def main():
	print("Start Time:", datetime.now())
	sale_list = []
	my_url = 'https://sacramento.craigslist.org/search/cta'
	with Pool() as pool:
		car_async = pool.apply_async(get_cars)
		sale_async = pool.apply_async(get_total_sale, (my_url,))
		car_list = car_async.get()
		# total_sales = sale_async.get()
		sale_list.append(extract(my_url, car_list)[1])
		cpu = os.cpu_count()
		for x in range(120, sale_async.get(), 120*cpu):
			pools = [pool.apply_async(extract, (my_url + "?s={}".format(x*y), car_list)) for y in range(1, cpu+1)]
			sale_list.extend([z.get()[1] for z in pools])
		# for x in range(120, total_sales, 480):
		# 	first = pool.apply_async(extract, (my_url + "?s={}".format(x), car_list))
		# 	second = pool.apply_async(extract, (my_url + "?s={}".format(x+120), car_list))
		# 	third = pool.apply_async(extract, (my_url + "?s={}".format(x+240), car_list))
		# 	fourth = pool.apply_async(extract, (my_url + "?s={}".format(x+360), car_list))
		# 	sale_list.extend((first.get()[1], second.get()[1], third.get()[1], fourth.get()[1]))

	print("Process Time:", datetime.now())
	d = set()
	if os.path.isfile('cars.csv'):
		with open('cars.csv', 'r', newline='') as csvfile:
			reader = DictReader(csvfile)
			d = {(row['title'], row['date'], row['link']) for row in reader}

	with open('f.csv', 'a', newline='') as csvfile:
		fieldnames = ['title', 'date', 'make', 'year', 'price', 'VIN', 'fuel', 'odometer', 'cylinders', 'paint color',
					'title status', 'transmission', 'type', 'link']
		writer = DictWriter(csvfile, fieldnames=fieldnames, restval='NOT INDICATED', extrasaction='ignore')

		if not d:
			writer.writeheader()
		for x in sale_list:
			for y in x:
				if d:
					if (y['title'], y['date'], y['link']) not in d:
						writer.writerow(y)
					else:
						continue
				else:
					writer.writerow(y)
	print("End Time:", datetime.now())


if __name__ == '__main__':
	freeze_support()
	main()

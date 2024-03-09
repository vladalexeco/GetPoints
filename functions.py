from lxml import etree as et
import xlwt, xlrd
import re
import pyproj
import datetime
import os, sys 
from decimal import Decimal, ROUND_HALF_UP
from journal_point import JournalPoint
import random

def getFormatedDateAndTime(journalPoint):
	splitDate = journalPoint.date.split(".")
	splitDate.reverse()
	formatedDate = "-".join(splitDate)

	return f"{formatedDate}T{journalPoint.time}Z"

def xls_journal_to_gpx(a_list, address, createComments = False):
	
	file = a_list[0]
	book = xlrd.open_workbook(file)
	sheet = book.sheet_by_index(0)

	result = []

	for i in range(1, sheet.nrows):
		
		currentList = []
		
		for j in range(sheet.ncols):
			currentList.append(sheet.cell(i, j).value)

	
		if createComments:

			result.append(JournalPoint(
				id = int(currentList[0]), 
				latitude = coordMinToCoordFract(currentList[1].replace(",", ".")), 
				longitude = coordMinToCoordFract(currentList[2].replace(",", ".")), 
				search = currentList[3],
				mad = currentList[4],
				madError = round(currentList[5], 3),
				date = currentList[6],
				time = currentList[7]
				))

		else:

			result.append(JournalPoint(
				id = int(currentList[0]),
				latitude = coordMinToCoordFract(currentList[1]),
				longitude = coordMinToCoordFract(currentList[2]),
				search = "",
				mad = "",
				madError = "",
				date = currentList[3],
				time = currentList[4]
				))

	
	tree = et.parse('templates/gpx_template.gpx')
	root = tree.getroot()

	metadata = root.find("metadata")

	bounds = metadata.find("bounds")

	if bounds is not None:
		metadata.remove(bounds)

	link = et.SubElement(metadata, "link", href="http://www.garmin.com")
	textLink = et.SubElement(link, "text")
	textLink.text = "Garmin International"

	time = metadata.find("time")

	if time is not None:
		metadata.remove(time)

	time = et.SubElement(metadata, "time")

	firstJournalPoint = result[0]

	time.text = getFormatedDateAndTime(firstJournalPoint)

	for journalPoint in result:
		
		wpt = et.SubElement(root, "wpt", lat=str(journalPoint.latitude), lon=str(journalPoint.longitude))
		
		ele = et.SubElement(wpt, "ele")
		ele.text = str(round(random.randint(-2, 6) + random.random(), 6))

		time = et.SubElement(wpt, "time")
		time.text = getFormatedDateAndTime(journalPoint)

		name = et.SubElement(wpt, "name")
		name.text = str(journalPoint.id)

		if createComments:
			cmt = et.SubElement(wpt, "cmt")
			cmt.text = f"Поиск {journalPoint.search} МАЭД {journalPoint.mad} Ошибка МАД {journalPoint.madError}"

		sym = et.SubElement(wpt, "sym")
		sym.text = "Block, Red"

	tree.write(address, encoding='utf-8')

	with open(address, 'r', encoding='utf-8') as file:
		list_file = file.readlines()
		list_file[0] = '<gpx xmlns="http://www.topografix.com/GPX/1/1" xmlns:gpxx="http://www.garmin.com/xmlschemas/GpxExtensions/v3" xmlns:wptx1="http://www.garmin.com/xmlschemas/WaypointExtension/v1" xmlns:gpxtpx="http://www.garmin.com/xmlschemas/TrackPointExtension/v1" creator="GPSMAP 64" version="1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd http://www.garmin.com/xmlschemas/GpxExtensions/v3 http://www8.garmin.com/xmlschemas/GpxExtensionsv3.xsd http://www.garmin.com/xmlschemas/TrackStatsExtension/v1 http://www8.garmin.com/xmlschemas/TrackStatsExtension.xsd http://www.garmin.com/xmlschemas/WaypointExtension/v1 http://www8.garmin.com/xmlschemas/WaypointExtensionv1.xsd http://www.garmin.com/xmlschemas/TrackPointExtension/v1 http://www.garmin.com/xmlschemas/TrackPointExtensionv1.xsd">\n' 
		list_file.insert(0, '<?xml version="1.0" encoding="UTF-8" standalone="no" ?>\n')

	with open(address, 'w', encoding='utf-8') as file:
		for string in list_file:
			file.write(string)

	format_file(address)

	
def coordMinToCoordFract(coord, separators = ["º", "\'", "\""]):
	"""Transform coordinate with minutes and seconds to coordinate with fraction part"""
	listOfCoord = coordMinStrToList(coord, separators)
	coordFract = float(listOfCoord[0]) + float(listOfCoord[1]) / 60 + float(listOfCoord[2]) / 3600
	return round(coordFract, 7)


def coordMinStrToList(coord, separators):
	"""Transform string coordinate with minutes and seconds to list. Example 36°24'59.124'' converts to [36, 24, 59.124]"""
	degree, rightPart = coord.split(separators[0])
	minute, second = rightPart.split(separators[1], 1)
	second = second.split(separators[2])[0]
	return [degree, minute, second]

def trueRound(num):
	decimal_num = Decimal(str(num))
	for i in range(7, 1, -1):
		decimal_num = decimal_num.quantize(Decimal("1." + "1" * i), rounding = ROUND_HALF_UP)
	return float(decimal_num)


def check_extension(ex, a_list):
	for path in a_list:
		extension = os.path.splitext(path)[-1]
		if extension != ex:
			return False
	return True


def coordList(coord):
	'''Converts geografic coordinates with fractions parts to coordinates with minutes and seconds'''
	coord = float(coord)
	degree = int(coord)
	fract_part = coord - degree
	minutes_fract = fract_part * 60
	minutes = int(minutes_fract)
	# sec = round((minutes_fract - int(minutes_fract)) * 60, 2)
	sec = (minutes_fract - int(minutes_fract)) * 60
	return [degree, minutes, round(sec, 2)]


def  get_net_adress(string):
	a_list = string.split('}')
	return a_list[0] + '}'


def gpx_to_dict(file):
	'''Converts gpx file to python dictionary'''
	total = dict() 
	order = 0

	tree = et.parse(file)
	root = tree.getroot()

	net_adress = get_net_adress(root.tag)

	for point in root.findall(net_adress + 'wpt'):
		order += 1
		dict_point = dict()
		dict_point['coord'] = point.attrib
		for obj in point:
			name = obj.tag.split('}')[1]
			if name == 'time':
				a_list = obj.text.split('T')
				dict_point['date'] = a_list[0]
				dict_point['time'] = a_list[1][:-1]
			else:
				dict_point[name] = obj.text
				total[str(order)] = dict_point
	return total


def kml_to_dict(file):
	'''Converts kml file to python list'''
	dom = et.parse(file)

	root = dom.getroot()

	net_adress = get_net_adress(root.tag)

	folder_with_points = []

	total = []

	for elem in root.findall(".//" + net_adress + "Folder"):
		if net_adress + "Placemark" in [x.tag for x in elem.getchildren()]:
			folder_with_points.append(elem)

	for folder in folder_with_points:
		for placemark in folder.findall(net_adress + 'Placemark'):
			dict_point = {}
			dict_point['name'] = placemark.find(net_adress + 'name').text
			description = placemark.find(net_adress + 'description')
			dict_point['cmt'] = description.text if description is not None else ''
			point = placemark.find(net_adress + 'Point')
			temp_coords = point.find(net_adress + 'coordinates').text.split(',')
			dict_point['lat'] = temp_coords[1] 
			dict_point['lon'] = temp_coords[0]
			
			if placemark.find(net_adress + 'end') is not None:
				time = placemark.find(net_adress + 'end')[0][0][0].text 
				dict_point['time'] = time
			else:
				now = datetime.datetime.now()
				dict_point['time'] = now.strftime("%Y-%m-%dT%H:%M:%SZ")
			
			dict_point['ele'] = '0.000000' 

			total.append(dict_point)

	return total 
	

def format_to(a_list, param=10):
	'''Formate fraction part of float numbers'''

	for i in range(len(a_list)):
		for  j in range(len(a_list[i])):
			if type(a_list[i][j]) == float:
				int_part, fract_part = str(a_list[i][j]).split('.')
				num_of_zero  = param - len(fract_part)
				if num_of_zero >= 0:
					fract_part = fract_part +  num_of_zero * '0'
					a_list[i][j] = int_part + '.' + fract_part
				else:
					new_num = str(round(a_list[i][j], param))
					int_part, fract_part = new_num.split('.')
					if len(fract_part) < param:
						num_of_zero = param - len(fract_part)
						fract_part = fract_part + num_of_zero * '0'
						a_list[i][j] = int_part + '.' + fract_part
					else:
						a_list[i][j] = new_num


def global_dict(a_list, for_exl=True, for_transpotation=False):
	'''Add dictionares in total dictionary'''

	symb = ['\u00BA', "\'",  '\"']

	list_of_dict = list()

	if check_extension('.gpx', a_list):
		for file in a_list:
			list_of_dict.append(gpx_to_dict(file))
	elif check_extension('.kml', a_list):
		for file in a_list:
			list_of_dict.append(kml_to_dict(file))
		
		temp_main_list = []
		for lst in list_of_dict:
			temp_dictionary = {}
			
			for ind, dct in enumerate(lst):
				temp_dictionary[ind + 1] = dct 

			temp_main_list.append(temp_dictionary)

		list_of_dict = temp_main_list 	

	else:
		raise Exception("wrong file format")


	if for_transpotation and check_extension('.gpx', a_list):
		return list_of_dict # ??????

	new_dict = list_of_dict[0]
	tail_dict =  list_of_dict[1:]

	for dictionary in tail_dict:
		num = len(new_dict)
		for elem in dictionary:
			num += 1
			new_dict[str(num)] = dictionary[elem]


	if for_exl and check_extension('.gpx', a_list):
		for num in new_dict:
			coordinates = new_dict[num]['coord']
			lat_list = coordList(coordinates['lat'])
			lon_list = coordList(coordinates['lon'])

			lat = ''
			lon = ''

			for x, y in zip(lat_list, symb):
				lat += str(x) + y 

			for x, y in zip(lon_list, symb):
				lon += str(x) + y 

			new_dict[num]['lat'] = lat
			new_dict[num]['lon'] = lon
			new_dict[num].pop('coord')
			if not 'cmt' in new_dict[num]:
				new_dict[num]['cmt'] = ''
			if not 'time' in new_dict[num]:
				new_dict[num]['time'] = ''
			if not 'date' in new_dict[num]:
				new_dict[num]['date'] = ''

	if for_exl and check_extension('.kml', a_list):
		for num in new_dict:
			lat_list = coordList(new_dict[num]['lat'])
			lon_list = coordList(new_dict[num]['lon'])

			lat, lon = ('', '')

			for x, y in zip(lat_list, symb): lat += str(x) + y

			for x, y in zip(lon_list, symb): lon += str(x) + y 

			new_dict[num]['lat'] = lat 
			new_dict[num]['lon'] = lon 

			time_date_list = new_dict[num]['time'].split('T')
			new_dict[num]['date'] = time_date_list[0]
			new_dict[num]['time'] = time_date_list[1][:-1]

	return new_dict


def gpx_to_excel(a_list, address, headers, list_of_param, comment=True, srt=True):
	'''Converts list with filenames to table excel'''

	if comment:
		headers_temp = headers
		list_of_param_temp = list_of_param
	else:
		headers_temp = headers[:-1]
		list_of_param_temp = list_of_param[:-1]

	

	new_dict = global_dict(a_list)

	new_list = list(new_dict.values())

	if srt:
		new_list_numbers = []
		new_list_string = []

		for ind, item in enumerate(new_list):

			if item['name'].isdigit():
				new_list_numbers.append(item)
			else:
				new_list_string.append(item)

		new_list_numbers = sorted(new_list_numbers, key=lambda x: int(x['name']))
		new_list_numbers.extend(new_list_string)
		new_list = new_list_numbers

	warning_rows = []

	for i in range(len(new_list) - 1):
		if new_list[i]['name'].isdigit() and new_list[i + 1]['name'].isdigit():
			if int(new_list[i]['name']) != int(new_list[i + 1]['name']) - 1:
				warning_rows.append(i + 2)

	wb = xlwt.Workbook()
	ws = wb.add_sheet('Points')

	for i in range(len(headers_temp)):
		ws.write(0, i, headers_temp[i])

	for i in range(len(new_list)):
		cur_dict = new_list[i]
		cur_dict['ord']  = str(i + 1)

		for k in range(len(list_of_param_temp)):
			ws.write(i + 1, k, cur_dict[list_of_param_temp[k]])

	ws.write(0, 7, "Возможно пропущенные строки")
	ws.write(0, 8, " ".join([str(x) for x in warning_rows]))

	wb.save(address)


def proj4_to_dict(filename):
	'''Converts proj4.txt file to python dictionary where keys is names of zones and values is parametres of proj4'''
	with open(filename, 'r', encoding='utf-8') as file:
		total_list = file.readlines()

	total_dict = {}

	for string in total_list:
		a_list = re.split(r' {2,}', string)
		total_dict[a_list[0]] = a_list[1].rstrip()

	return total_dict


def gpx_to_msk_txt(a_list, address, zone_key, fl_param, xy, add_name=True):
	'''Converts data to txt file with MSK projections'''
	new_dict  = global_dict(a_list, for_exl=False)
	dict_proj4_data = proj4_to_dict('proj4.txt')

	temp_list = []
	not_digit_list = []
	digit_list = []

	if check_extension('.gpx', a_list):
		for key in new_dict:
			temp_dict = new_dict[key]
			name, lat, lon =  temp_dict['name'], temp_dict['coord']['lat'], temp_dict['coord']['lon']
			temp_list.append((name, lat, lon))

	elif check_extension('.kml', a_list):
		for key in new_dict:
			temp_dict = new_dict[key]
			name, lat, lon = temp_dict['name'], temp_dict['lat'], temp_dict['lon']
			temp_list.append((name, lat, lon))

	for ind, tup in enumerate(temp_list):
		if tup[0].isdigit():
			digit_list.append(temp_list[ind])
		else:
			not_digit_list.append(temp_list[ind])

	for ind, tup in enumerate(digit_list):
		current = list(tup)
		current[0] = str(int(current[0]))
		digit_list[ind] = tuple(current)

	digit_list = sorted(digit_list, key=lambda x: int(x[0]))

	digit_list.extend(not_digit_list)

	temp_list = digit_list

	inProj = pyproj.Proj(init='epsg:4326')
	outProj = inProj if zone_key[1:] == 'WGS-84' else pyproj.Proj(dict_proj4_data[zone_key])
	# outProj = pyproj.Proj(dict_proj4_data[zone_key])

	msk_list = []

	for name, lat, lon in temp_list:
		new_lon, new_lat = pyproj.transform(inProj, outProj, lon, lat)
		if xy == 'XY':
			msk_list.append([name, new_lat, new_lon])
		elif xy == 'YX':
			msk_list.append([name, new_lon, new_lat])


	format_to(msk_list, param=fl_param)

	# temp_digit = []
	# temp_non_digit = []

	# for lst in msk_list:
	# 	if lst[0].isdigit():
	# 		temp_digit.append(lst)
	# 	else:
	# 		temp_non_digit.append(lst)


	# temp_digit = sorted(temp_digit, key=lambda x: int(x[0]))

	# temp_digit.extend(temp_non_digit)

	# msk_list = temp_digit

	if not add_name:
		for i in range(len(msk_list)):
			msk_list[i] = msk_list[i][1:]

	with open(address, 'w', encoding='utf-8') as file:
		for tup in msk_list:
			string = ' '.join(tup) + '\n'
			file.write(string)

	return True


def add_n_to_tag(string):
	new = ''
	
	for ind, char in enumerate(string):
		if char == '>' and ind == len(string) - 1:
			new_char = char + '\n'
			new += new_char
		elif char == '>' and string[ind + 1] == '<':
			new_char = char + '\n'
			new += new_char
		else:
			new += char

	return new 	


def format_file(filename):	 

	with open(filename, 'r', encoding='utf-8') as file:
		list_file = file.readlines()

	for ind, string in enumerate(list_file):
		list_file[ind] = add_n_to_tag(string)

	with open(filename, 'w', encoding='utf-8') as file:
		for string in list_file:
			file.write(string)			


def gpx_to_kml(a_list, address):
	list_of_dict = global_dict(a_list, for_exl=False, for_transpotation=True)
	
	for dictionary in list_of_dict:
	
		for i in range(len(dictionary)):
			dictionary[str(i + 1)]['lat'] = dictionary[str(i + 1)]['coord']['lat']
			dictionary[str(i + 1)]['lon'] = dictionary[str(i + 1)]['coord']['lon']
			dictionary[str(i + 1)].pop('coord')


	tree = et.parse('templates/kml_template.kml')
	root = tree.getroot()

	lookAT = root[0].find('LookAt')
	lat = lookAT.find('latitude')
	lon = lookAT.find('longitude')

	lat.text = list_of_dict[0]['1']['lat']
	lon.text = list_of_dict[0]['1']['lon']

	num = 0

	for i in root[0]: num += 1

	for ind, dictionary in enumerate(list_of_dict):
		et.SubElement(root[0], 'Folder')
		folder = root[0][num]
		name = et.SubElement(folder, 'name')
		name.text = a_list[ind].split('/')[-1]

		count = 0

		for point in dictionary:
			count += 1
			
			coord_text = dictionary[point]['lon'] + ',' + dictionary[point]['lat']
			time_text = dictionary[point]['date'] + 'T' + dictionary[point]['time'] + 'Z'
			
			placemark = et.SubElement(folder, 'Placemark')
			name = et.SubElement(folder[count], 'name')
			name.text = dictionary[point]['name']
			description = et.SubElement(folder[count], 'description')
			if 'cmt' in  dictionary[point]:
				description.text = dictionary[point]['cmt']
			styleUrl = et.SubElement(folder[count], 'styleUrl')
			styleUrl.text = '#waypoint'
			point = et.SubElement(folder[count], 'Point')
			coordinates = et.SubElement(point, 'coordinates')
			coordinates.text = coord_text
			end = et.SubElement(folder[count], 'end')
			timeInstant = et.SubElement(end, 'TimeInstant')
			timePosition = et.SubElement(timeInstant, 'timePosition')
			time = et.SubElement(timePosition, 'time')
			time.text = time_text

		num += 1

	tree.write(address, encoding='utf-8')

	with open(address, 'r', encoding='utf-8') as file:
		list_file = file.readlines()
		list_file[0] = '<kml xmlns="http://earth.google.com/kml/2.0">\n' 
		list_file.insert(0, '<?xml version="1.0" encoding="UTF-8"?>\n')

	with open(address, 'w', encoding='utf-8') as file:
		for string in list_file:
			file.write(string)

	format_file(address)


def kml_to_gpx(a_list, address):
	total = []
	for file in a_list:
		list_of_points = kml_to_dict(file)
		total.extend(list_of_points)

	now = datetime.datetime.now()

	time_string = now.strftime("%Y-%m-%dT%H:%M:%SZ")

	tree = et.parse('templates/gpx_template.gpx')
	root = tree.getroot()

	max_lat = total[0]['lat']
	max_lon = total[0]['lon']
	min_lat = total[0]['lat']
	min_lon = total[0]['lon']

	for dictionary in total[1:]:
		lat = dictionary['lat']
		lon = dictionary['lon']
		max_lat = lat if float(lat) > float(max_lat) else max_lat
		max_lon = lon if float(lon) > float(max_lon) else max_lon
		min_lat = lat if float(lat) < float(min_lat) else min_lat
		min_lon = lon if float(lon) < float(min_lon) else min_lon

	creation_time_tag  = root[0].find('time')
	creation_time_tag.text = time_string

	bounds = root[0].find('bounds')

	bounds.set('minlat', min_lat)
	bounds.set('minlon', min_lon)
	bounds.set('maxlat', max_lat)
	bounds.set('maxlon', max_lon)

	for point in total:
		wpt = et.SubElement(root, 'wpt')
		wpt.set('lat', point['lat'])
		wpt.set('lon', point['lon'])
		ele = et.SubElement(wpt, 'ele')
		ele.text = point['ele']
		time = et.SubElement(wpt, 'time')
		time.text = point['time']
		name = et.SubElement(wpt, 'name')
		name.text = point['name']
		cmt = et.SubElement(wpt, 'cmt')
		cmt.text = point['cmt']
 
	tree.write(address, encoding='utf-8')

	with open(address, 'r', encoding='utf-8') as file:
		list_file = file.readlines()
		list_file[0] = '<gpx xmlns="http://www.topografix.com/GPX/1/1" creator="OziExplorer Version 3955k - http://www.oziexplorer.com" version="1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd">\n' 
		list_file.insert(0, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')

	with open(address, 'w', encoding='utf-8') as file:
		for string in list_file:
			file.write(string)

	format_file(address)


def total_table(a_list, address, error):
	plus_minus = '\u00B1'
	grad = '\u00BA'

	ultimately = []

	more_than_mid_search = [] # rows where search values more than doubled average of search 

	total = {"ord":[], "lat":[], "lon":[], "mad":[], "ind":[], "comm":[], "err":[], "mad_err":[], 'search': []}

	list_of_total = ['ord', 'lat', 'lon', 'mad_err', 'ind', 'comm']

	file1, file2 = a_list[0], a_list[1]

	book1 = xlrd.open_workbook(file1)
	book2 = xlrd.open_workbook(file2)

	sheet1 = book1.sheet_by_index(0)
	sheet2 = book2.sheet_by_index(0)

	cell_1_2 = sheet1.cell(1, 2).value

	if not grad in str(cell_1_2):
		sheet1, sheet2 = sheet2, sheet1

	if sheet1.nrows != sheet2.nrows:
		return "Разное количество строк у файлов"		

	for i in range(1, sheet1.nrows):
		total["ord"].append(sheet1.cell(i, 0).value)
		total["lat"].append(sheet1.cell(i, 2).value)
		total["lon"].append(sheet1.cell(i, 3).value)
		total["mad"].append(sheet2.cell(i, 2).value)
		total["ind"].append(str(sheet2.cell(i, 3).value))
		total["comm"].append(sheet2.cell(i, 4).value)
		total['search'].append(sheet2.cell(i, 1).value)

	mid = lambda x: float(Decimal(str(sum(x) / len(x))).quantize(Decimal('1.00'), rounding = ROUND_HALF_UP))

	# mid = lambda x: roundRus(sum(x) / len(x)) if x else 0

	response = checkTypeInList(total["search"])

	if response != "success":
		return "Неверное значение в файле журнала в графе ПОИСК на строке " + str(response + 1)

	min_search = min(total['search'])
	max_search = max(total['search'])
	mid_search = trueRound(sum(total['search']) / len(total['search']))

	doubled_search = mid_search * 2

	response = checkTypeInList(total["mad"])

	if response != "success":
		return "Неверное значение в файле журнала в графе МАД на строке " + str(response + 1)

	min_mad = min(total["mad"])
	max_mad  =max(total["mad"])
	mid_mad = mid(total["mad"])

	ultimately.append(("", "ПОИСК", "МАД"))
	ultimately.append(("МИН", min_search, min_mad))
	ultimately.append(("МАКС", max_search, max_mad))
	ultimately.append(("СРЕД", mid_search, mid_mad))

	for i in range(len(total["ind"])):
		try:
			if total["ind"][i][-1] == '0':
				total["ind"][i] = total["ind"][i][0]
			else:
				total["ind"][i] = total["ind"][i].replace('.', ',')
		except:
			return str("Пустой или неверный индекс в файле журнала в строке " + str(i + 1))

	for i in range(len(total["mad"])):
		total["err"].append((str(round(total["mad"][i] * error/100, 2))).replace('.', ','))
		total["mad"][i] = str(total["mad"][i]).replace('.', ',')
		total['mad_err'].append(total['mad'][i] + " " + plus_minus + " " + total['err'][i])

	for ind, val in enumerate(total['search']):
		if val >= doubled_search:
			more_than_mid_search.append(ind + 1)

	wb = xlwt.Workbook()
	ws = wb.add_sheet('Table')

	header = ["N", "Северная", "Восточная", "МАД", "Индекс", "Комментарий"]

	for i in range(len(header)):
		ws.write(0, i, header[i])

	for i in range(len(list_of_total)):
		col = total[list_of_total[i]]
		for j in range(len(col)):
			ws.write(j + 1, i, col[j])

	temp_col = 8

	for i in range(len(ultimately)):
		row = ultimately[i]
		for j in range(len(row)):
			ws.write(i, temp_col + j, row[j])

	ws.write(0, 12, "Точки значение, которых превышает удвоенное среднее значение по 'ПОИСКУ': ")
	
	if more_than_mid_search:
		points = " ".join([str(x) for x in more_than_mid_search])
	else: 
		points = "0"
		
	ws.write(0, 13, points)

	wb.save(address)

	return "success"

def transition_to_excel(list_data, address, zone_key):

	symb = ['\u00BA', "\'",  '\"']
	# symb = ['°', "\'",  '\"']
	dict_proj4_data = proj4_to_dict('proj4.txt')
	headers = ['Номер', 'Имя', 'Северная широта', 'Восточная долгота']

	outProj = pyproj.Proj(init='epsg:4326')
	inProj = outProj if zone_key[1:] == 'WGS-84' else pyproj.Proj(dict_proj4_data[zone_key])

	total_list = []
	
	for name, lat, lon in list_data:
		new_lon, new_lat = pyproj.transform(inProj, outProj, lon, lat)
		total_list.append([name, new_lat, new_lon])
			
	format_to(total_list, param=6)

	for i in range(len(total_list)):
		current = total_list[i]
		lat_lst = coordList(current[1])
		lon_lst = coordList(current[2])

		format_lat_lst = ''
		format_lon_lst = ''
		
		for j in range(len(lat_lst)):
			format_lat_lst += str(lat_lst[j])
			format_lat_lst += symb[j]
			format_lon_lst += str(lon_lst[j])
			format_lon_lst +=  symb[j]
		
		current[1] = format_lat_lst
		current[2] = format_lon_lst
		total_list[i] = current 

	wb = xlwt.Workbook()
	ws = wb.add_sheet('Points')

	for i in range(len(headers)):
		ws.write(0, i, headers[i])

	for j in range(len(total_list)):
		ws.write(j + 1, 0, j + 1)
		current_el = total_list[j]
		for k in range(len(current_el)):
			ws.write(j + 1, k + 1, current_el[k])

	wb.save(address)

def txt_msk_to_exl(a_list, address, zone_key):
		"""converts txt file with MSK coordinates to excel file with WGS-84 coordinates"""

		total_list = []
		total_list_with_tuples = []
		
		for add in a_list:
			with open(add, 'r') as file:
				lst = file.readlines()
				total_list.extend(lst)

		for string in total_list:
			tup = tuple(string.rstrip().split(" "))
			total_list_with_tuples.append(tup)

		# try:
		# 	transition_to_excel(total_list_with_tuples, address, zone_key)
		# except:
		# 	return -1
		# else:
		# 	return 1

		transition_to_excel(total_list_with_tuples, address, zone_key)

def exl_to_lst(a_list):

	file = a_list[0]
	book = xlrd.open_workbook(file)
	sheet = book.sheet_by_index(0)

	total = []

	temp_digit = []
	temp_non_digit = []

	strMad = "НОМЕР_ТОЧКИ_МАД"
	strNum = "НОМЕР"
	strLat = "Положение Y"
	strLon = "Положение X"

	numOfNameString = None
	flag = False 

	for j in range(sheet.ncols):
		currentString = sheet.cell(0, j).value
		
		if currentString == strMad:
			numOfNameString = j
			flag = True
		elif currentString == strLat:
			numOfLatString = j 
		elif currentString == strLon:
			numOfLonString = j
		else:
			pass 

	if not numOfNameString:
		for j in range(sheet.ncols):
			currentString = sheet.cell(0, j).value

			if currentString == strNum:
				numOfNameString = j
				break 

	for i in range(1, sheet.nrows):
		name = sheet.cell(i, numOfNameString).value
		
		if flag:
			name = name.split(";")[-1]

		lat = sheet.cell(i, numOfLatString).value 
		lon = sheet.cell(i, numOfLonString).value
		
		total.append((name, lat, lon))

	for tup in total:
		if tup[0].isdigit():
			temp_digit.append(tup)
		else:
			temp_non_digit.append(tup)

	temp_digit = sorted(temp_digit, key=lambda x: int(x[0]))

	temp_digit.extend(temp_non_digit)

	total = temp_digit

	return total 

def autocad_msk_to_exl(a_list, address, zone_key, fl_param):
	
	total = exl_to_lst(a_list)

	try:
		transition_to_excel(total, address, zone_key)
	except:
		return -1
	else:
		return 1

	# transition_to_excel(total, address, zone_key)

def xls_to_kml(a_list, address, zone_key):
	total = exl_to_lst(a_list)
	dict_proj4_data = proj4_to_dict('proj4.txt')

	outProj = pyproj.Proj(init='epsg:4326')
	inProj = outProj if zone_key[1:] == 'WGS-84' else pyproj.Proj(dict_proj4_data[zone_key])

	total_list = []
	
	for name, lat, lon in total:
		new_lon, new_lat = pyproj.transform(inProj, outProj, lon, lat)
		total_list.append([name, new_lat, new_lon])
			
	format_to(total_list, param=6)



	tree = et.parse('templates/kml_template.kml')
	root = tree.getroot()

	lookAT = root[0].find('LookAt')
	lat = lookAT.find('latitude')
	lon = lookAT.find('longitude')

	lat.text = total_list[0][1]
	lon.text = total_list[0][2]

	document = root[0]

	et.SubElement(document, "Folder")
	folder = document.find("Folder")

	et.SubElement(folder, "name")
	name = folder.find("name")
	name.text = "Waypoints"

	now = datetime.datetime.now()
	current_time = now.strftime("%Y-%m-%dT%H:%M:%SZ")

	for wpoint in total_list:
		
		placemark = et.SubElement(folder, "Placemark")
		name_of_placemark = et.SubElement(placemark, "name")
		name_of_placemark.text = wpoint[0]
		et.SubElement(placemark, "description")
		styleUrl = et.SubElement(placemark, "styleUrl")
		styleUrl.text = "#waypoint"
		point = et.SubElement(placemark, "Point")
		coordinates = et.SubElement(point, "coordinates")
		coordinates.text = wpoint[2] + ',' + wpoint[1]
		end = et.SubElement(placemark, "end")
		timeInstant = et.SubElement(end, "TimeInstant")
		timePosition = et.SubElement(timeInstant, "timePosition")
		time = et.SubElement(timePosition, "time")
		time.text = current_time

	tree.write(address, encoding='utf-8')

	with open(address, 'r') as file:
		list_file = file.readlines()
		list_file[0] = '<kml xmlns="http://earth.google.com/kml/2.0">\n' 
		list_file.insert(0, '<?xml version="1.0" encoding="UTF-8"?>\n')

	with open(address, 'w') as file:
		for string in list_file:
			file.write(string)

	format_file(address)

def xls_to_gpx(a_list, address, zone_key):
	total = exl_to_lst(a_list)
	dict_proj4_data = proj4_to_dict('proj4.txt')

	outProj = pyproj.Proj(init='epsg:4326')
	inProj = outProj if zone_key[1:] == 'WGS-84' else pyproj.Proj(dict_proj4_data[zone_key])

	total_list = []
	
	for name, lat, lon in total:
		new_lon, new_lat = pyproj.transform(inProj, outProj, lon, lat)
		total_list.append([name, new_lat, new_lon])
			
	format_to(total_list, param=6)

	now = datetime.datetime.now()

	time_string = now.strftime("%Y-%m-%dT%H:%M:%SZ")

	tree = et.parse('templates/gpx_template.gpx')
	root = tree.getroot()

	max_lat = total_list[0][1]
	max_lon = total_list[0][2]
	min_lat = total_list[0][1]
	min_lon = total_list[0][2]

	for lst in total_list[1:]:
		lat = lst[1]
		lon = lst[2]
		max_lat = lat if float(lat) > float(max_lat) else max_lat
		max_lon = lon if float(lon) > float(max_lon) else max_lon
		min_lat = lat if float(lat) < float(min_lat) else min_lat
		min_lon = lon if float(lon) < float(min_lon) else min_lon

	creation_time_tag  = root[0].find('time')
	creation_time_tag.text = time_string

	bounds = root[0].find('bounds')

	bounds.set('minlat', min_lat)
	bounds.set('minlon', min_lon)
	bounds.set('maxlat', max_lat)
	bounds.set('maxlon', max_lon)

	for point in total_list:
		wpt = et.SubElement(root, 'wpt')
		wpt.set('lat', point[1])
		wpt.set('lon', point[2])
		ele = et.SubElement(wpt, 'ele')
		ele.text = "0.0"
		time = et.SubElement(wpt, 'time')
		time.text = time_string
		name = et.SubElement(wpt, 'name')
		name.text = point[0]
		cmt = et.SubElement(wpt, 'cmt')
		cmt.text = '0'
 
	tree.write(address, encoding='utf-8')

	with open(address, 'r', encoding='utf-8') as file:
		list_file = file.readlines()
		list_file[0] = '<gpx xmlns="http://www.topografix.com/GPX/1/1" creator="OziExplorer Version 3955k - http://www.oziexplorer.com" version="1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd">\n' 
		list_file.insert(0, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')

	with open(address, 'w', encoding='utf-8') as file:
		for string in list_file:
			file.write(string)

	format_file(address)

def checkTypeInList(lst):
	for ind, obj in enumerate(lst):
		if not isinstance(obj, float):
			return ind
	return "success"

if __name__ == '__main__':
	n = "°"
	# numbers = [5, 17.1, 0.05445, 0.544848, 0.054848, 0.064545]
	

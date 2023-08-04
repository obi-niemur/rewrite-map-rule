import xlrd

def parse_input(in_str):
	replace_dict = {
            #Replace these with domains
		'https://www.domain-name.com/': '',
                'www.domain-name/': '',
                'domain-name-without-www.com/':'',
		'%20': ' ',
		'&': '&amp;'
	}
	in_str = in_str.strip()
	for key, val in replace_dict.items():
		in_str = in_str.replace(key, val)
	return in_str

def generate_map():
	fout = open('C:/rewrite/config/web.config','w')
	wb = xlrd.open_workbook('C:/Users/niemu/Downloads/config.xls')
	sheet = wb.sheet_by_name('Sheet1')

	fout.write('<rewriteMaps>\n<rewriteMap name="Redirects">\n')

	

	key_list = []

	for rownum in range(sheet.nrows):
		keyval = sheet.row_values(rownum,0,3)

		rowkey = parse_input(keyval[0])

		# skip duplicates
		if rowkey in key_list:
			
			continue

		key_list.append(rowkey)

		mapval = "<add key=\"/"
		mapval += parse_input(keyval[0])
		mapval += "\" value=\""
		mapval += parse_input(keyval[1])
		mapval += "\" />\n"

		# one off for lakeforest
		# mapval += "<add key=\""
		# mapval += parse_input(keyval[0][0:-1])
		# mapval += "\" value=\""
		# mapval += parse_input(keyval[1])
		# mapval += "\" />\n"
		fout.write(mapval)

	fout.write('</rewriteMap>\n</rewriteMaps>')
	fout.close()

generate_map()

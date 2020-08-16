import re
import os
import sys
import string

import QBibSearch as qbs
import urllib.request as urllib2
import sqlite3 as db

hebrew_db_name = 'wlc.db'
num_ot = 39
nau_map_file = 'nau-vmf.txt'
py_map_file = 'wttmap.py'

book_to_key = {
'Gen': '01',
'Exo': '02',
'Lev': '03',
'Num': '04',
'Deu': '05',
'Jos': '06',
'Jdg': '07',
'Rut': '08',
'1Sa': '09',
'2Sa': '10',
'1Ki': '11',
'2Ki': '12',
'1Ch': '13',
'2Ch': '14',
'Ezr': '15',
'Neh': '16',
'Est': '17',
'Job': '18',
'Psa': '19',
'Pro': '20',
'Ecc': '21',
'Sol': '22',
'Isa': '23',
'Jer': '24',
'Lam': '25',
'Eze': '26',
'Dan': '27',
'Hos': '28',
'Joe': '29',
'Amo': '30',
'Oba': '31',
'Jon': '32',
'Mic': '33',
'Nah': '34',
'Hab': '35',
'Zep': '36',
'Hag': '37',
'Zec': '38',
'Mal': '39',
'Mat': '40',
'Mar': '41',
'Luk': '42',
'Joh': '43',
'Act': '44',
'Rom': '45',
'1Co': '46',
'2Co': '47',
'Gal': '48',
'Eph': '49',
'Phi': '50',
'Col': '51',
'1Th': '52',
'2Th': '53',
'1Ti': '54',
'2Ti': '55',
'Tit': '56',
'Phl': '57',
'Heb': '58',
'Jam': '59',
'1Pet':'60',
'2Pet':'61',
'1Joh':'62',
'2Joh':'63',
'3Joh':'64',
'Jud': '65',
'Rev': '66'
}

book_table_keys = list(qbs.book_table.keys())
book_table_keys.sort()
txt_list = list()
url_list = list()
book_table = dict()

class VerseValue:
	def __init__(self, verse):
		self.nverse = verse
		
class BookInfo:
	def __init__(self):
		self.chap = dict()

def create_txt_list():
	for i in range(num_ot):
		book_obj = qbs.book_table[book_table_keys[i]]
			
		if re.search(r'[Ss]ong', book_obj[3]):
			book_name = 'Song of Songs'
		else:
			book_name = book_obj[3]
			
		url_name = book_name
		item_list = re.finditer(r'\s', url_name)
	
		if item_list:
			p1 = 0
			temp = ''
			for item in item_list:
				temp += url_name[p1:item.start()]
				temp += '%20'
				p1 = item.end()
			temp += url_name[p1:]
			url_name = temp
			
		#url_list.append('http://tanach.us/TextServer?{}*&content=Accents'.format(url_name))
		url_list.append('http://tanach.us/TextServer?{}*&content=Vowels'.format(url_name))
		txt_list.append(os.path.join(os.getcwd(), 'bible', 'OT', '{}.txt'.format(book_name)))
		
def get_wlc_text():

	for i in range(num_ot):
		book_obj = qbs.book_table[book_table_keys[i]]
		print('Connect to {}'.format(url_list[i]))
		
		page = urllib2.urlopen(url_list[i])
		page_content = page.read().decode('utf-8')
		
		with open(txt_list[i], mode='w', encoding='utf-8') as file:
			file.write(page_content)
			file.close()

def create_wlc_db():
	if os.path.isfile(hebrew_db_name): os.remove(hebrew_db_name)
	db_con = db.connect(hebrew_db_name)
	db_cur = db_con.cursor()
	db_cur.execute("CREATE TABLE HebBible(book INT, chap INT, verse INT, vtext TEXT)")
			
	for i in range(num_ot):
		book_obj = qbs.book_table[book_table_keys[i]]
		
		try:
			print("... Open {}".format(txt_list[i]))
			file = open(txt_list[i], encoding='utf-8')
		except IOError as err:
			errno, strerror = err.args
			print("... I/O error({0}): {1}".format(errno, strerror))
			print("... Can't open {}".format(fname))
			break
		
		for line in file:
			skip_comment = re.search(r'xxxx', line)
			if skip_comment: continue
			chap = re.search(r'(\d+).*?(\d+)', line)
			
			if chap:
				#print(chap.group(1), chap.group(2), chap.end(2))
				# skip spaces
				pos = chap.end(2)
				while line[pos].isspace(): pos = pos+1
				db_cur.execute("INSERT INTO HebBible VALUES('{0}', '{1}', '{2}', '{3}')"\
                .format(i+1,chap.group(2), chap.group(1), line[pos:len(line)-2]));
	
	db_con.commit()
	db_con.close()

def createStdVerseTable():
	
	out = open('vlist.txt', 'wt')
	file = open(qbs.qbs_datafile)
	temp = file.readline()
	bdf_path = temp.replace('\n', "")
	file.close()
	
	for i in range(0, 66):
		book_table[book_table_keys[i]] = BookInfo()
	
	# open NASB 
	for i in range(qbs.max_bdf_number):
		file = open(os.path.join(bdf_path, '{}{}.bdf'.format(qbs.english_bible_prefix[0], i+1)))
		for line in file:
			item = re.search(r'^(\d+).*?(\d+):(\d+)', line)
			if item:
				book = item.group(1)
				chap = item.group(2)
				vers = item.group(3)
				book_table[book].chap[chap] = int(vers)
	
	
	for i in range(qbs.number_of_bible_books):
		book = book_table_keys[i]
		#sorted_chap = sorted(book_table[book].chap.items(), key = lambda x:x[0])
		#print(book,  sorted_chap)
		#print(book,  book_table[book].chap)
		
		divunit = 5
		nchap = len(book_table[book].chap)
		tenth = int(nchap / divunit)
		#leftover = nchap - tenth*10
		leftover = nchap % divunit
		#print(nchap, tenth, left)
		
		last_key = 0
		book_obj = qbs.book_table[book_table_keys[i]]

		print('%s: %s' % (book_obj[3], book_obj[2]))
		out.write('%s: %s\n' % (book_obj[3], book_obj[2]))

		for i in range(tenth):
			chap = list()
			for j in range(divunit):
				key = i*divunit+j+1
				print('%3d:%3d ' % (key, book_table[book].chap[str(key)]), end=' ')
				out.write('%3d:%3d ' % (key, book_table[book].chap[str(key)]))
				last_key = key
			print('')
			out.write('\n')
		
		if leftover:
			chap = list()
			key = last_key
			for j in range(leftover):
				key = key +1
				print('%3d:%3d ' % (key, book_table[book].chap[str(key)]), end=' ')
				out.write('%3d:%3d ' % (key, book_table[book].chap[str(key)]))
			print('')
			out.write('\n')
		out.write('\n')
	out.close()
		
def create_otmap_table():
	file1 = open(nau_map_file)
	file2 = open(py_map_file, 'w')
	
	# key:chap:vers
	file2.write('wtt_table = {\n')
	
	for line in file1:
		list = line.split(' ')
		key = book_to_key[list[0]]
		
		item1 = list[1].split(':')
		item2 = list[3].split(':')

		chap_src  = item1[0]
		chap_dest = item2[0]
		
		if item1[1].find('-') != -1:
			vers_src  = item1[1].split('-')
			vers_dest = item2[1].split('-')
			
			vers1 = int(vers_src[0])
			vers2 = int(vers_src[1])
			vers3 = int(vers_dest[0])

			map_vers1 = vers1			
			map_vers2 = vers3

			for i in range(vers1, vers2+1):
				file2.write('\'%s:%s:%s\' : [%s, %d],\n' %\
				(key, chap_src, map_vers1, chap_dest, map_vers2))
				map_vers1 = map_vers1 + 1
				map_vers2 = map_vers2 + 1
		else:
			vers_src  = item1[1]
			vers_dest = item2[1]
			map_vers = vers_dest
			file2.write('\'%s:%s:%s\' : [%s, %s],\n' %\
			(key, chap_src, vers_src, chap_dest, map_vers))
	file2.write('}')
			
def main():
	createStdVerseTable()
	
if __name__ == '__main__':
	main()


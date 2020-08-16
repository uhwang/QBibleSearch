import re
import os
import sys
import string

import QBibSearch as qbs
import urllib.request as urllib2
import sqlite3 as db

num_ot = 39
num_nt = 27
greek_db_name = 'sblgnt.db'
greek_bible_db_table_name = 'GrkBible'

book_table_keys = list(qbs.book_table.keys())
book_table_keys.sort()
txt_list = list()
url_list = list()

def create_txt_list():
	directory = os.path.join(os.getcwd(), 'bible', 'SBLGNTtxt')
	for root, directories, files in os.walk(directory):
		for filename in files:
			txt_list.append(os.path.join(directory, filename))
		
def create_gnt_db():
	if os.path.isfile(greek_db_name): os.remove(greek_db_name)
	
	nt_index = num_ot
	db_con = db.connect(greek_db_name)
	db_cur = db_con.cursor()
	db_cur.execute("CREATE TABLE GrkBible(book INT, chap INT, verse INT, vtext TEXT)")
			
	for i in range(num_nt):
		nt_index = nt_index+1
		book_obj = qbs.book_table[book_table_keys[num_ot+i]]
		
		try:
			print("... Open {}".format(txt_list[i]))
			file = open(txt_list[i], mode = 'r', encoding='utf-8')
		except IOError as err:
			errno, strerror = err.args
			print("... I/O error({0}): {1}".format(errno, strerror))
			print("... Can't open {}".format(fname))
			break
		
		for line in file:
			chap = re.search(r'(\d+):(\d+)', line)
			if chap:
				pos = chap.end(2)
				print(nt_index,chap.group(1), chap.group(2))
				
				while line[pos].isspace(): pos = pos+1
				try:
					db_cur.execute("INSERT INTO GrkBible VALUES('{0}', '{1}', '{2}', '{3}')"\
                    .format(nt_index,chap.group(1), chap.group(2), line[pos:]));
				except db.OperationalError as err:
					#err1, err2, err3 = err.args
					#print("... SQlite3 error --> ({0}): {1}".format(errno, strerror))
					print("{} {}:{}".format(book_obj[3], chap.group(1), chap.group(2)))
					print(line.encode('utf-8'))
					break
	
	db_con.commit()
	db_con.close()

def main():
	out = open('nt-out.txt', mode='w', encoding='utf-8')
	db_con = db.connect(greek_db_name)
	db_cur = db_con.cursor()
	for vtext in db_cur.execute("SELECT vtext from GrkBible where book = 40 and chap = 2 and verse = 1"):
		out.write('John 1:1 {}'.format(vtext[0]))
	db_con.close()
	out.close()
	
if __name__ == '__main__':
	create_txt_list()
	create_gnt_db()
	main()


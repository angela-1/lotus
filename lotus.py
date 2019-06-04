#coding:utf-8 
import os
import shutil

def docx2pdf(filepath):
	cmd = ['morning-glory.exe -f pdf ' + '"' + filepath + '"']
	os.system(' '.join(cmd))
	filename, ext = os.path.splitext(filepath)
	return f'{filename}.pdf'
	

def filter_docx(filepath):
	filename, ext = os.path.splitext(filepath)
	return ext == '.docx' or ext == '.doc'

def filter_pdf(filepath):
	filename, ext = os.path.splitext(filepath)
	return ext == '.pdf'


def get_folder_name(filepath):
	dirname = os.path.dirname(filepath)
	folder = os.path.split(dirname)
	return folder[1]


def modify_bookmark(filepath):
	filename, ext = os.path.splitext(filepath)

	title = os.path.split(filename)[1]

	bookmark_file = '_bookmark.txt'

	cmd = ['pdftk\\pdftk.exe', filepath, 'dump_data', 'output', bookmark_file]
	os.system(' '.join(cmd))

	f = open(bookmark_file, 'r')
	lines = f.readlines()
	f.close()
	new_lines = []
	for l in lines:
		if l.strip()[0:14] == 'NumberOfPages:':
			new_lines.append(l)
			level1 = f'''BookmarkBegin
BookmarkTitle: {str2ascii(title)}
BookmarkLevel: 1
BookmarkPageNumber: 1
'''			
			new_lines.append(level1)
			continue

		if l.strip() == 'BookmarkLevel: 1':
			ll = 'BookmarkLevel: 2\n'
			new_lines.append(ll)
		else:
			new_lines.append(l)

	level1_bookmark_file = '_level1_bookmark_file.txt'
	ff = open(level1_bookmark_file, 'w')
	ff.write(''.join(new_lines))
	ff.close()

	new_file = os.path.split(filename)[1] + '_bookmarked' + ext
	target = os.path.join('dist_bookmarked', new_file)

	if not os.path.exists('dist_bookmarked'):
		os.mkdir('dist_bookmarked')

	cmd = ['pdftk\\pdftk.exe', '"' + filepath + '"', 
		'update_info', level1_bookmark_file, 'output', '"' +target + '"']
	os.system(' '.join(cmd))

	return target

def odd2even(file):
	cmd = ['odd2even\\odd2even.exe', f'"{file}"']
	os.system(' '.join(cmd))
	fn, ext = os.path.splitext(file)
	target = fn + '_even' + ext
	return target

def merge_files(files_list):
	folder = get_folder_name(files_list[0])

	target = os.path.join('dist', f'{folder}.pdf')

	files = '" "'.join(files_list)
	cmd = ['pdftk\\pdftk.exe', '"' + files + '"', 'cat', 'output', target]
	os.system(' '.join(cmd))

	return target


def write_bookmark_file(filepath, title):
	bookmark_template = '''BookmarkBegin
BookmarkTitle: {title}
BookmarkLevel: 1
BookmarkPageNumber: 1
'''
	bookmark = bookmark_template.replace('{title}', str2ascii(title))
	f=open(filepath, 'w')
	f.write(bookmark)
	f.close()



def add_bookmark(filepath):
	filename, ext = os.path.splitext(filepath)
	bookmark_file = '_bookmark.txt'
	title = os.path.splitext(os.path.basename(filepath))
	write_bookmark_file(bookmark_file, title[0].replace('_even', ''))
	target = filename + '_bookmarked' + ext
	cmd = ['pdftk', '"' + filepath + '"', 
		'update_info', bookmark_file, 'output', '"' + target + '"']
	os.system(' '.join(cmd))
	return target


def str2ascii(str):
	return ''.join(list(map(char2ascii, str)))

def char2ascii(char):
	code = ord(char)
	return f'&#{code};'


def clean(files_list):
	for f in files_list:
		os.remove(f)


def do(folder):
	# folder = input('请输入目录：')
	print(f'文件夹：{folder}')
	merge = input('是否仅合并：（y/N）')
	justMerge = False if merge.lower() == 'n' or merge == '' else True


	files_list = map(lambda f: os.path.join(folder, f), os.listdir(folder))
	docx_files = list(filter(filter_docx, files_list))
	pdf_files = list(map(docx2pdf, docx_files))

	pdf_files_list = list(filter(filter_pdf, 
		map(lambda f: os.path.join(folder, f), os.listdir(folder))))

	even_files_list = list(map(odd2even, pdf_files_list))

	if not os.path.exists('dist'):
		os.mkdir('dist')


	if justMerge:
		target = merge_files(even_files_list)
		clean(even_files_list)
	else:
		bookmarked_files_list = list(map(add_bookmark, even_files_list))
		target = merge_files(bookmarked_files_list)
		modify_bookmark(target)
		clean(even_files_list)
		clean(bookmarked_files_list)


def read_folder_list():
    file = open("foo.txt", 'r', encoding='UTF-8')
    list = map(lambda x: x.strip(), file.readlines())
    file.close()
    return list

if __name__ == '__main__':
	print('...')
	
	folder_list = read_folder_list()

	for x in folder_list:
		do(x)

	print('Done')

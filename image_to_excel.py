"""img_table.py: Converts provided table image into excel sheet."""

from PIL import Image, ImageOps
from pytesseract import image_to_string
import itertools
import re
import xlsxwriter
from string import ascii_uppercase as ALP
import zipfile as zp
import time
import os
import logging
from pathlib import Path
import shutil

BASIC = {
	'bold': 1,
	'border': 1,
	'align': 'center',
	'valign': 'vcenter',
}


def make_zip(abspath,ext):

	global OUTPUT_IMG_ZIP_NAME
	global INPUT_IMAGE_DIR

	output_filename = OUTPUT_IMG_ZIP_NAME

	zip_file_path = os.path.join(OUTPUT_IMG_ZIP,output_filename)
	zfile = zp.ZipFile(zip_file_path+'.zip','a')
	arcname = os.path.basename(abspath).split('.')[0]+time.strftime("_%m_%d_%y__%H_%M_%S")+'.'+ext
	zfile.write(abspath,arcname)
	print(f'--> Removing {os.path.basename(abspath)}')
	os.remove(abspath)
	zfile.close()

	print(f'--> ZIP archive created at {OUTPUT_IMG_ZIP}')


def convert_to_b_w(img):
	'''
		Helper function for converting images to black and white with threshold to sharpen.
		Used when pytesseract fails to recognize minute details in image.
	'''
	try:
		thresh = 200

		def fn(x): return 255 if x > thresh else 0

		return img.convert('L').point(fn, mode='1')
	except Exception as e:
		print("Couldn't convert image into black and white.")


def create_worksheet(name) -> tuple:
	'''
			Desc: Create a xlsxwriter worksheet and workbook
			Return: Tuple of workbook and worksheet and validated boolean
	'''

	# creating a excel workbook(document)
	workbook = xlsxwriter.Workbook(name)
	# creating a sheet in excel workbook
	worksheet = workbook.add_worksheet()


	# writing headers in created worksheet
	hd_format = BASIC.copy()
	hd_format.update({'bg_color':'#488A99','font_color':'white'})
	header_format = workbook.add_format(hd_format)
	del hd_format

	title_border = (0, 0, 2199, 147)
	title_img = convert_to_b_w(im.crop(title_border))
	title = image_to_string(title_img).strip().strip()
	# print(title)
	worksheet.merge_range(0, 0, 0, 7, title, workbook.add_format(BASIC))
	header_border = (49, 180, 2150, 290)

	# height of header
	h = 110

	# height of nested header (e.g. in witholding and deductions)
	h1 = 44

	# width of each column
	w = [321, 557, 485, 485, 254]

	# starting co-ordinates of the header row
	start_x = 49
	start_y = 180

	# column index in headers
	ci = 1
	validated = False
	for i in range(5):

		if i in [1, 2, 3]:
			# print('UPP',(start_x,start_y,start_x+w[i],start_y+h1))
			hcell_img1 = im.crop((start_x, start_y, start_x + w[i], start_y + h1))
			data1 = image_to_string(hcell_img1).strip()
			if data1 == '':
				hcell_img1 = convert_to_b_w(hcell_img1)
				data1 = image_to_string(hcell_img1).strip()
			# print(repr(data1))
			
			# print('DOWN',(start_x,start_y+h1,start_x+w[i],start_y+h1))
			hcell_img2 = im.crop((start_x, start_y + h1, start_x + w[i], start_y + h))
			data2 = image_to_string(hcell_img2).strip().split()
			if 'EARNINGS' in data1:
				data1='EARNINGS'
			worksheet.write(grow-1, ci, data1+' '+data2[0], header_format)
			worksheet.write(grow-1, ci+1, data1+' '+data2[1], header_format)
			# print(repr(data2))
			ci += 2

		else:
			hcell_img = im.crop(
				(start_x, start_y, start_x + w[i], start_y + h))
			data = image_to_string(hcell_img).strip().replace('\n\n', '\n')
			if i == 0:
				if 'EMPLOYEE NAME' in data:
					validated = True
					data = 'EMPLOYEE ID'
				worksheet.write(grow-1,0,data,header_format)
			elif i == 4:
				worksheet.write(grow-1,7,data,header_format)
			# print(data)

		start_x += w[i]

	worksheet.set_row(0, 40)
	worksheet.set_row(grow - 2, 30)
	worksheet.set_row(grow - 1, 20)
	
	worksheet.set_column('A:A', 20)
	worksheet.set_column('B:C', 20)
	worksheet.set_column('D:E', 30)
	worksheet.set_column('F:G', 22)
	worksheet.set_column('H:H', 15)

	return workbook, worksheet, validated


def _parse_ssn_id(data):

	if 'ID:' in data:
		return data.strip()


def _parse_common(data):


	nest_li = [i.rsplit(' ', 1) for i in re.split('\n+', data) if i is not '']
	for li in nest_li:
		li[-1] = float(li[-1].replace(',', '_'))

	return nest_li


def _parse_netpay(data):
	return list(itertools.chain(*_parse_common(data)))


def write_to_excel(data, col):


	global aex
	if len(aex) <= 5:
		aex.append(data)
	else:
		if len(aex) == 5:
			# print('FULL')
			pass
		else:
			print('Trying to overwrite aex..')
			logging.error('Trying to overwrite aex..')

	if len(aex) == 5:

		global grow
		global record
		data_format = workbook.add_format({'align': 'center','valign': 'vcenter','num_format': '0.00'})
		if isinstance(aex[0], str):
			print(f'Processing record {record+1}..')
			m = max(len(aex[1]), len(aex[2]), len(aex[3]))
			if 'ID:' in aex[0]:
				aex[0]=aex[0][3:]
			for i in range(grow,grow+m):
				worksheet.write(i,0,aex[0], workbook.add_format(BASIC))

			col = 1
			tt = ['Total Earnings','Total Witholdings','Total Deductions']
			total_row = []
			total_row_format = BASIC.copy()
			total_row_format.update({'fg_color':'#DADADA','num_format': '0.00'})
			for idx,ind in enumerate(aex[1:4]):
				row = grow
				sm = 0
				for i, j in ind:
					if idx==0:
						if 'TOTAL EARNINGS' in i:
							total_row.append(tt[idx])
							total_row.append(j)
					else:
						sm+=j

					if not 'TOTAL EARNINGS' in i:
						worksheet.write(row,col,i,data_format)
						worksheet.write(row,col+1,j,data_format)
					row += 1
				if idx!=0:
					if sm==0:
						sm=''
					total_row.append(tt[idx])
					total_row.append(sm)
				col += 2

			# appending net_pay
			total_row.append(aex[-1][0])
			
			# writing the extra row for total
			worksheet.write_row(grow+m,1,total_row,workbook.add_format(total_row_format))
			
			del total_row_format

		aex = []
		grow += m+1
		record += 1


def parse_and_write(data: str, col):
	# a dictionary to call parse function according to the cell index
	dic = {
		0: _parse_ssn_id,
		1: _parse_common,
		2: _parse_common,
		3: _parse_common,
		4: _parse_netpay,
	}

	write_to_excel(dic[col](data), col)


if __name__ == '__main__':

	from imageconfig import IMAGE as config


	OUTPUT_IMG_EXCEL_DIR = config.OUTPUT_IMG_EXCEL_DIR

	OUTPUT_IMG_ZIP_NAME = config.OUTPUT_IMG_ZIP_NAME

	OUTPUT_IMG_ZIP = config.OUTPUT_IMG_ZIP

	OUTPUT_LOG_DIR = config.OUTPUT_LOG_DIR

	INPUT_IMAGE_DIR = config.INPUT_IMAGE_DIR

	OUTPUT_ERROR_DIR = config.OUTPUT_ERROR_DIR

	# setting up the logging
	log_filename = 'image_to_excel.log'
	filepath = os.path.join(OUTPUT_LOG_DIR,log_filename)

	logging.basicConfig(filename=filepath,format='[%(asctime)s] %(levelname)s: %(message)s',level=logging.DEBUG)


	if isinstance(INPUT_IMAGE_DIR,str):
		input_path = Path(INPUT_IMAGE_DIR)
		types = ('*.png','*.jpg')
		imgs = []
		for i in types:
			imgs.extend(input_path.glob(i))
		# print(imgs)
		temp_s = 0
		for img in imgs:
			print('='*25)
			print(f'Processing Image {temp_s+1}',os.path.basename(img))
			inp_filename, extension = os.path.basename(img).split('.')
			excel_file = inp_filename+time.strftime("_%m_%d_%Y__%H_%M_%S")+'.xlsx'
			NAME = os.path.join(OUTPUT_IMG_EXCEL_DIR,excel_file)
			# print(img,NAME)
			im = Image.open(img)
			temp_s+=1
			# container for storing parsed values of each row. it'll will have max 5 elements as there are 5 colums.
			# used in function write_to_excel
			aex = []

			# this acts as global row from which main table data will start.
			# all the rows above variable grow will be header rows
			grow = 3

			# count of all the records
			record = 0

			# creating the workbook and worksheet
			workbook, worksheet, validated = create_worksheet(NAME)
			if not validated:
				print(img,OUTPUT_ERROR_DIR)
				shutil.move(str(img),OUTPUT_ERROR_DIR)
				message = f'[Error] {os.path.basename(img)} is not a valid image file. Moved to error folder.'
				print(message)
				logging.error(message)
			else:

				# start and end point of each row
				rows = [(50, 290, 2151, 443),
						(50, 443, 2151, 569),
						(50, 569, 2151, 721),
						(50, 721, 2151, 899),
						(50, 899, 2151, 1025),
						(50, 1025, 2151, 1203),
						(50, 1203, 2151, 1329),
						(50, 1329, 2151, 1456)]

				# column width of each cell
				col = [320, 556, 484, 483, 252]
				for idx, rb in enumerate(rows):
					start_xy = rb[:2]
					for i in range(5):
						end_xy = (start_xy[0] + col[i], rb[3])
						cell_border = start_xy + end_xy
						cell_img = im.crop(cell_border)
						cell_data = image_to_string(cell_img)
						try:
							parse_and_write(cell_data, i)
						except Exception as e:
							print('Exception occured at area: ', cell_border,
								  '\nProcessing again with b/w threshold image...')
							cell_img = convert_to_b_w(cell_img)
							cell_data = image_to_string(cell_img)
							parse_and_write(cell_data, i)
						# print(f"[Row : {idx+1} Col : {i+1}]",repr(cell_data))
						start_xy = (end_xy[0], start_xy[1])

				msg = f"--> Total Records Processed: {record} from image :{os.path.basename(img)}"
				print(msg)
				print('--> LOG File: ',filepath)
				logging.info(msg)
				print('--> Please open: ', NAME)
				workbook.close()
				make_zip(os.path.abspath(img),extension)
		
		if temp_s==0:
			raise Exception(f'{input_path} has no image files.')
	

	

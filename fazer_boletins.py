# Warning: shit coding ahead

make_image_files = True
remove_excel_files = False

data_filename = "dados.xlsx"
template_filename = "modelo.xlsx"
excel_prefix = "temp "
image_prefix = "boletim "

###########

'''
TODO:
- Search for tags only once in template, instead of each time. Like, save the cells and positions of tags and just use them later.
- Remove excel2img, since it reopens the file. Make it so we use openpyxl, maybe in conjunction with win32 API.
- Save and restore clipboard state. Yep, it actually puts stuff in there to save the image. I don't think that there is a function on the API to save the image directly.
- Put all them prints in a disableable debug function
- Make those options above into command line args
- Maybe make a GUI
- Fix/add other things that I can't remember right now
'''

import excel2img
import openpyxl
import os
from collections import namedtuple

fields = []

print("Opening data file '{}'".format(data_filename))

wb = openpyxl.load_workbook(data_filename)
ws = wb.active

print("Getting field names:")

for cell in ws[1]:
	fields.append(str(cell.value))

print(fields)

print("Getting each data row:")

for data_row in ws.iter_rows(min_row=2):

	print("- (Re)opening template file '{}'".format(template_filename))

	wb = openpyxl.load_workbook(template_filename)
	ws = wb.active

	print("- Looping through all non-empty cells to find tags:")

	for template_row in ws.rows:
		for cell in template_row:
			if cell.value != None:

				# find tags in that cell

				cell_str = str(cell.value)
				pos = 0

				print("  - Cell {}: {}".format(cell.coordinate, cell_str))

				while True:

					# find next tag after pos
					percent_pos_start = cell_str.find("%", pos)

					if percent_pos_start != -1:
						# find end of tag
						percent_pos_end = cell_str.find("%", percent_pos_start+1)

						if percent_pos_end == -1:
							print("    - percent_pos_start = {}, no ending percent".format(percent_pos_start))
							break  # no ending %, so there is no more tags (no point in warning, probably an actual percentage sign there

						# find name of tag
						tag_name = cell_str[percent_pos_start+1:percent_pos_end]

						print("    - percent_pos_start = {}, percent_pos_end = {}, tag_name = {}".format(percent_pos_start, percent_pos_end, tag_name))

						# find which field is this tag
						for field_i, field in enumerate(fields):

							if tag_name == field:

								# replace tag name and %s with value in data

								data_value = data_row[field_i].value
								data_type = data_row[field_i].data_type

								print("      - data_value = {}, data_type = {}".format(repr(data_value), data_type), )

								data_value_str = str(data_value) if data_value != None else ""

								# check if the cell only contains the tag, if so copy the type directly

								if percent_pos_start == 0 and percent_pos_end+1 == len(cell_str):

									cell_str = data_value_str
									cell.value = data_value
									# cell.data_type = data_type

									print("        - whole tag cell copied directly, cell.data_type = {}".format(cell.data_type))

								else:

									cell_str = cell_str[:percent_pos_start] + data_value_str + cell_str[percent_pos_end+1:]
									cell.value = cell_str

								pos = percent_pos_start + len(data_value_str)

								print("      - replaced, cell_str = {}".format(cell_str))

								break

						else:
							print("      - no field called {}".format(tag_name))
							pos = percent_pos_end+1

						print("      - pos = {}".format(pos))

					else:
						# no more %s
						break

	excel_filename = excel_prefix + str(data_row[0].value) + ".xlsx"
	image_filename = image_prefix + str(data_row[0].value) + ".png"

	print("- Saving excel file {}".format(excel_filename))

	wb.save(excel_filename)

	if make_image_files:
		print("- Saving image file {}".format(image_filename))
		excel2img.export_img(excel_filename, image_filename, "Plan1", None)

	if remove_excel_files:
		print("- Removing excel file {}".format(excel_filename))
		os.remove(excel_filename)
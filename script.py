import pandas
from openpyxl import load_workbook

def main():
	df = pandas.read_excel(io="raw_data_1.xlsx", sheet_name='Data')

	# parsing input data
	d1 = list(df['Geranic Acid|Amount'])[2:]
	d1s1, d1s2, d1s3, d1b1, d1b2, d1b3 = save_data(d1)
	d1b1_avg_list = calc_avg(d1b1)
	d1b2_avg_list = calc_avg(d1b2)
	d1b3_avg_list = calc_avg(d1b3)

	d2 = list(df['Ibuprofen|Amount'])[2:]
	d2s1, d2s2, d2s3, d2b1, d2b2, d2b3 = save_data(d2)
	d2b1_avg_list = calc_avg(d2b1)
	d2b2_avg_list = calc_avg(d2b2)
	d2b3_avg_list = calc_avg(d2b3)

	# from 48h sheet
	ss = 5
	df2 = pandas.read_excel(io="raw_data_2.xlsx", sheet_name='Data')
	d1_48 = list(df2['Geranic Acid|Amount'])[2:]
	save_48_data(d1_48, d1s1, d1s2, d1s3, d1b1, d1b2, d1b3,
		d1b1_avg_list, d1b2_avg_list, d1b3_avg_list)

	d2_48 = list(df2['Ibuprofen|Amount'])[2:]
	save_48_data(d2_48, d2s1, d2s2, d2s3, d2b1, d2b2, d2b3,
		d2b1_avg_list, d2b2_avg_list, d2b3_avg_list)

	# outputting data

	wb = load_workbook('empty_template.xlsx')

	paste_blank_data(wb, 'Blank S1', d1b1)
	paste_blank_data(wb, 'Blank S2', d2b1)
	paste_data(wb, 'S1', d1s1, d1b1_avg_list)
	paste_data(wb, 'S2', d2s1, d2b1_avg_list)


def save_data(data):
	ss = 15  # samplesize
	return (data[0:ss], data[ss:ss * 2],
		data[ss * 2:ss * 3], data[ss * 3:ss * 4],
		data[ss * 4:ss * 5], data[ss * 5:ss * 6])


def save_48_data(data, s1, s2, s3, b1, b2, b3,
		b1_avg_list, b2_avg_list, b3_avg_list):
	ss = 5
	s1 += data[0:ss]
	s2 += data[ss:ss * 2]
	s3 += data[ss * 2: ss * 3]
	b1 += data[ss * 3: ss * 5]
	b2 += data[ss * 4: ss * 5]
	b3 += data[ss * 5: ss * 6]
	b1_avg_list += [sum(data[ss * 3: ss * 5]) / 5]
	b2_avg_list += [sum(data[ss * 4: ss * 5]) / 5]
	b3_avg_list += [sum(data[ss * 5: ss * 6]) / 5]

def paste_blank_data(wb, worksheet, data):
	ws = wb[worksheet]
	i = 0
	for row in ws.iter_rows(min_row=5, min_col=8, max_row=9, max_col=10):
		for cell in row:
			cell.value = data[i]
			i += 1
	wb.save('empty_template.xlsx')


def paste_data(wb, worksheet, data, avg_list):
	ws = wb[worksheet]
	i, j = 0, 0
	for row in ws.iter_rows(min_row=5, min_col=8, max_row=9, max_col=10):
		for cell in row:
			# print('for ' , cell , data[i] , avg_list[j])
			if data[i] - avg_list[j] < 0:
				cell.value = data[i]
			else:
				cell.value = data[i] - avg_list[j]
			i += 1
			j = j + 1 if j < 2 else 0
	wb.save('empty_template.xlsx')


# append 48h
# returns list of averages
def calc_avg(data):
	lst = [0, 0, 0]
	for i in range(0, 15, 3):
		lst[0] += data[i]
		lst[1] += data[i + 1]
		lst[2] += data[i + 2]
	return [x / 5 for x in lst]


if __name__ == "__main__":
	main()

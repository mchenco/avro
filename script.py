import pandas
from openpyxl import load_workbook

def main():
	df = pandas.read_excel(io="raw_data_1.xlsx", sheet_name='Data')
	df2 = pandas.read_excel(io="raw_data_2.xlsx", sheet_name='Data')

	# parsing input data
	d1 = list(df['Geranic Acid|Amount'])[2:]
	d1_extra = list(df2['Geranic Acid|Amount'])[2:]
	d1s1, d1s2, d1s3, d1b1, d1b2, d1b3 = save_data(d1, d1_extra)
	d1b1_avg = calc_avg(d1b1)
	d1b2_avg = calc_avg(d1b2)
	d1b3_avg = calc_avg(d1b3)

	d2 = list(df['Ibuprofen|Amount'])[2:]
	d2_extra = list(df2['Ibuprofen|Amount'])[2:]

	d2s1, d2s2, d2s3, d2b1, d2b2, d2b3 = save_data(d2, d2_extra)
	d2b1_avg = calc_avg(d2b1)
	d2b2_avg = calc_avg(d2b2)
	d2b3_avg = calc_avg(d2b3)

	# outputting data

	wb = load_workbook('empty_template.xlsx')

	paste_blank_data(wb, 'Blank S1', d1b1)
	paste_blank_data(wb, 'Blank S2', d2b1)
	paste_data(wb, 'S1', d1s1, d1b1_avg)
	paste_data(wb, 'S2', d2s1, d2b1_avg)


def save_data(data1, data2):
	ss1 = 15  # samplesize
	ss2 = 5
	return (
		data1[0:ss1] + data2[0:ss2],
		data1[ss1:ss1 * 2] + data2[ss2:ss2 * 2],
		data1[ss1 * 2:ss1 * 3] + data2[ss2 * 2:ss2 * 3],
		data1[ss1 * 3:ss1 * 4] + data2[ss2 * 3:ss2 * 4],
		data1[ss1 * 4:ss1 * 5] + data2[ss2 * 4:ss2 * 5],
		data1[ss1 * 5:ss1 * 6] + data2[ss2 * 5:ss2 * 6]
	)


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


# returns list of averages
def calc_avg(data):
	lst = [0, 0, 0]
	for i in range(0, 20, 4):
		lst[0] += data[i]
		lst[1] += data[i + 1]
		lst[2] += data[i + 2]
		lst[3] += data[i + 3]
	return [x / 5 for x in lst]


if __name__ == "__main__":
	main()

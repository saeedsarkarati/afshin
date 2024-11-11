import pandas as pd
file = ('f2.xlsx')
data = pd.read_excel(file, dtype= {'phone':str, 'idnum':str, 'name': str, 'surnam': str, 'email': str})
print (data)
data.insert(3, 'iddif' , '')
data.insert(5, 'phonedif' , '')
nrows = data.shape[0]
print (nrows, type(nrows))
# ~ d = data.iloc[1, 0]
for i in range (nrows):
	if isinstance(data.iloc[i, 0], str):
		data.iloc[i, 0] = data.iloc[i,0].strip()
	if isinstance(data.iloc[i, 1], str):
		data.iloc[i, 1] = data.iloc[i,1].strip()
data = data.drop('event', axis = 1)
writer = pd.ExcelWriter('o2.xlsx', engine='xlsxwriter')
d2 = data.sort_values(['phone', 'idnum'])
print(d2)
for i in range (1, nrows):
	d2.iloc[i,3] = str ( int(d2.iloc[i,2]) - int(d2.iloc[i-1,2]) )
for i in range (1, nrows):
	d2.iloc[i,5] = str ( int(d2.iloc[i,4]) - int(d2.iloc[i-1,4]) )
drop = []
for i in range (1, nrows):
	if d2.iloc[i, 0] == d2.iloc[i-1,0]:
		if d2.iloc[i, 1] == d2.iloc[i-1, 1]:
			if d2.iloc[i, 2] == d2.iloc[i-1, 2]:
				if d2.iloc[i, 4] == d2.iloc[i-1, 4]:
					if d2.iloc[i, 6] == d2.iloc[i-1, 6]:
						drop.append(i-1)
					if pd.isna(d2.iloc[i, 6]) and pd.isna(d2.iloc[i-1, 6]):
						drop.append(i-1)


print ('----')
print (drop)
d2 = d2.drop(d2.index[drop])

# ~ while i < data.shape[0]:
	# ~ if 
	# ~ i += 1

d2.to_excel(writer, sheet_name='sheet1')
worksheet = writer.sheets['sheet1']
worksheet.right_to_left()
worksheet.set_column(3, 6, 15)
worksheet.set_column(1, 2, 12)
worksheet.set_column(7, 7, 30)
writer.save()


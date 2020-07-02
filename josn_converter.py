import xlrd
from collections import OrderedDict
import simplejson as json
# Open the workbook and select the first worksheet
wb = xlrd.open_workbook(r'C:\Users\i532970\OneDrive - SAP SE\Desktop\final.xlsx')
sh = wb.sheet_by_index(1)
# List to hold dictionaries
data_list = []
# Iterate through each row in worksheet and fetch values into dict
for rownum in range(1, sh.nrows):
    data = OrderedDict()
    row_values = sh.row_values(rownum)
    data['equipmentID'] = row_values[4]
    data['locationID'] = "4BCD957295B547EE9BF74FE16D0A61D6"
    data['breakdown'] = "0"
    data['type'] = "M2"
    data['priority'] = "15"
    data['status'] =  [
		"CPT"
	]
    data['startDate'] = row_values[0]
    data['endDate'] = row_values[0]
    data['malfunctionStartDate'] = row_values[0]
    data['malfunctionEndDate'] = row_values[0]
    data['shortDescription'] =row_values[1]
    data['longDescription'] =row_values[2]
    data['confirmedFailureModeID'] = row_values[3]
    data_list.append(data)
# Serialize the list of dicts to JSON
j = json.dumps(data_list)
# Write to file
with open(r'C:\Users\i532970\OneDrive - SAP SE\Desktop\data1.json', 'w') as f:
    f.write(j)
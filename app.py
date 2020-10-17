import requests, xlrd, json
from requests.auth import HTTPBasicAuth

with open("config.json", "r") as f:
    my_vars = json.load(f)

instanceURL = my_vars["instanceURL"]
username = my_vars["username"]
password = my_vars["password"]

payrollInfoURL1 = instanceURL + "/hcmRestApi/resources/11.13.18.05/payrollRelationships?q=PersonNumber="
payrollInfoURL2 = "&expand=payrollAssignments,payrollAssignments.assignedPayrolls"
dobURL = instanceURL + "/hcmRestApi/resources/11.13.18.05/emps?onlyData=true&fields=HireDate,DisplayName&q=PersonNumber="
payrollDefURL = instanceURL + "/hcmRestApi/resources/11.13.18.05/payrollDefinitionsLOV?fields=PayrollId,PayrollName&onlyData=true&limit=100"


payrollDefInfo = requests.request("GET", payrollDefURL, auth=HTTPBasicAuth(username, password))
payrollDefData = payrollDefInfo.json()
payrolls = dict()

for i in range(len(payrollDefData['items'])):
	payrolls[payrollDefData['items'][i]['PayrollName']] = str(payrollDefData['items'][i]['PayrollId'])

filename = 'input.xlsx'
book = xlrd.open_workbook(filename)
sheet = book.sheet_by_index(0)
xlsData = [[str(sheet.cell_value(r, c)) for c in range(sheet.ncols)] for r in range(sheet.nrows) if r!=0]
emps = dict()
for i in range(len(xlsData)):
	emps[str(xlsData[i][0]).split('.')[0]] = str(xlsData[i][1]).replace(r"\\xa0", " ")
for p in emps.keys():
	try:
		dobInfo = requests.request("GET", dobURL + p, auth=HTTPBasicAuth(username, password))
		dobData = dobInfo.json()
		payrollInfo = requests.request("GET", payrollInfoURL1 + p + payrollInfoURL2, auth=HTTPBasicAuth(username, password))
		payrollData = payrollInfo.json()
		if len(payrollData['items'][0]['payrollAssignments'][0]['assignedPayrolls']) == 0:
			payrollLinks = payrollData['items'][0]['payrollAssignments'][0]['links']
			for link in payrollLinks:
				if link['name'] == 'assignedPayrolls':
					assignedPayrollURL = link['href']
					upd_payload = "{\r\n    \"PayrollId\": " + payrolls[emps[p].replace(r"\\xa0", " ")] + ",\r\n    \"StartDate\": \"" + dobData['items'][0]['HireDate'] + "\",\r\n    \"EffectiveStartDate\": \"" + dobData['items'][0]['HireDate'] + "\"\r\n}"
					headers = {'content-type': 'application/json' }
					response = requests.post(assignedPayrollURL, auth=HTTPBasicAuth(username, password), headers=headers, data = upd_payload)
					if response.status_code == 200:
						print('Payroll added for ' + dobData['items'][0]['DisplayName'] + ' (' + p)
					elif response.status_code == 401:
						self.textBrowser.setText('Authorization failed. Please validate the instance URL, Username and password (Basic Auth)')
					elif response.status_code == 403:
						self.textBrowser.setText('Insufficient privileges. Please ensure your user account has sufficient access privileges')
		else:
			print('Employee - ' + p + ' already has a payroll assigned')
         
	except:
		print('Failed to add payroll for ' + p + '. Please validate the employee\'s payroll relationship manually.')
		continue


#print(response.text.encode('utf8'))

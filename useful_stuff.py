import win32api, win32con
from win32api import GetSystemMetrics
import datetime
import win32com.client
import teradata
import os

#Simulate Mouse Click
def click(x,y):
    #win32api.SetCursorPos((x,y))
    SCREEN_WIDTH, SCREEN_HEIGHT = GetSystemMetrics(0), GetSystemMetrics(1)
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, int(x/SCREEN_WIDTH*65535.0), int(y/SCREEN_HEIGHT*65535.0))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)

#Simulate Moving the Mouse
def setCursor(x, y):
    SCREEN_WIDTH, SCREEN_HEIGHT = GetSystemMetrics(0), GetSystemMetrics(1)
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, int(x/SCREEN_WIDTH*65535.0), int(y/SCREEN_HEIGHT*65535.0))


UDA_EXEC = teradata.UdaExec(appName="HelloWorld", version="1.0", logConsole=False)
SESSION = UDA_EXEC.connect(method="", dsn="", username="", password="")

xl = win32com.client.Dispatch("Excel.Application")
xl.DisplayAlerts = False

def replace_non_alnum( s_in, s_sub ):
	s = s_in
	for i in range(0, len(s)):
		if s[i:i+1] != "_" and not s[i:i+1].isalnum():
			s = s.replace(s[i:i+1], s_sub)
	
	if s.strip().upper() == "NONE":
		s = "No_Name"
	return s

#upload excel file to teradata
def excel_to_teradata(xl_app, teradata_table, excel_file, excel_sheet, header_location, num_columns, pw=""):
	global UDA_EXEC
	global SESSION
	fname = ""
	
	try:
		import os
	except:
		pass
	
	try:
		wb = xl_app.Workbooks(excel_file)
	except:
		try:
			fname = excel_file[excel_file.rfind("\\")+1:excel_file.rfind("\\") + (len(excel_file)-excel_file.rfind("\\")) ]
			wb = xl_app.Workbooks(fname)
		except:
			try:
				try:
					xl_app.Workbooks.Open(Filename=os.path.abspath(excel_file), Password=pw)
				except:
					xl_app.Workbooks.Open(Filename=os.path.abspath(excel_file))
				fname = excel_file[excel_file.rfind("\\")+1:excel_file.rfind("\\") + (len(excel_file)-excel_file.rfind("\\")) ]
				wb = xl_app.Workbooks(fname)
			except:
				print ("Could not open file:", fname)
				return False
		
	try:
		import teradata
	except:
		pass
	
	try:
		SESSION.execute("DROP TABLE "+teradata_table)
	except:
		pass
	
	tbl_query = """
			CREATE MULTISET TABLE &STAGE_TABLE ,NO FALLBACK ,
			 NO BEFORE JOURNAL,
			 NO AFTER JOURNAL,
			 CHECKSUM = DEFAULT,
			 DEFAULT MERGEBLOCKRATIO
			 (
				&COLUMN_LIST
			 )
			PRIMARY INDEX ( &PRI_INDEX );
			"""
	
	sht = wb.Sheets(excel_sheet)
	
	h = []
	
	for i in range(0, num_columns):
		h.append(replace_non_alnum(str(sht.Range(header_location).GetOffset(0, i).Value).strip(), "_"))
	
	for i in range(0, len(h)-1):
		num = 1
		for j in range(i+1, len(h)):
			if h[j] == h[i]:
				h[j] = h[j] + str(num)
				num = num + 1
	
	col_list = ""
	for i in range(0, len(h)):
		if i > 0:
			col_list = col_list + ",\n"
		col_list = col_list + h[i] + " VARCHAR(255) CHARACTER SET LATIN CASESPECIFIC"
	
	tbl_query = tbl_query.replace("&STAGE_TABLE", teradata_table)
	tbl_query = tbl_query.replace("&COLUMN_LIST", col_list)
	tbl_query = tbl_query.replace("&PRI_INDEX", h[0])
	
	try:
		SESSION.execute(tbl_query)
	except:
		print("could not create table")
		return False
	
	insert_sql = """
				INSERT INTO &STAGE_TABLE
				( 
				&COLUMN_LIST 
				)
				VALUES 
				(
				&VALUE_LIST 
				)
				"""
	
	col_list = ""
	for i in range(0, len(h)):
		if i > 0:
			col_list = col_list + ",\n"
		col_list = col_list + h[i]
	
	insert_sql = insert_sql.replace("&STAGE_TABLE", teradata_table)
	insert_sql = insert_sql.replace("&COLUMN_LIST", col_list)
	
	for i in range(sht.Range(header_location).Row+1, sht.Range(header_location).Row + sht.Range(header_location).CurrentRegion.Rows.Count):
		k = str(sht.Range(header_location).Offset(i-1,0).Value).strip()
		if k == "":
			break
		v_list = ""
		for col in range(0, len(h)):
			print ("db = ", i, col)
			if col > 0:
				v_list = v_list + ",\n"
			if sht.Range(header_location).GetOffset(i-1, col).Value == None:
				v_list = v_list + "NULL"
			else:
				v_list = v_list + "'" + str(sht.Range(header_location).GetOffset(i-1, col).Value).replace("'","''").replace("\"","").strip() + "'"
	
		ins_query = insert_sql.replace("&VALUE_LIST", v_list)
		try:
			#break
			SESSION.execute(ins_query)
		except:
			print("Could Not Insert Row ", i)
	

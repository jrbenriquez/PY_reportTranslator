""" Script Aid for The O2 Weekly Report """
import sys
import codecs
import pandas as pd

def readdb():
	""" Read Database """
	df1 = pd.read_csv('eventDB.csv', encoding='utf-8')

	pidlist = df1['pid'].tolist()
	translateList = df1['translated'].tolist()
	dbdict = {}

	if len(pidlist) == len(translateList):
		for i in range(0,len(pidlist)):
			dbdict[pidlist[i]] = translateList[i]

	return dbdict

def translatereport(dbdict):
	"""Translate the Report """
	df = pd.read_csv(sys.argv[1], encoding='utf-8')
	print "Length of dataframe:" + str(len(df))
	
	#Translate O2 Report
	print "Doing Translation...."
	translatedlist = []
	for i in df['pid']:
		if i in dbdict:
			translatedlist.append(dbdict[i])
		else:
			addtodb(i)
			return 'restart'
	
	df['pid'] = translatedlist
	print 'Done!'
	return df

def removedups(df,duptransdict):
	""" Remove the Duplicates"""
	rowsDeleted = 0
	for trans in duptransdict:
		#While number of duplicates is not zero, drop a row with matching transid
		while duptransdict[trans] > 0:
			dupdf = df[df['transid'] == trans]
			rowToDelete = df.index.get_loc(dupdf.iloc[0].name)
			print "Found Duplicate" + str(trans) + " in ROW: " + str(rowToDelete)
			df = df.drop(df.index[rowToDelete])
			rowsDeleted = rowsDeleted + 1
			duptransdict[trans] = duptransdict[trans] - 1
	print "Rows Deleted: " + str(rowsDeleted)
	return df

def addtodb(trans):
	""" Add to Database """
	df = pd.read_csv('eventDB.csv', encoding='utf-8')
	eventSeries =  pd.Series(df['translated'].unique()).sort_values(ascending=True)
	print eventSeries
	print trans.encode('utf-8') + " not found in Database"
	eventNum = raw_input(trans.encode('utf-8') + ": Enter Event Number if you can see it from the list above. (If not enter 'no')")
	try:
		eventNum = int(eventNum)
		appendDF = pd.DataFrame([[trans.encode('utf-8'),eventSeries[eventNum]]], columns=['pid','translated'])
	except:
		eventName = newevent()
		appendDF = pd.DataFrame([[trans.encode('utf-8'),eventName]], columns=['pid','translated'])
	df = df.append(appendDF)
	df.to_csv('eventDB.csv', index=False, encoding='utf-8')

def newevent():
	""" New Event"""
	return raw_input("What do you want to call this event?")


def main_translate():
	""" Main Translation Workflow """
	dbdict = readdb()
	finalreport = translatereport(dbdict)
	try:
		while finalreport == 'restart':
			dbdict = readdb()
			finalreport = translatereport(dbdict)
	except:
		finalreport.to_csv(sys.argv[1][:-3]+ 'done.csv',index=False, encoding='utf-8')
		return finalreport

def event_pivot():
	""" Main Pivoting Workflow"""
	df = pd.read_csv(sys.argv[1][:-3]+ 'done.csv', encoding='utf-8')

	#Basic Pivot Sum of Tickets per Event
	eventlist = df['pid'].unique()
	eventqtydict = {}
	for event in eventlist:
	    sum = df[df['pid'] == event]['qty'].sum()
	    eventqtydict[event] = sum
	eventqtydf = pd.DataFrame.from_dict(eventqtydict, orient='index')
	eventqtydf.columns = ['Quantity']
	totalqty = int(eventqtydf.sum(0,1))
	totalDict = {'Total' : totalqty }
	dfx = pd.DataFrame.from_dict(totalDict, orient='index')
	dfx.columns = ['Quantity']
	eventqtydf = eventqtydf.sort_values('Quantity',axis=0,ascending=False)
	eventqtydf = eventqtydf.append(dfx)
	return eventqtydf

def quantity_pivot():
	""" Main Pivoting Workflow"""
	df = pd.read_csv(sys.argv[1][:-3]+ 'done.csv', encoding='utf-8')
	#Basic Pivot Sum of Tickets per Day
	datelist = df['conversion_ts'].unique()
	dateqtydict = {}
	for event in datelist:
	    sum = df[df['conversion_ts'] == event]['qty'].sum()
	    dateqtydict[event] = sum
	dateqtydf = pd.DataFrame.from_dict(dateqtydict, orient='index')
	dateqtydf.columns = ['Quantity']
	totalqty = int(dateqtydf.sum(0,1))
	totalDict = {'Total' : totalqty }
	dfx = pd.DataFrame.from_dict(totalDict, orient='index')
	dfx.columns = ['Quantity']
	dateqtydf = dateqtydf.sort_index(axis=0,ascending=True)
	dateqtydf = dateqtydf.append(dfx)
	return dateqtydf

def createreport(dframe, df_event, df_qty):
        """ Creates the final report"""
	# Create a Pandas Excel writer using XlsxWriter as the engine.
	writer = pd.ExcelWriter(sys.argv[1][:-3]+ 'done.xlsx', engine='xlsxwriter')
	# Convert the dataframe to an XlsxWriter Excel object.
	dframe = pd.read_csv(sys.argv[1][:-3]+ 'done.csv', encoding='utf-8')
        dframe.to_excel(writer, sheet_name='Main', index=False)
	df_event.to_excel(writer, sheet_name='Event Pivot')
	df_qty.to_excel(writer, sheet_name='Quantity Pivot')
	eventSheet = writer.sheets['Event Pivot']
	quantitySheet = writer.sheets['Quantity Pivot']
	mainSheet = writer.sheets['Main']
	#Styling
	eventSheet.set_column('A:A', 70)
	quantitySheet.set_column('A:A', 20)
	mainSheet.set_column('A:B', 40)
	mainSheet.set_column('A:A', 20)
        writer.save()

def main():
	""" Main Program Workflow """
        dframe = main_translate()
        df_event = event_pivot()
        df_qty = quantity_pivot()
        createreport(dframe,df_event, df_qty)


if __name__ == '__main__':
	""" Main Program Flow """
	main()
	


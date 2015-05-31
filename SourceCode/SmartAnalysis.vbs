Class AnalyserClass
	Public strSource,strExceptionFilter,arrSourceWords,strTotalWords,arrDictKey
	Function Initial
		strTotalWords=""
		strSource=""
		strExceptionFilter=""
	End Function			
	Function GetStrTotalWords
		arrExceptionFilter=Split(strExceptionFilter,";",-1.1)
		For i=0 To UBound(arrSourceWords)
			strconverttolowercase=LCase(arrSourceWords(i))			 
			For b=0 To UBound(arrExceptionFilter)
				If InStr(strconverttolowercase,LCase(arrExceptionFilter(b)))>0 Then
					ExceptionFlag=False 
					Exit For 
				Else
					ExceptionFlag=True 
				End If
			Next
			If ExceptionFlag And InStr(strTotalWords,strconverttolowercase)=0 Then
				strTotalWords=strTotalWords&strconverttolowercase&";"				
				If Not objDict.Exists(strconverttolowercase) Then
					objDict.Add strconverttolowercase,1
				Else
					objDict.Item(strconverttolowercase)=objDict.Item(strconverttolowercase)+1	
				End If							
			End If 
		Next
	End Function
	Function strCompareWithTotalWords		
		For k=0 To UBound(arrSourceWords)-1
			If objDict.Exists(arrSourceWords(k)) Then
				objDict.Item(arrSourceWords(k))=objDict.Item(arrSourceWords(k))+1
				fff=objDict.Item(arrSourceWords(k))				
			End If
		Next
	End Function	
	Function GetHighFrequencyWords()
		arrDictKey=objDict.Keys
		For h=0 To UBound(arrDictKey)
			For l=h+1 To UBound(arrDictKey)
				If objDict.Item(arrDictKey(l))>objDict.Item(arrDictKey(h)) Then
					temp=arrDictKey(h)
					arrDictKey(h)=arrDictKey(l)
					arrDictKey(l)=temp
				End If 
			Next			
		Next
	End Function
	Function DictReverse
		arrDictKeys=objDict.Keys
		For a=0 To UBound(arrDictKeys)
			If objDictReverse.Exists(objDict.Item(arrDictKeys(a))) Then
				objDictReverse.Item(objDict.Item(arrDictKeys(a)))=objDictReverse.Item(objDict.Item(arrDictKeys(a)))&";"&arrDictKeys(a)
			Else
				objDictReverse.Add objDict.Item(arrDictKeys(a)),arrDictKeys(a)
			End If 
		Next
	End Function
End Class
Class ExcelOperationClass
	Public strExcelFilePath
	Function GetTimeStamp() 'YYYYMMDD 20150320110033
		Dim strTime, strYear, strMonth, strDay, strHour, strMinute, strSecond
		strYear = Cstr(year(Now))
		strMonth = Pad(Cstr(Month(Now)), 2)
		strDay = Pad(Cstr(Day(Now)), 2)
		strHour = Pad(Cstr(Hour(Now)), 2)
		strMinute = Pad(Cstr(Minute(Now)), 2)
		strSecond = Pad(Cstr(Second(Now)), 2)		
		strTime = strYear & strMonth & strDay & strHour & strMinute &strSecond
		GetTimeStamp = strTime
	End Function
	Function Pad(strText, intLen)
		If Len(strText) >= intLen Then 
			strText = strText
		Else
			Do Until Len(strText) = intLen
				strText = "0" & strText
			Loop
		End If		
		Pad = strText
	End Function
	Function Initial
		Set objFSO = CreateObject("Scripting.FileSystemObject") 
		strFolderPath=left(wscript.scriptfullname,instrrev(wscript.scriptfullname,"\")-1) 
		reportfilename=strFolderPath&"\"&"report_"&GetTimeStamp()&".xlsx"
		oExcelReport.Workbooks.Add()
		oExcelReport.ActiveWorkbook.SaveAs reportfilename
		oExcelReport.Quit
		oExcelReport.Workbooks.Open(reportfilename)
		oExcel.Workbooks.Open(strExcelFilePath)
		oExcel.Visible = False
		oExcelReport.Visible=False 
	End Function
	Function GenerateReport
		'set title
		oExcelReport.Worksheets(1).Activate
		oExcelReport.Cells(1,1).value="Smart Analysis Report"
		oExcelReport.Cells(1,1).Interior.colorIndex=41
		oExcelReport.Cells(1,1).font.colorIndex=2
		oExcelReport.Cells(1,1).Font.Size=30
		oExcelReport.Range("A1:G1").merge			
		oExcelReport.Cells(2,1).value="High Frequent Words Exist Time"
		oExcelReport.Cells(2,2).value="High Frequent Words"
		arrWordsNum=objDictReverse.Keys
		arrWords=objDictReverse.Items
		intstep=0
		For intstep=0 To UBound(arrWordsNum)			
			oExcelReport.Cells(intstep+3,1).value=arrWords(intstep)
			oExcelReport.Cells(intstep+3,2).value=arrWordsNum(intstep)			
		Next		
		Set oSheet=oExcelReport.Workbooks(1).Worksheets(1)
		Set achart=oSheet.chartobjects.add(110,55,800,400)
		achart.chart.charttype=5
		set series=achart.chart.seriescollection		
		range="Sheet1!$A$2:$B$"&intstep+2
		series.add range,true 		
	End Function
	Function GetData
		intsheetscount=oExcel.sheets.Count	
		For i = 1 To intsheetscount		
			oExcel.Worksheets(i).Activate			
			'get selected column name
			intcol=1		
			While Not oExcel.Cells(1,intcol).value="Summary" And oExcel.Cells(1,intcol).value<>""
				intcol=intcol+1
			Wend
			'get data
			introw=2		
			While oExcel.Cells(introw,intcol).value<>""				
				LoadData(oExcel.Cells(introw,intcol).value)
				introw=introw+1
			Wend 	
		Next		
	End Function
	Function LoadData(strData)
		analyser.strSource=strData
		analyser.arrSourceWords=Split(analyser.strSource," ",-1,1)
		analyser.GetStrTotalWords
		analyser.strCompareWithTotalWords
	End Function
	Function ClearExcel
		oExcelReport.ActiveWorkbook.Save	
		oExcel.ActiveWorkbook.Save	
		oExcelReport.Quit 
		oExcel.Quit
	End Function
End Class
Class XmlOperationClass
	Function GetExeptionDict
		Set xDoc=CreateObject("MSXML2.DOMDocument")
		If xDoc.load("ExceptionDict.xml") Then 
			strprep=xDoc.selectSingleNode(".//prep").text
			strverb=xDoc.selectSingleNode(".//verb").text
			stradv=xDoc.selectSingleNode(".//adv").text
			strarticle=xDoc.selectSingleNode(".//article").text
			strpron=xDoc.selectSingleNode(".//pron").text
			strothers=xDoc.selectSingleNode(".//others").text
			GetExeptionDict=strprep&";"&strverb&";"&stradv&";"&strarticle&";"&strpron&";"&strothers
		Else 
			MsgBox "ExceptionDict.xml missing please check it and confirm whether it exists in current folder"
			WScript.Quit
		End If 
	End Function
	Function GetFilePath
		Set xDoc=CreateObject("MSXML2.DOMDocument")
		xDoc.load "config.xml"
		If xDoc.load("config.xml") Then
			Set fso1=WScript.CreateObject("scripting.filesystemobject")
			strFolderPath=left(wscript.scriptfullname,instrrev(wscript.scriptfullname,"\")-1) 
			Set fsoFolder1=fso1.GetFolder(strFolderPath) 			
			GetFilePath=fsoFolder1&"\"&xDoc.selectSingleNode(".//FilePath").text
'			If GetFilePath="" Then
'				Set fso=CreateObject("Scripting.FileSystemObject")
'				strFolderPath=left(wscript.scriptfullname,instrrev(wscript.scriptfullname,"\")-1) 
'				Set fso=WScript.CreateObject("scripting.filesystemobject")
'				Set fsoFolder=fso.GetFolder(strFolderPath)
'				For Each filename In fsofolder.Files
'					If InStr(filename,".xls")>0 Then 
'						a=a&";"&filename
'					End If 
'				Next	
'				arra=Split(a,";",-1,1)
'				GetFilePath=strFolderPath&"\"&arra(0)
'			End If			
		Else 
			MsgBox "config.xml missing please check it and confirm whether it exists in current folder"
			WScript.Quit
		End If 		
		
	End Function
End Class

Set EO=New ExcelOperationClass
Set analyser=New AnalyserClass
Set XO=New XmlOperationClass

Set objDict=CreateObject("Scripting.Dictionary")
Set objDictReverse=CreateObject("Scripting.Dictionary")
Set oExcel=WScript.CreateObject("Excel.Application")
Set oExcelReport=WScript.CreateObject("Excel.Application")
EO.strExcelFilePath=xo.GetFilePath
analyser.Initial
analyser.strExceptionFilter=xo.GetExeptionDict
EO.Initial
EO.GetData
analyser.GetHighFrequencyWords
analyser.DictReverse
EO.GenerateReport
EO.ClearExcel
aa=objDictReverse.Keys
bb=objDictReverse.Items

MsgBox "Analysis Finished!!!"&vbCrLf&"@Neo Support"&vbCrlf&"Neo's WebSite: http://xneo123.github.io/"&vbCrlf&"Neo's Github:https://github.com/xneo123"


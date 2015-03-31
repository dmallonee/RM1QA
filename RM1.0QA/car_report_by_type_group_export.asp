<%@ Language=VBScript %>

<%
  
	'On Error Resume Next
 
 	Dim strReportType
 	Dim intReportType
 	Dim strFileName
 	
 	
 	Select Case Request.QueryString("reportformat")
 	
 		Case 0
 			intReportType = 0
 			strReportType = "xls"
 			strFileName = "Rate-Monitor Report " & Request.QueryString("reportrequestid") & "." & strReportType 
 		
 		
 		Case 1
 			intReportType = 1
 			strReportType = "csv"
 			strFileName = "Rate-Monitor Report " & Request.QueryString("reportrequestid") & "." & strReportType 
 		

 		Case 2
 			intReportType = 2
 			strReportType = "txt"
 			strFileName = "Rate-Monitor Report " & Request.QueryString("reportrequestid") & "." & strReportType 

 		Case 3
 			intReportType = 3
 			strReportType = "htm"
 			strFileName = "Rate-Monitor Report " & Request.QueryString("reportrequestid") & "." & strReportType 


 		Case 6
 			intReportType = 6
 			strReportType = "xml"
 			strFileName = "Rate-Monitor Report " & Request.QueryString("reportrequestid") & "." & strReportType 

 		
 		Case Else
 			intReportType = 1
 			strReportType = "csv"
 			strFileName = "Rate-Monitor Report " & Request.QueryString("reportrequestid") & "." & strReportType 
 		
 		
 	End Select

 
	set conn = Server.CreateObject("ADODB.Connection")
	conn.Open Session("pro_con")
	Set rs = Server.CreateObject("ADODB.Recordset")

	rs.Open  "select * from view_car_rate_for_export_group WHERE [Report Number] = " & Request.QueryString("reportrequestid") & " ORDER BY [Pick-up City], [Pick-up Date], [Car Group], Vendor", conn 
	
	'Set adoCmd5 = CreateObject("ADODB.Command")

	'adoCmd5.ActiveConnection = Session("pro_con")

	'adoCmd5.CommandText = "select * from view_car_rate_for_export_group WHERE [Report Number] = 75264"
	'adoCmd5.CommandType = 4
		
	'Set rs = adoCmd5.Execute

    Set Obj = Server.CreateObject("nonnoi_ASPExport2ExcelPack.ASPExport2ExcelPack")

	Obj.RegisterName = "Rate-Highway, Inc."
	Obj.RegisterKey = "127161A4159B9CAC-7398"

	' return CSV to file
	Obj.DateFormat = 1
	Obj.FourDigitYear = False
	Obj.Separator = ","
	If intReportType = 0 Then
		Obj.Header =  "Please Note: Excel will not display the dates properly unless you format the date columns with a date format"
	End If
	Obj.Footer =  vbCrLf & "Rate-Monitor.com is a product of Rate-Highway, Inc. (c) 1999 - 2013 " & vbCrLf & ""
	'Obj.ExportTypeStr = strReportType 
	Obj.ExportType = intReportType
	Obj.ShowExport rs, -1, strFileName
	'Obj.SaveToFile rs, "Rate-Monitor Report " & Request.QueryString("reportrequestid") & "." & strReportType, -1

	rs.Close
	conn.Close

	   set rs = nothing
	   set Obj = nothing

			If err.number = 0 Then
				'Server.Transfer "search_queue_car.asp"	
			Else
			
			   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
			   response.write "<b>VBScript Errors Occured!<br>"
			   response.write "</b><br>"
			   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
			   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
			   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
			   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
			   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
			
			'	Session("error_msg") = "An error was encountered while request your search. Please contact Rate-Highway support"
				'Server.Transfer "search_criteria_car.asp"
			End If


%>
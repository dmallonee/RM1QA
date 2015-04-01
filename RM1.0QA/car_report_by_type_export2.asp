<%@ Language=VBScript %>

<%
Response.ContentType = "application/octet-stream"
Response.ContentType = "application/vnd.ms-excel"


	'On Error Resume Next
 
 	Dim strReportType
 	Dim intReportType
 	Dim strFileName
 	
 	
 	Select Case Request.QueryString("reportformat")
 	
 		Case 0
 			intReportType = 0
 			strReportType = "xls"
 			strFileName = "Rate-Monitor Report " & Request.QueryString("reportrequestid") & "." & strReportType 
            Response.AddHeader "Content-Disposition", "attachment; filename=" & strFileName		
 		
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

	rs.Open  "select * from view_car_rate_for_export_simple WHERE [Report Number] = " & Request.QueryString("reportrequestid"), conn 
    Response.Write "<table>"
    Response.Write "<tr>"
    Response.Write "<td>Report Number</td>"
    Response.Write "<td>Pick-Up City</td>"
    Response.Write "<td>Return City</td>"
    Response.Write "<td>Car Type</td>"
    Response.Write "<td>Data Source</td>"
    Response.Write "<td>Currency</td>"
    Response.Write "<td>LOR</td>"
    Response.Write "<td>Pick-Up Date</td>"
    Response.Write "<td>Pick-Up Time</td>"
    Response.Write "<td>Return Date</td>"
    Response.Write "<td>Return Time</td>"
    Response.Write "<td>Vendor</td>"
    Response.Write "<td>Base Rate</td>"
    Response.Write "<td>Total Rate</td>"
    Response.Write "<td>Total Price</td>"
    Response.Write "</tr>"
    While NOT rs.EOF 
    Response.Write "<tr>"
    Response.Write "<td>" & rs.Fields("Report Number").Value & "</td>"
    Response.Write "<td>" & rs.Fields("Pick-up City").Value & "</td>"
    Response.Write "<td>" & rs.Fields("Return City").Value & "</td>"
    Response.Write "<td>" & rs.Fields("Car Type").Value & "</td>"
    Response.Write "<td>" & rs.Fields("Data Source").Value & "</td>"
    Response.Write "<td>" & rs.Fields("Currency").Value & "</td>"
    Response.Write "<td>" & rs.Fields("lor").Value & "</td>"
    Response.Write "<td>" & rs.Fields("Pick-up Date").Value & "</td>"
    Response.Write "<td>" & rs.Fields("Pick-up Time").Value & "</td>"
    Response.Write "<td>" & rs.Fields("Return Date").Value & "</td>"
    Response.Write "<td>" & rs.Fields("Return Time").Value & "</td>"    
    Response.Write "<td>" & rs.Fields("Vendor").Value & "</td>"    
    Response.Write "<td>" & rs.Fields("Base Rate").Value & "</td>"    
    Response.Write "<td>" & rs.Fields("Total Rate").Value & "</td>"    
    Response.Write "<td>" & rs.Fields("Total Price").Value & "</td>"    
    Response.Write "</tr>"
    Response.Flush
	    rs.MoveNext
    Wend
    Response.Write "</table>"
	'Set adoCmd5 = CreateObject("ADODB.Command")

	'adoCmd5.ActiveConnection = Session("pro_con")

	'adoCmd5.CommandText = "select * from view_car_rate_for_export_simple WHERE [Report Number] = 75264"
	'adoCmd5.CommandType = 4
		
	'Set rs = adoCmd5.Execute

'    Set Obj = Server.CreateObject("nonnoi_ASPExport2ExcelPack.ASPExport2ExcelPack")

'	Obj.RegisterName = "Rate-Highway, Inc."
'	Obj.RegisterKey = "127161A4159B9CAC-7398"

	' return CSV to file
'	Obj.DateFormat = 1
'	Obj.FourDigitYear = False
'	Obj.Separator = ","
'	If intReportType = 0 Then
'		Obj.Header =  "Please Note: Excel will not display the dates properly unless you format the date columns with a date format"
'	End If
'	Obj.Footer =  vbCrLf & "Rate-Monitor.com is a product of Rate-Highway, Inc. (c) 1999 - 2005 " & vbCrLf & ""
	'Obj.ExportTypeStr = strReportType 
'	Obj.ExportType = intReportType
'	Obj.ShowExport rs, -1, strFileName
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
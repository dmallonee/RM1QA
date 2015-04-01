<%@ Language=VBScript %>
<%
	'on error resume next
 
	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd2 = Server.CreateObject("ADODB.Command")
	Set adoRS2  = Server.CreateObject("ADODB.Recordset")
	Set adoConn = Server.CreateObject("ADODB.Connection")
	
	adoConn.Open Session("pro_con")

	
	'adoCmd2.ActiveConnection =  strConn
	'adoCmd2.CommandText = "car_rate_rule_grid"
	'adoCmd2.CommandType = 4

	'adoCmd2.Parameters.Append adoCmd2.CreateParameter("@user_id", 3, 1, 0, strUserId)
		
	'Set adoRS2 = adoCmd2.Execute
  
  	adoRS2.Open  "EXEC car_rate_rule_grid " & strUserId, adoConn
 
    Set Obj = Server.CreateObject("nonnoi_ASPExport2ExcelPack.ASPExport2ExcelPack")

	Obj.RegisterName = "Rate-Highway, Inc."
	Obj.RegisterKey = "127161A4159B9CAC-7398"

	' return CSV to browser
	Obj.DateFormat = 1
	Obj.FourDigitYear = False
	Obj.Separator = ","
'	Obj.Footer = "Hello folks" & vbCrLf & "Hello again "
	Obj.ExportTypeStr = "xls"
	Obj.ShowExport adoRS2, -1, "Rate_Monitor_rule_cross_reference.xls"

	adoRS2.Close
	adoConn.Close

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
<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180
   
   On Error Resume Next
 
	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim adoPrices
	Dim intIndex
	Dim blnExit
	Dim intReportRequestID
	Dim intErrorCount
	Dim intTimeoutCount
	Dim strCarType 
	Dim intResults
	Dim intPrice
	Rem we have no clue how many, so cross your fingers
	Dim varCarTypes()
	Dim varDataSources()
	Dim varDates()
	Dim intRefresh
	Dim strTime

	strClientUserid = Request.Cookies("rate-monitor.com")("user_id") 'Request("userid")
	strCity = Request("city")
	strCarType = Request("car_type")
	strCompany = Request("company")
	intDaysBack = Request("days_back")
	
	If IsNumeric(intDaysBack) = False Then
		intDaysBack = 0
	End If
	
	strSearched = False
	
	If Request("refresh") = "" Then
		intRefresh = 300
	Else
		intRefresh = Request("refresh")
		If IsNumeric(intRefresh) = False Then
			intRefresh = 300
		End If
	End If
	
		strConn = Session("pro_con")
	
  		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_request_rule_select"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("@org_id", 3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@days_back", 3, 1, 0)
		
	
		adoCmd.Parameters("@org_id").Value = strClientUserid 
		adoCmd.Parameters("@days_back") = intDaysBack 
		
		Set adoRS = adoCmd.Execute
	

	If strClientUserId = "" Then
		strClientUserId = Request.Cookies("rate-monitor.com")("client_userid")
	End If

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write parm_msg & "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	End If


	%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Search Queue</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<% If intDaysBack > 0 Then %>
<% Else %>
<meta http-equiv="refresh" content="<%=intRefresh %>;url=request_rule_queue_car.asp">
<% End If %>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script type='text/javascript' language='javascript' src="inc/sitewide.js" ></script>
<script type='text/javascript' language='javascript' src="inc/header_menu_support.js" ></script>
<script type='text/javascript' language='javascript' >
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//
// Page submition section
//
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

function confirmSubmit(SubmitType) {

	if (SubmitType == 'cancel'){
		document.queue.action = 'cancel_search_car.asp'
		//document.search_criteria.submit;
		//document.queue.submit;
		//return true;   
		}

	if (SubmitType == 'redo'){
		document.queue.action = 'redo_search_car.asp'
		//document.search_criteria.submit;
		//document.queue.submit;
		//return true;   
		}


	if (SubmitType == 'redoremail'){
		//alert('redoremail');
		document.queue.action = 'redoremail_search_car.asp'
		//document.search_criteria.submit;
		//document.queue.submit;
		//return true;   
		}

	if (SubmitType == 'forcecomplete'){
		document.queue.action = 'forcecomplete_search_car.asp'
		//document.search_criteria.submit;
		//document.queue.submit;
		//return true;   
		}



}

</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif">
    <img src="images/top.jpg" width="770" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/b_tile.gif">
    <!-- #INCLUDE FILE="inc/page_header_buttons.htm" -->
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/med_bar_tile.gif">
    <img src="images/med_bar.gif" width="12" height="8"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/user_tile.gif">
    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/user_left.gif" width="580" height="31"></td>
        <td background="images/user_tile.gif">
        <table width="100" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td valign="bottom">
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
                <div align="right">
                  <a href="default.asp"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">User: <%=Request.Cookies("rate-monitor.com")("user_name")%></font></div>
                </td>
                <!--
                <td><a href="http://www.rate-monitor.com">
                <img src="images/logout.gif" width="54" height="19" align="middle" border="0" ></a>
                </td>
                -->
              </tr>
              <tr>
                <td><img src="images/separator.gif" width="183" height="6"></td>
              </tr>
            </table>
            </td>
            <td><img src="images/user_tile.gif" width="7" height="31"></td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/h_tile.gif">
    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/h_search_que.gif" width="368" height="31"></td>
        <td><img src="images/h_right.gif" width="402" height="31"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
&nbsp;
<form method="POST" action="request_rule_queue_car.asp" name="search" class="search">
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1710" height="4">
    <tr>
      <td background="images/ruler.gif"></td>
     </tr>
  </table>
  <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1710" id="AutoNumber1" background="images/alt_color.gif">
    <tr>
      <td width="19" height="18">&nbsp;</td>
      <td width="178">&nbsp;</td>
      <td width="478" height="18" colspan="3">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">To search enter 
      the last name, or a portion of. You may also optionally enter city, car type 
      and/or the car company.</font></td>
      <td width="572" height="18">&nbsp;</td>
      <td width="571" height="18">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="26">&nbsp;</td>
      <td width="178"><img border="0" src="images/search.GIF"></td>
      <td width="137" height="26">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><label for="userid" >User Name:</label>
      </font></td>
      <td width="211" height="26">
      <input id="userid" type="text" name="userid" size="20" value="<%=strClientUserid %>" onfocus="this.className='focus';" onblur="this.className='';">
      </td>
      <td width="1207" height="26" colspan="3">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <input type="submit" value="  Display  " name="submit" class="rh_button"></font></td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="178">&nbsp;</td>
      <td width="137" height="22">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">City:</font></td>
      <td width="211" height="22">
      <input type="text" name="city" size="20" value="<%=strCity %>" onfocus="this.className='focus';cl(this,'<%=strCity %>');" onblur="this.className='';fl(this,'<%=strstrCity %>');"></td>
      <td width="1207" height="22" colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="178">&nbsp;</td>
      <td width="137" height="22">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Car Type:</font></td>
      <td width="211" height="22">
      <input type="text" name="car_type" size="20" value="<%=strCarType %>" onfocus="this.className='focus';cl(this,'<%=strCarType %>');" onblur="this.className='';fl(this,'<%=strCarType %>');"></td>
      <td width="1207" height="22" colspan="3"></td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="178">&nbsp;</td>
      <td width="137" height="22">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Company:</font></td>
      <td width="211" height="22">
      <input type="text" name="company" size="20" value="<%=strCompany %>" onfocus="this.className='focus';cl(this,'<%=strCompany %>');" onblur="this.className='';fl(this,'<%=strCompany %>');"></td>
      <td width="1207" height="22" colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="18">&nbsp;</td>
      <td width="178">&nbsp;</td>
      <td width="137" height="18">&nbsp;</td>
      <td width="211" height="18">&nbsp;</td>
      <td width="1207" height="18" colspan="3">&nbsp;</td>
    </tr>
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1710" height="4">
    <tr>
      <td background="images/ruler.gif"></td>
    </tr>
  </table>
</form>
<table width="1710" border="0" cellpadding="2" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
  <tr valign="bottom">
    <td width="625"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
    &nbsp;&nbsp;| <a href="javascript:confirmSubmit('cancel');document.queue.submit();" >Cancel</a> 
<% If Request.Cookies("rate-monitor.com")("user_id") = 3 Then %>
| <a href="javascript:confirmSubmit('redo');document.queue.submit();" >Redo</a> | <a href="javascript:confirmSubmit('redoremail');document.queue.submit();">
Redo &amp; Remail</a> | <a href="javascript:confirmSubmit('forcecomplete');document.queue.submit();">Force Complete</a>   
<% End If %>
| <a href="search_queue_car_ex.asp">Extended view</a> | Standard view |</font></td><td >
    <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
    <!--
    <a href="search_queue_car.asp">|&lt;</a>
    <a href="search_queue_car.asp">&lt;</a> Page 1 of 1
    <a href="search_queue_car.asp">&gt;</a> <a href="search_queue_car.asp">&gt;|</a>
    -->
    <a href="search_queue_car.asp?days_back=0&user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&company=<%=strCompany%>">Today searches</a> or number of days back:
    <a href="search_queue_car.asp?days_back=1&user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&company=<%=strCompany%>">1</a> | 
    <a href="search_queue_car.asp?days_back=2&user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&company=<%=strCompany%>">2</a> | 
    <a href="search_queue_car.asp?days_back=3&user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&company=<%=strCompany%>">3</a> | 
    <a href="search_queue_car.asp?days_back=4&user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&company=<%=strCompany%>">4</a> | 
    <a href="search_queue_car.asp?days_back=5&user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&company=<%=strCompany%>">5</a>
    
    </font></td>
  </tr>
</table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1710" height="4">
  <tr>
    <td background="images/ruler.gif"></td>
  </tr>
</table>
<form name="queue" method="POST" >
<table border="1" style="border-collapse: collapse" bordercolor="#FFFFFF" id="profiles" width="1710">
<thead >
  <tr>
    <td align="left" valign="bottom" bgcolor="#879AA2" height="45" width="26">&nbsp;</td>
    <td class="profile_header" width="63" style="background-color: #E07D1A" height="45">Selected</td>
    <td class="profile_header" width="45" height="45">Search ID</td>
<% If Request.Cookies("rate-monitor.com")("user_id") = 3 Then %>
    <% End If %>    
    <td class="profile_header" width="86" height="45">User</td>
    <td class="profile_header" width="400" height="45">Profile<br>(hover over to view complete name)</td>
    <td class="profile_header" width="400" height="45">Rule<br>(hover over to view complete name)</td>
    <td class="profile_header" width="93" height="45">Source</td>
    <td class="profile_header" width="70" height="45">Pickup City</td>
    <td class="profile_header" width="77" height="45">First Rental Date</td>
    <td class="profile_header" width="77" height="45">Last Rental Date</td>
    <td class="profile_header" height="45">Car Types<br>(hover over to view all types)</td>
    <td class="profile_header" height="45">Companies<br>(hover to view all companies)</td>
  </tr>
  </thead> 
  <%
        
        Dim strClass
        Dim strOrange
        Dim intCount

		If (adoRS.State = adStateOpen) Then

		While adoRS.EOF = False
		
			If strClass = "profile_light" Then
				strClass = "profile_dark"
				strOrange = "bgcolor='#E07D1A'"
			Else
				strClass = "profile_light"
				strOrange = "bgcolor='#FDC677'"
			End If
			
			intCount = intCount + 1
			
		%>
  <tr>
    <td class="<%=strClass %>" height="20"><%=intCount  %></td>
    <td   bgcolor="#FDC677" align="center" height="20">
    <input type="checkbox" value='<%=adoRS.Fields("shop_request_id").Value %>' name="shop_request_id"></td>
   
    
    <td class="<%=strClass %>" height="20" width="45" title="click report number to view in new window">
    <a href="car_report_by_type.asp?ReportRequestID=<%=adoRS.Fields("shop_request_id").Value %>" target="_blank">
    <%=adoRS.Fields("shop_request_id").Value %></a></td>
<!--
    <a href="view_report_car2.asp?ReportRequestID=<%=adoRS.Fields("shop_request_id").Value %>" target="_blank">
    <%=adoRS.Fields("shop_request_id").Value %></a></td>
-->
<% If Request.Cookies("rate-monitor.com")("user_id") = 3 Then %>
	<% End If %>

    <td class="<%=strClass %>" height="20" width="86"><%=adoRS.Fields("client_userid").Value %></td>
    <td class="<%=strClass %>" height="20" width="400" title="<%=adoRS.Fields("profile_desc").Value %>">
		<a href="car_report_by_type.asp?ReportRequestID=<%=adoRS.Fields("shop_request_id").Value %>" target="_blank">
		<% If Len(adoRS.Fields("profile_desc").Value) > 35 Then %>
		  <%=Left(adoRS.Fields("profile_desc").Value, 35) & "..." %>
		<% Else %>
		  <%=adoRS.Fields("profile_desc").Value %>
		<% End If %>
		</a>
	</td>
    <td class="<%=strClass %>" height="20" width="400" title="<%=adoRS.Fields("alert_desc").Value %>">
    	<% If adoRS.Fields("alert_desc").Value <> "No rules assigned" Then %>
		 <a target="_self" title="<%=adoRS.Fields("alert_desc").Value %>" href="alerts_rate_management_car.asp?rateruleid=<%=adoRS.Fields("rate_rule_id").Value %>">		
		 <% If Len(adoRS.Fields("alert_desc").Value) > 35 Then %>
		  <%=Left(adoRS.Fields("alert_desc").Value, 35) & "..." %>
		<% Else %>
		  <%=adoRS.Fields("alert_desc").Value %>
		<% End If %>
		</a>
		<% Else %>
		  <%=adoRS.Fields("alert_desc").Value %>
		<% End If %>
	</td>

    <td class="<%=strClass %>" height="20" title="<%=adoRS.Fields("data_sources").Value %>">
    <% Select Case adoRS.Fields("data_sources").Value
    
    	Case "SA1"
    %>
	    <%="GDS" %></td>
    <%
    	Case "SA2"
    %>
	    <%="GDS Gov" %></td>
    <%
    	Case "BRD"
    %>
	    <%="Brand" %></td>
	<%
	    Case "TDT"
	%>    
		<%="GDS" %></td>
    <%	
    	Case Else
    	
    		If Left(adoRS.Fields("data_sources").Value, 1) = "V" Then
    %>
		    <%="Brand"  %> <% Rem Left(adoRS.Fields("data_sources").Value, 4) & "..." %></td>
    <%	
    
    		Else

    %>
		    <%=adoRS.Fields("data_sources").Value %></td>
    <%	
    		
    		
    		End If
    		
    	End Select
   	%>
    
    <td class="<%=strClass %>" height="20" title="<%=adoRS.Fields("city_cd").Value %>">
    <% If Len(adoRS.Fields("city_cd").Value) > 18 Then %>
    	<%=Left(adoRS.Fields("city_cd").Value, 16) & "..." %>
	<% Else %>
    	<%=adoRS.Fields("city_cd").Value %>
	<% End If %>
	</td> 
    <td class="<%=strClass %>" height="20">
    <% If Len(Month(adoRS.Fields("begin_arv_dt").Value)) = 1 Then
    	strTime = "0" & Month(adoRS.Fields("begin_arv_dt").Value) & "/"
       Else
    	strTime = Month(adoRS.Fields("begin_arv_dt").Value) & "/"
       End If
       
	   If Len(Day(adoRS.Fields("begin_arv_dt").Value)) = 1 Then
    	strTime = strTime & "0" & Day(adoRS.Fields("begin_arv_dt").Value) & "/"
       Else
    	strTime = strTime & Day(adoRS.Fields("begin_arv_dt").Value) & "/"
       End If
       
 	   strTime = strTime & "0" & Year(adoRS.Fields("begin_arv_dt").Value) - 2000

	%>
	<%=strTime %>
    <!--   
    <%=adoRS.Fields("begin_arv_dt").Value %>
    -->
    </td>
    <td class="<%=strClass %>" height="20">
    <% If Len(Month(adoRS.Fields("end_arv_dt").Value)) = 1 Then
    	strTime = "0" & Month(adoRS.Fields("end_arv_dt").Value) & "/"
       Else
    	strTime = Month(adoRS.Fields("end_arv_dt").Value) & "/"
       End If
       
	   If Len(Day(adoRS.Fields("end_arv_dt").Value)) = 1 Then
    	strTime = strTime & "0" & Day(adoRS.Fields("end_arv_dt").Value) & "/"
       Else
    	strTime = strTime & Day(adoRS.Fields("end_arv_dt").Value) & "/"
       End If
       
 	   strTime = strTime & "0" & Year(adoRS.Fields("end_arv_dt").Value) - 2000

	%>
	<%=strTime %>
    <!--   
    
    <%=adoRS.Fields("end_arv_dt").Value %>
    -->
    </td>
    <td class="<%=strClass %>" height="20" title="<%=adoRS.Fields("shop_car_type_cds").Value %>">
    <% If Len(adoRS.Fields("shop_car_type_cds").Value) > 25 Then %>
    	<%=Left(adoRS.Fields("shop_car_type_cds").Value, 25) & "..." %>
	<% Else %>
    	<%=adoRS.Fields("shop_car_type_cds").Value %>
	<% End If %>
	</td> 
    <td class="<%=strClass %>" height="20" title="<%=adoRS.Fields("vend_cd").Value %>">
    <% If adoRS.Fields("vend_cd").Value = "" Then %>
    All
    <% Else %>
	    <% If Len(adoRS.Fields("vend_cd").Value) > 26 Then %>
		    <%=Left(adoRS.Fields("vend_cd").Value, 24) & "..." %>
		<% Else %>
	    	<%=adoRS.Fields("vend_cd").Value %>
		<% End If %>
    <% End If %>
    </td>
  </tr>
  <%
        
        	adoRS.MoveNext
        	
        Wend
        
   		adoRS.Close
		Set adoRS = Nothing
		Set adoCmd = Nothing

		Else

		%>
		
		
  <tr>
    <td class="profile_light" height="20"></td>
    <td bgcolor="#FDC677" align="center" height="20">
    <input type="radio" value="V1" name="selected"></td>
    <td class="profile_light" height="20" width="45"></td>
    <td class="profile_light" height="20" width="86"></td>
    <td class="profile_light" height="20"></td>
    <td class="profile_light" height="20"></td>
    <td class="profile_light" height="20"></td>
    <td class="profile_light" height="20"></td>
    <td class="profile_light" height="20"></td>
    <td class="profile_light" height="20"></td>
    <td class="profile_light" height="20"></td>
  </tr>
<!--
  <tr>
    <td width="26" class="profile_light" height="20"></td>
    <td width="63" bgcolor="#FDC677" align="center" height="20">
    <input type="radio" value="V1" name="selected"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="54" class="profile_light" height="20"></td>
    <td width="60" class="profile_light" height="20"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
    <td width="13" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
  </tr>

-->
  <%

		End If
		        
        %>
</table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1710" height="4">
  <tr>
      <td background="images/ruler.gif"></td>
  </tr>
</table>
<p>&nbsp;| <a href="javascript:confirmSubmit('cancel');document.queue.submit();" >Cancel</a> 
<% If Request.Cookies("rate-monitor.com")("user_id") = 3 Then %>
| <a href="javascript:confirmSubmit('redo');document.queue.submit();" >Redo</a> | <a href="javascript:confirmSubmit('redoremail');document.queue.submit();">
Redo &amp; Remail</a> | <a href="javascript:confirmSubmit('forcecomplete');document.queue.submit();">Force Complete</a>    
<% End If %>
| <a href="search_queue_car_ex.asp">Extended view</a> | Standard view |</p>
</form>
<% If intDaysBack > 0 Then %>
<font size="2" face="Verdana, Arial, Helvetica, sans-serif">
<p align="center"><blink><i>You are viewing a prior day's queue, this page will not 
auto-refresh</i></blink></p>
</font>
<% End If %>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
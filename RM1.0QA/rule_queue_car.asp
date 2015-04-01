<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check_ex.asp" -->
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
	Dim strDaysBack 

'Declare variables
Dim intCurrentPage
Dim iPageSize
Dim i
Dim oConnection
Dim oRecordSet
Dim oTableField
Dim sPageURL

'Retrieve the name of the current ASP document
sPageURL = Request.ServerVariables("SCRIPT_NAME")

'Retrieve the current page number from the QueryString
intCurrentPage = Request.QueryString("page")
If intCurrentPage = "" Or intCurrentPage = 0 Then intCurrentPage = 1

'Set the number of records to be displayed on each page
iPageSize = 200

	strClientUserid = Request("userid")
	strCity = Trim(Request("city"))
	strCarType = Request("car_type")
	strCompany = Request("company")
	intDaysBack = Request("days_back")
	intDayToInclude = Request("days_to_include")
		
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
	
	If (strClientUserid = "") And (strCity = "") And (strCarType  = "") And (strCompany = "") And (strDaysBack = "") Then
	
		strConn = Session("pro_con")
	
  		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_request_select1"
		adoCmd.CommandType = 4

		'adoCmd.Parameters.Refresh 
		adoCmd.Parameters.Append adoCmd.CreateParameter("@client_userid",      200, 1, 20)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",            200, 1, 6)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds",  200, 1, 255)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vendor_cd",          200, 1, 255)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",              3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@days_to_include",      3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@linked_to_send_dttm", 11, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@day_back",             3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_role",            3, 1, 0)
		
		adoCmd.Parameters("@user_id").Value = Request.Cookies("rate-monitor.com")("user_id")
		adoCmd.Parameters("@client_userid").Value = Null
		adoCmd.Parameters("@city_cd").Value = Null
		adoCmd.Parameters("@shop_car_type_cds").Value = Null
		adoCmd.Parameters("@vendor_cd").Value = Null 
		adoCmd.Parameters("@days_to_include") = 0
		adoCmd.Parameters("@linked_to_send_dttm") = Null
		adoCmd.Parameters("@day_back") = intDaysBack 
		adoCmd.Parameters("@user_role").Value = Request.Cookies("rate-monitor.com")("user_role")
		
		'Create an ADO RecordSet object
		Set adoRS = Server.CreateObject("ADODB.Recordset")

		'Set the RecordSet PageSize property
		adoRS.PageSize = iPageSize 
		adoRS.CursorLocation = adUseClient 

		'Set the RecordSet CacheSize property to the
		'number of records that are returned on each page of results
		adoRS.CacheSize = iPageSize 

		'Open the RecordSet
		adoRS.Open adoCmd, , adOpenStatic, adLockReadOnly

		If adoRS.EOF = False Then
		
			adoRS.PageSize = iPageSize 
			intPageCount = adoRS.PageCount
			adoRS.AbsolutePage = intCurrentPage 

		Else
			intPageCount = 1
	
		End If


		'Set adoRS = adoCmd.Execute
	
		strSearched = True
 
	Else
	
		strConn = Session("pro_con")
	
  		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_request_select1"
		adoCmd.CommandType = 4

		'adoCmd.Parameters.Refresh 
		adoCmd.Parameters.Append adoCmd.CreateParameter("@client_userid",      200, 1, 20)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",            200, 1, 6)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds",  200, 1, 255)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vendor_cd",          200, 1, 255)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",              3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@days_to_include",      3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@linked_to_send_dttm", 11, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@day_back",             3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_role",            3, 1, 0)
		
		adoCmd.Parameters("@user_id").Value = Request.Cookies("rate-monitor.com")("user_id")

		If Trim(strClientUserid) <> "" Then
			adoCmd.Parameters("@client_userid").Value = strClientUserid 
		Else
			adoCmd.Parameters("@client_userid").Value = Null
		End If


		If Trim(strCity) <> "" Then
			adoCmd.Parameters("@city_cd").Value = strCity 
		Else
			adoCmd.Parameters("@city_cd").Value = Null
		End If


		If Trim(strCarType) <> "" Then
			adoCmd.Parameters("@shop_car_type_cds").Value = strCarType 
		Else
			adoCmd.Parameters("@shop_car_type_cds").Value = Null
		End If
	

		If Trim(strCompany) <> "" Then
			adoCmd.Parameters("@vendor_cd").Value = strCompany 
		Else
			adoCmd.Parameters("@vendor_cd").Value = Null 
		End If


		If IsNumeric(strDaysBack) = False Then
			strDaysBack = 0
		ElseIf strDaysBack < 0 Then
			strDaysBack = 0
		ElseIf strDaysBack > 5 Then
			strDaysBack = 5
		End If
		adoCmd.Parameters("@day_back").Value = strDaysBack 

		If IsNumeric(intDayToInclude) = False Then
			intDayToInclude = 0
		ElseIf intDayToInclude < 0 Then
			intDayToInclude = 0
		ElseIf intDayToInclude > 5 Then
			intDayToInclude = 5
		End If
		adoCmd.Parameters("@days_to_include").Value = intDayToInclude 

		adoCmd.Parameters("@user_role").Value = Request.Cookies("rate-monitor.com")("user_role")





		'Create an ADO RecordSet object
		Set adoRS = Server.CreateObject("ADODB.Recordset")

		'Set the RecordSet PageSize property
		adoRS.PageSize = iPageSize 
		adoRS.CursorLocation = adUseClient 

		'Set the RecordSet CacheSize property to the
		'number of records that are returned on each page of results
		adoRS.CacheSize = iPageSize 

		'Open the RecordSet
		adoRS.Open adoCmd, , adOpenStatic, adLockReadOnly

		If adoRS.EOF = False Then
		
			adoRS.PageSize = iPageSize 
			intPageCount = adoRS.PageCount
			adoRS.AbsolutePage = intCurrentPage 

		Else
			intPageCount = 1
	
		End If

		'Set adoRS = adoCmd.Execute
		'Set adoRS1 = adoCmd.Execute
	
		strSearched = True
		

	End If

	'If strClientUserId = "" Then
	'	strClientUserId = Request.Cookies("rate-monitor.com")("client_userid")
	'End If


	'If adoRS Is Nothing Then
	'	Set adoRS = CreateObject("ADODB.Recordset")
	'End If


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
<title>Rate-Monitor by Rate-Highway, Inc. | Rule Queue</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<% If intDaysBack > 0 Then %>
<% Else %>
<meta http-equiv="refresh" content="<%=intRefresh %>;url=search_queue_car.asp?userid=<%=Request("userid") %>">
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


	if (SubmitType == 'resetretrycount'){
		document.queue.action = 'resetretrycount_search_car.asp'
		//document.search_criteria.submit;
		//document.queue.submit;
		//return true;   
		}

	if (SubmitType == 'suspend'){
		document.queue.action = 'suspend_search_car.asp'
		//document.search_criteria.submit;
		//document.queue.submit;
		//return true;   
		}

	if (SubmitType == 'resume'){
		document.queue.action = 'resume_search_car.asp'
		//document.search_criteria.submit;
		//document.queue.submit;
		//return true;   
		}


}

</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">

<table border="0" cellspacing="0" cellpadding="0" style="width: 1710px">
  <tr>
    <td background="images/top_tile.gif">
    <img src="images/top.jpg" width="770" height="91"></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" style="width: 1710px">
  <tr>
    <td background="images/b_tile.gif">
    <!-- #INCLUDE FILE="inc/page_header_buttons.htm" -->
    </td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" style="width: 1710px">
  <tr>
    <td background="images/med_bar_tile.gif">
    <img src="images/med_bar.gif" width="12" height="8"></td>
  </tr>
</table>
<table width="1710" border="0" cellspacing="0" cellpadding="0">
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
<table border="0" cellspacing="0" cellpadding="0" style="width: 1710px">
  <tr>
    <td background="images/h_tile.gif">
    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/h_search_que.gif" width="368" height="31"></td>
        <td><map name="logout_map">
        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25">
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
&nbsp;
<form method="POST" action="search_queue_car.asp" name="search" class="search">
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
      <font size="2" >To search enter 
      the last name, or a portion of. You may also optionally enter city, car type 
      and/or the car company.</font></td>
      <td width="572" height="18">&nbsp;</td>
      <td width="571" height="18">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="26">&nbsp;</td>
      <td width="178"><img border="0" src="images/search.GIF"></td>
      <td width="137" height="26">
      <font size="2" ><label for="userid" >User Name:</label>
      </font></td>
      <td width="211" height="26">
      <input id="userid" type="text" name="userid" size="20" value="<%=strClientUserid %>" onfocus="this.className='focus';" onblur="this.className='';">
      </td>
      <td width="1207" height="26" colspan="3">
      <font size="2" >
      <input type="submit" value="  Display  " name="submit" class="rh_button"></font></td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="178">&nbsp;</td>
      <td width="137" height="22">
      <font size="2" >City:</font></td>
      <td width="211" height="22">
      <input type="text" name="city" size="20" value="<%=strCity %>" onfocus="this.className='focus';cl(this,'<%=strCity %>');" onblur="this.className='';fl(this,'<%=strstrCity %>');"></td>
      <td width="1207" height="22" colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="178">&nbsp;</td>
      <td width="137" height="22">
      <font size="2" >Car Type:</font></td>
      <td width="211" height="22">
      <input type="text" name="car_type" size="20" value="<%=strCarType %>" onfocus="this.className='focus';cl(this,'<%=strCarType %>');" onblur="this.className='';fl(this,'<%=strCarType %>');"></td>
      <td width="1207" height="22" colspan="3"></td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="178">&nbsp;</td>
      <td width="137" height="22">
      <font size="2" >Company:</font></td>
      <td width="211" height="22">
      <input type="text" name="company" size="20" value="<%=strCompany %>" onfocus="this.className='focus';cl(this,'<%=strCompany %>');" onblur="this.className='';fl(this,'<%=strCompany %>');"></td>
      <td width="1207" height="22" colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="18">&nbsp;</td>
      <td width="178">&nbsp;</td>
      <td width="137" height="18">
      <font size="2" >Days Back:</font></td>
      <td width="211" height="18">

      <select size="1" name="days_back" style="border:1px solid #000000; width:145; background-color:#FF9933">
		  <% Select Case intDaysBack %>
		  <%	Case 0 %>
          <option  value="0" selected>Today</option>
          <option  value="1" >Yesterday</option>
          <option  value="2" >2 days ago</option>
          <option  value="3" >3 days ago</option>
          <option  value="4" >4 days ago</option>
          <option  value="5" >5 days ago</option>
          <option  value="6" >6 days ago</option>
          <option  value="7" >7 days ago</option>
          <%	Case 1 %>
          <option  value="0" >Today</option>
          <option  value="1" selected>Yesterday</option>
          <option  value="2" >2 days ago</option>
          <option  value="3" >3 days ago</option>
          <option  value="4" >4 days ago</option>
          <option  value="5" >5 days ago</option>
          <option  value="6" >6 days ago</option>
          <option  value="7" >7 days ago</option>
		  <%	Case 2 %>
          <option  value="0" >Today</option>
          <option  value="1" >Yesterday</option>
          <option  value="2" selected>2 days ago</option>
          <option  value="3" >3 days ago</option>
          <option  value="4" >4 days ago</option>
          <option  value="5" >5 days ago</option>
          <option  value="6" >6 days ago</option>
          <option  value="7" >7 days ago</option>
		  <%	Case 3 %>
          <option  value="0" >Today</option>
          <option  value="1" >Yesterday</option>
          <option  value="2" >2 days ago</option>
          <option  value="3" selected>3 days ago</option>
          <option  value="4" >4 days ago</option>
          <option  value="5" >5 days ago</option>
          <option  value="6" >6 days ago</option>
          <option  value="7" >7 days ago</option>
		  <%	Case 4 %>
          <option  value="0" >Today</option>
          <option  value="1" >Yesterday</option>
          <option  value="2" >2 days ago</option>
          <option  value="3" >3 days ago</option>
          <option  value="4" selected>4 days ago</option>
          <option  value="5" >5 days ago</option>
          <option  value="6" >6 days ago</option>
          <option  value="7" >7 days ago</option>
		  <%	Case 5 %>
          <option  value="0" >Today</option>
          <option  value="1" >Yesterday</option>
          <option  value="2" >2 days ago</option>
          <option  value="3" >3 days ago</option>
          <option  value="4" >4 days ago</option>
          <option  value="5" selected>5 days ago</option>
          <option  value="6" >6 days ago</option>
          <option  value="7" >7 days ago</option>
		  <%	Case 6 %>
          <option  value="0" >Today</option>
          <option  value="1" >Yesterday</option>
          <option  value="2" >2 days ago</option>
          <option  value="3" >3 days ago</option>
          <option  value="4" >4 days ago</option>
          <option  value="5" >5 days ago</option>
          <option  value="6" selected>6 days ago</option>
          <option  value="7" >7 days ago</option>
		  <%	Case 7 %>
          <option  value="0" >Today</option>
          <option  value="1" >Yesterday</option>
          <option  value="2" >2 days ago</option>
          <option  value="3" >3 days ago</option>
          <option  value="4" >4 days ago</option>
          <option  value="5" >5 days ago</option>
          <option  value="6" >6 days ago</option>
          <option  value="7" selected>7 days ago</option>
          <%	Case Else %>
          <option  value="0" selected>Today</option>
          <option  value="1" >Yesterday</option>
          <option  value="2" >2 days ago</option>
          <option  value="3" >3 days ago</option>
          <option  value="4" >4 days ago</option>
          <option  value="5" >5 days ago</option>
          <option  value="6" >6 days ago</option>
          <option  value="7" >7 days ago</option>
          
          <% End Select %>
          </select></font></td>
      <td width="1207" height="18" colspan="3">&nbsp;</td>
    </tr>
<!--    
    <tr>
      <td width="19" height="18">&nbsp;</td>
      <td width="178">&nbsp;</td>
      <td width="137" height="18">
      <font size="2" >Days to 
		include:</font></td>
      <td width="211" height="18">

      <select size="1" name="days_to_include" style="border:1px solid #000000; width:145; background-color:#FF9933">
		  <% Select Case intDayToInclude %>
		  <%	Case 1 %>
          <option  value="1" selected>1 day</option>
          <option  value="2" >2 days/option>
          <option  value="3" >3 days</option>
          <option  value="4" >4 days</option>
          <option  value="5" >5 days</option>
		  <%	Case 2 %>
          <option  value="1" >1 day</option>
          <option  value="2" selected>2 days</option>
          <option  value="3" >3 days</option>
          <option  value="4" >4 days</option>
          <option  value="5" >5 days</option>
		  <%	Case 3 %>
          <option  value="1" >1 day</option>
          <option  value="2" >2 days</option>
          <option  value="3" selected>3 days</option>
          <option  value="4" >4 days</option>
          <option  value="5" >5 days</option>
		  <%	Case 4 %>
          <option  value="1" >1 day</option>
          <option  value="2" >2 days</option>
          <option  value="3" >3 days</option>
          <option  value="4" selected>4 days</option>
          <option  value="5" >5 days</option>
		  <%	Case 5 %>
          <option  value="1" >1 day</option>
          <option  value="2" >2 days</option>
          <option  value="3" >3 days</option>
          <option  value="4" >4 days</option>
          <option  value="5" selected>5 days</option>
          <%	Case Else %>
          <option  value="1" selected>1 day</option>
          <option  value="2" >2 days</option>
          <option  value="3" >3 days</option>
          <option  value="4" >4 days</option>
          <option  value="5" >5 days</option>
          
          <% End Select %>
          </select></font></td>
      <td width="1207" height="18" colspan="3">&nbsp;</td>
    </tr>
-->
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
    <td width="797"><font size="2" >
    &nbsp;&nbsp;| <a href="javascript:confirmSubmit('cancel');document.queue.submit();" >Cancel</a> | <a href="javascript:confirmSubmit('suspend');document.queue.submit();" >Suspend</a> | <a href="javascript:confirmSubmit('resume');document.queue.submit();" >Resume</a> 
<% If Request.Cookies("rate-monitor.com")("user_id") = 3 Then %>
| Advanced options listed below data grid 
<% End If %>
| <a href="search_queue_car_ex.asp">Extended view</a> | Standard view |</font></td><td >
    <font size="2" >

    <% If intCurrentPage <= 1 Then %>
	    |&lt;
	    &lt; 
	<% Else  %>
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&company=<%=strCompany%>&page=1">|&lt;</a>
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&company=<%=strCompany%>&page=<%=intCurrentPage - 1%>">&lt;</a> 
    <% End If %>
    Page <%=intCurrentPage %> of <%=intPageCount %>
    <% If CInt(intCurrentPage) >= CInt(intPageCount) Then %>
	    &gt;
	    &gt;|
    <% Else %>
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&company=<%=strCompany%>&page=<%=intCurrentPage + 1%>">&gt;</a>
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&company=<%=strCompany%>&page=<%=intPageCount %>">&gt;|</a>
    <% End If %>

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
    <td class="profile_header" width="50" height="45">Search Status</td>
    <td class="profile_header" width="75" height="45">Request<br>Time</td>
<% If Request.Cookies("rate-monitor.com")("user_id") = 3 Then %>
    <td class="profile_header" width="75" height="45">Emailed<br>(pst)<br>
    rts|alts</td>
<% End If %>    
    <td class="profile_header" width="75" height="45">User Name</td>
    <td class="profile_header" width="300" height="45">Profile<br>(hover over to view complete name)</td>
    <td class="profile_header" width="60" height="45">Source</td>
    <td class="profile_header" width="73" height="45">Rates Expected</td>
    <td class="profile_header" width="79" height="45">Rates Complete</td>
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

		If (strSearched = True) And (adoRS.State = adStateOpen) Then

		While (adoRS.EOF = False) And (intCount < iPageSize)
		
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
    <% If (adoRS.Fields("request_status").Value = "C") Then %>
    <input type="checkbox" value='<%=adoRS.Fields("shop_request_id").Value %>' name="shop_request_id" disabled></td>
    <% Else %>
    <input type="checkbox" value='<%=adoRS.Fields("shop_request_id").Value %>' name="shop_request_id"></td>
    <% End If %>
    
    
    <td class="<%=strClass %>" height="20" width="45" title="click report number to view in new window">
    <a href="car_report_by_type.asp?ReportRequestID=<%=adoRS.Fields("shop_request_id").Value %>&security_code=<%=Escape(adoRS.Fields("security_code").Value) %>" target="_blank">
    <%=adoRS.Fields("shop_request_id").Value %></a></td>
<!--
    <a href="view_report_car2.asp?ReportRequestID=<%=adoRS.Fields("shop_request_id").Value %>" target="_blank">
    <%=adoRS.Fields("shop_request_id").Value %></a></td>
-->
    <td class="<%=strClass %>" height="20" width="50">
    <% Select Case adoRS.Fields("request_status").Value %>
	<%     Case "R" %>
	Running
	<%     Case "C" %>
	Cancelled
	<%     Case "P" %>
	Suspended
	<%     Case "N" %>
	New
	<%     Case "S" %>
	Successful
	<%     Case "F" %>
	Failure
	<%     Case Else %>
	<%=adoRS.Fields("request_status").Value %>
	<% End Select %></td>
    <td class="<%=strClass %>" height="20" width="75" align="center">
    <% If DateDiff("d", Now, adoRS.Fields("scheduled_dttm").Value) = 0 Then %>
    
    	<% 
    	
    	strTime = ""
        	
		If (Hour(adoRS.Fields("scheduled_dttm").Value) = 0) Then 
    		strTime = "12:"

    	ElseIf (Len(Hour(adoRS.Fields("scheduled_dttm").Value)) = 1) Then
    		strTime = "0" & Hour(adoRS.Fields("scheduled_dttm").Value ) & ":"

		ElseIf (Hour(adoRS.Fields("scheduled_dttm").Value) > 12) Then 
			If (Len(Hour(adoRS.Fields("scheduled_dttm").Value ) - 12) = 1) Then	
	    		strTime = "0" & Hour(adoRS.Fields("scheduled_dttm").Value ) - 12 & ":"
			
			Else
	    		strTime = Hour(adoRS.Fields("scheduled_dttm").Value ) - 12 & ":"
			
			End If

		Else
    		strTime = Hour(adoRS.Fields("scheduled_dttm").Value ) & ":"
   	
    	End If

    	'strTime = strTime & Hour(adoRS.Fields("scheduled_dttm").Value ) & ":"
    
    	If Len(Minute(adoRS.Fields("scheduled_dttm").Value)) = 1 Then
    		strTime = strTime & "0"  
    		
    	End If
    	
    	strTime = strTime & Minute(adoRS.Fields("scheduled_dttm").Value) '& ":"

    	'If Len(Second(adoRS.Fields("scheduled_dttm").Value)) = 1 Then
    	'	strTime = strTime & "0"  
    	'	
    	'End If
    	
    	'strTime = strTime & Second(adoRS.Fields("scheduled_dttm").Value)
    	
    	
    	If  Hour(adoRS.Fields("scheduled_dttm").Value) > 11 Then
    		strTime = strTime & " pm"
    	
    	Else
    		strTime = strTime & " am"
    	
    	
    	End If
   	
    	%>	
    	
    	<%=strTime %>
    	
    	<!--
      <%=DatePart("h", adoRS.Fields("scheduled_dttm").Value) & ":" & DatePart("n", adoRS.Fields("scheduled_dttm").Value) & ":" & DatePart("s", adoRS.Fields("scheduled_dttm").Value) %>
    	-->
    	
    <% Else %>
    
	    <% If Len(Month(adoRS.Fields("scheduled_dttm").Value)) = 1 Then
	    	strTime = "0" & Month(adoRS.Fields("scheduled_dttm").Value) & "/"
	       Else
	    	strTime = Month(adoRS.Fields("scheduled_dttm").Value) & "/"
	       End If
       
		   If Len(Day(adoRS.Fields("scheduled_dttm").Value)) = 1 Then
	    	strTime = strTime & "0" & Day(adoRS.Fields("scheduled_dttm").Value) & "/"
	       Else
	    	strTime = strTime & Day(adoRS.Fields("scheduled_dttm").Value) & "/"
	       End If
       
	 	   strTime = strTime & "0" & Year(adoRS.Fields("scheduled_dttm").Value) - 2000

		%>
		<%=strTime %>
    <!--      
    
	  <%=FormatDateTime(adoRS.Fields("scheduled_dttm").Value, 2) %>
	-->
    <% End If %>
    
    </td>
<% If Request.Cookies("rate-monitor.com")("user_id") = 3 Then %>
	<td class="<%=strClass %>" height="20" width="75" align="center">
	
	<% If IsNull(adoRS.Fields("email_dttm").Value) And IsNull(adoRS.Fields("alert_dttm").Value) Then %>
		<% strTime = "Not sent" %>
	
	<% Else %>

    	<% 
    	
    	strTime = ""

IF False Then
        	
		If (Hour(adoRS.Fields("email_dttm").Value) = 0) Then 
    		strTime = "12:"

    	ElseIf (Len(Hour(adoRS.Fields("email_dttm").Value)) = 1) Then
    		strTime = "0" & Hour(adoRS.Fields("email_dttm").Value ) & ":"

		ElseIf (Hour(adoRS.Fields("email_dttm").Value) > 12) Then 
			If (Len(Hour(adoRS.Fields("email_dttm").Value ) - 12) = 1) Then	
	    		strTime = "0" & Hour(adoRS.Fields("email_dttm").Value ) - 12 & ":"
			
			Else
	    		strTime = Hour(adoRS.Fields("email_dttm").Value ) - 12 & ":"
			
			End If

		Else
    		strTime = Hour(adoRS.Fields("email_dttm").Value ) & ":"
   	
    	End If

   
    	If Len(Minute(adoRS.Fields("email_dttm").Value)) = 1 Then
    		strTime = strTime & "0"  
    		
    	End If
    	
    	strTime = strTime & Minute(adoRS.Fields("email_dttm").Value) '& ":"

    	If  Hour(adoRS.Fields("email_dttm").Value) > 11 Then
    		strTime = strTime & " pm"
    	
    	Else
    		strTime = strTime & " am"
    	
    	
    	End If

End If

	If IsNull(adoRS.Fields("email_dttm").Value) Then
		strTime = "n/a|"
	Else	
		strTime = Hour(adoRS.Fields("email_dttm").Value) & ":" 

		If Len(Minute(adoRS.Fields("email_dttm").Value)) = 1 Then
    		strTime = strTime & "0"  
    	End If

		strTime = strTime & Minute(adoRS.Fields("email_dttm").Value) & "|"

	End If

	If IsNull(adoRS.Fields("alert_dttm").Value) Then 
		strTime = strTime  & "n/a"
	Else	
		strTime = strTime & Hour(adoRS.Fields("alert_dttm").Value) & ":" 
		
		If Len(Minute(adoRS.Fields("alert_dttm").Value)) = 1 Then
    		strTime = strTime & "0"  
    	End If

		strTime = strTime & Minute(adoRS.Fields("alert_dttm").Value)		
		
	End If

   	
    	%>	
    	
    <% End If %>
    
    <%=strTime %>
   	
	</td>
<% End If %>

    <td class="<%=strClass %>" height="20" width="75"  title="<%="Requesting user was: " & adoRS.Fields("client_userid").Value %>">
		<% If Len(adoRS.Fields("client_userid").Value) > 9 Then %>
		  <%=Left(adoRS.Fields("client_userid").Value, 7 ) & "..." %>
		<% Else %>
		  <%=adoRS.Fields("client_userid").Value %>
		<% End If %>
     </td>
    <td class="<%=strClass %>" height="20" width="300" title="<%=adoRS.Fields("profile_desc").Value %>">
	<% If adoRS.Fields("profile_desc").Value = "" Then %>
	 <i>[none]</i> On-demand report request
	<% Else %>
		<% If Len(adoRS.Fields("profile_desc").Value) > 35 Then %>
		  <%=Left(adoRS.Fields("profile_desc").Value, 35) & "..." %>
		<% Else %>
		  <%=adoRS.Fields("profile_desc").Value %>
		<% End If %>
	<% End If %>
	</td>
    <td class="<%=strClass %>" height="20" title="<%=adoRS.Fields("data_sources").Value %>">
    <% Select Case adoRS.Fields("data_sources").Value
    	Case "WSA"
    %>
	    <%="GDS-A" %>
	<%
		Case "EXK"
	%>
		<%="EXP Pkg" %>
    <%
    	Case "SA2", "WSG"
    %>
	    <%="GDS Gov" %>
    <%
    	Case "BRD"
    %>
	    <%="Brand" %>

	<%
    	Case "UKB"
    %>
	    <%="UK Brnd" %>

	<%
	    Case "TDT", "SA1", "WSR"
	%>    
		<%="GDS" %>
    <%	
    	Case Else
   	
   	
    		If Left(adoRS.Fields("data_sources").Value, 1) = "V" Then
    %>
		    <%="Brand"  %> <% Rem Left(adoRS.Fields("data_sources").Value, 4) & "..." %>
    <%	

    		ElseIf InStr(adoRS.Fields("data_sources").Value, ",") Then
    %>
		    <%="Multiple"  %>
    <%	
    		Else
    %>
		    <%=adoRS.Fields("data_sources").Value %>
    <%	
    		End If
    		
    	End Select
   	%>
    </td>
    <td class="<%=strClass %>" height="20"><%=adoRS.Fields("work_units").Value %></td>
    <td class="<%=strClass %>" height="20"><%=adoRS.Fields("work_units_complete").Value %></td>
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
		Set adoRS1 = Nothing
		Set adoCmd = Nothing

		Else

		%>
		
		
  <tr>
    <td class="profile_light" height="20"></td>
    <td bgcolor="#FDC677" align="center" height="20">
    <input type="radio" value="V1" name="selected"></td>
    <td class="profile_light" height="20" width="45"></td>
    <td class="profile_light" height="20" width="50"></td>
    <td class="profile_light" height="20" width="75">&nbsp;</td>
    <td class="profile_light" height="20" width="75"></td>
    <td class="profile_light" height="20" width="300"></td>
    <td class="profile_light" height="20"></td>
    <td class="profile_light" height="20"></td>
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
<table width="1710" border="0" cellpadding="2" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
  <tr valign="bottom">
<td width="797">
<p>&nbsp;<font size="2">| <a href="javascript:confirmSubmit('cancel');document.queue.submit();" >Cancel</a> 
</font><font size="2" >
| <a href="javascript:confirmSubmit('suspend');document.queue.submit();" >Suspend</a> | <a href="javascript:confirmSubmit('resume');document.queue.submit();" >Resume</a></font><font size="2"> 
<% Select Case Request.Cookies("rate-monitor.com")("user_id")

	Case 3, 70
		%>
		| <a href="javascript:confirmSubmit('redo');document.queue.submit();" >Redo</a> | <a href="javascript:confirmSubmit('redoremail');document.queue.submit();">
		Redo &amp; Remail</a> | <a href="javascript:confirmSubmit('forcecomplete');document.queue.submit();">Force Complete</a> | <a href="javascript:confirmSubmit('resetretrycount');document.queue.submit();">Reset Retry Count</a>    
		<%
	Case 161, 102, 103, 104, 105, 106, 108, 237, 214
		%>
		| <a href="javascript:confirmSubmit('forcecomplete');document.queue.submit();">Force Complete</a>    
		<%

 	Case Else

  End Select %>
| <a href="search_queue_car_ex.asp">Extended view</a> | Standard view |</font></p>
</td>
<td >
    <font size="2" >

    <% If intCurrentPage <= 1 Then %>
	    |&lt;
	    &lt; 
	<% Else  %>
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&company=<%=strCompany%>&page=1">|&lt;</a>
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&company=<%=strCompany%>&page=<%=intCurrentPage - 1%>">&lt;</a> 
    <% End If %>
    Page <%=intCurrentPage %> of <%=intPageCount %>
    <% If CInt(intCurrentPage) >= CInt(intPageCount) Then %>
	    &gt;
	    &gt;|
    <% Else %>
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&company=<%=strCompany%>&page=<%=intCurrentPage + 1%>">&gt;</a>
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&company=<%=strCompany%>&page=<%=intPageCount %>">&gt;|</a>
    <% End If %>

    </font></td>

</tr> 
</table>
</form>
    <p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">© 
2002 - 2013 - All rights reserved<br>
<b>Rate-Highway, Inc.</b><br>
</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">18001 Cowan, 
    Suite F<br />
    Irvine, CA 92614<br />
    (949) 614-0751&nbsp; 
        <br />
        <br />
        </font>

        <font size="1">
        <a target="_blank" title="Click to open a new window and browse the support documents" href="https://na5.salesforce.com/sserv/login.jsp?orgId=00D70000000JFA6">
Support Center</a>
<!-- 
<a target="_blank" title="Click to open a new window and browse the support documents" href="support_files.asp">
Support Documents</a>
-->
</font></p>
<% If intDaysBack > 0 Then %>
<font size="2" >
<p align="center"><blink><i>You are viewing a prior day's queue, this page will not 
auto-refresh</i></blink></p>
</font>
<% End If %>
<font size="1" >
<p align="left"><%=Session("server_name") %></p>
</font>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
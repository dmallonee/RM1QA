<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<% Response.Expires = -1  
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache" 
   
   	on error resume next

   	Server.ScriptTimeout = 30

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	strCityCd = Request("city_cd") & ""
	strConn   = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "user_city_select"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",  3, 1,  0, strUserId)
		
	Set adoRSCities = adoCmd.Execute
	
	If strCityCd = "" Then
		strCityCd = adoRSCities.fields("city_cd").value
	End If
	

	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "car_type_select"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",  3, 1,  0, strUserId)
		
	Set adoRSCars = adoCmd.Execute

	If Request("delete") = "true" Then
		Set adoCmd = CreateObject("ADODB.Command")
		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "return_charge_delete"
		adoCmd.CommandType = adCmdStoredProc
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@return_chrg_id",      3, 1, 0, Request("return_chrg_id"))
			
		adoCmd.Execute
		
	End If
	
	If Request("rate_cd") <> "" Then
	
		Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "return_charge_insert"
		adoCmd.CommandType = adCmdStoredProc
		
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",            3, 1,  0, strUserId)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",          200, 1,  6, Request("city_cd"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rate_cd",          200, 1, 20, Request("rate_cd"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd",      200, 1,  4, Request("car_type_cd"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@begin_dt",         135, 1,  0, Request("begin_dt"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@end_dt",           135, 1,  0, Request("end_dt"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@early_amt",          6, 1,  0, Request("early_amt"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@early_is_dollar",   11, 1,  0, Request("early_is_dollar"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@late_amt",           6, 1,  0, Request("late_amt"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@late_is_dollar",    11, 1,  0, Request("late_is_dollar"))
			
		Set adoRS = adoCmd.Execute

	
	End If
	
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "return_charge_select"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",      3, 1, 0, strUserId)
				
	Set adoRS = adoCmd.Execute

	
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write pad & "</b><br>"
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
<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; Early &amp; Late Charge Settings</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script language="JavaScript" type="text/JavaScript" src="inc/sitewide.js" ></script>
<script language="JavaScript" type="text/JavaScript" src="inc/pupdate.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}


function openWindow(theURL,winName,features) 
{ //v2.0
  window.open(theURL,winName,features);
}

function IsDate(dateStr) {

var datePat = /^(\d{1,2})(\/|-)(\d{1,2})(\/|-)(\d{4})$/;
var matchArray = dateStr.match(datePat); // is the format ok?

if (matchArray == null) {
alert("Please enter date as either mm/dd/yyyy or mm-dd-yyyy.");
return false;
}

month = matchArray[1]; // p@rse date into variables
day = matchArray[3];
year = matchArray[5];

if (month < 1 || month > 12) { // check month range
alert("Month must be between 1 and 12.");
return false;
}

if (day < 1 || day > 31) {
alert("Day must be between 1 and 31.");
return false;
}

if ((month==4 || month==6 || month==9 || month==11) && day==31) {
alert("Month "+month+" doesn`t have 31 days!")
return false;
}

if (month == 2) { // check for february 29th
var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
if (day > 29 || (day==29 && !isleap)) {
alert("February " + year + " doesn`t have " + day + " days!");
return false;
}
}
return true; // date is valid
}


function ValidateForm(form)
{
   if(IsEmpty(form.rate_cd)) 
   { 
      alert('Please enter a rate code') 
      fee_grid.rate_cd.focus(); 
      return false; 
   } 
 

   if(!IsDate(form.begin_dt)) 
   { 
      alert('Please enter a begin date') 
      fee_grid.begin_dt.focus(); 
      return false; 
   } 

   if(!IsDate(form.end_dt)) 
   { 
      alert('Please enter an end date') 
      fee_grid.end_dt.focus(); 
      return false; 
   } 
  
   if (!IsNumeric(form.early_amt.value)) 
   { 
      alert('Please enter only numbers or decimal points in the early amount field') 
      form.early_amt.focus(); 
      return false; 
   } 


   if (!IsNumeric(form.late_amt.value)) 
   { 
      alert('Please enter only numbers or decimal points in the late amount field') 
      form.late_amt.focus(); 
      return false; 
   } 

 
return true;
 
} 

//-->
</script>

<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<style type="text/css">
<!--
.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.style5 {
	text-align: center;
	font-size: medium;
}
.style7 {
	border-collapse: collapse;
}
.style11 {
	text-align: center;
}
.style12 {
	font-size: xx-small;
}
.style13 {
	text-align: right;
}
.style14 {
	font-size: xx-small;
	text-align: center;
}
-->
</style>
<base target="_self">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')" >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif"><img src="images/top.jpg" width="770" height="91"></td>
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
    <td background="images/med_bar_tile.gif"><img src="images/med_bar.gif" width="12" height="8"></td>
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
                      <td><div align="right">
                  <font face="Vendana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">
                  User: <%=Request.Cookies("rate-monitor.com")("user_name")%></font></div></td>
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
    <table width="100" border="0" cellspacing="0" cellpadding="0" id="table1">
      <tr>
        <td><img src="images/h_system.gif" width="368" height="31"></td>
        <td><map name="logout_map">
        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25">
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<p class="style11">&nbsp;&nbsp;&nbsp; <br>
</p>
  <table border="0" style="width: 800px;" bordercolor="#FFFFFF"  align="center" class="style7">
    
    <tr>
      <td >&nbsp;</td>
     
    </tr>
    <tr>
      <td background="images/ruler.gif" height="4"></td>
     
    </tr>
  </table>
<p>&nbsp;<!-- 
	<p align="center">
		Peter - use this report for right now please =&gt;
	<a href="system_utilization_report.asp">utilization report</a></p>
	
	-->
	</p>
	<form method="get" name="fee_grid"  onsubmit="javascript:return ValidateForm(this)">
<p class="style5">Early / Late Return Charge Settings<br>
	<div align="center">
        <table border="0" cellpadding="0" style="width: 750;" bordercolor="#111111" class="style7">
          <tr>
           <td width="100%" class="boxtitle" colspan="8" style="height: 15px"><font size="2"><br><b>
           Directions:</b> Please enter a value for each car type that you would 
		   like to have a charge calculated. Positive or negative numbers may be 
		   entered. If you no longer wish to have a late charge for a car type, 
		   simply enter a zero or blank value and it will be ignored when your 
		   late charge is calculated.</font><p>
           &nbsp;</p>
           </td>
           
          </tr>
          <tr>
           	<td class="boxtitle">&nbsp;</td>
           	<td class="boxtitle">&nbsp;</td>
			<td class="style13"><font size="2">Location:&nbsp;&nbsp;</font></td>
            <td class="boxtitle">
			<select name="city_cd">
				  
                   <%   While (adoRSCities.EOF = False) 
 		                  If adoRSCities.Fields("city_cd").Value = strCityCd Then %>
		                    <option selected ><%=adoRSCities.Fields("city_cd").Value %></option>		           
		           <%     Else %>	 
		                    <option ><%=adoRSCities.Fields("city_cd").Value %></option>
		           <%     End If %>
		 		   <%     adoRSCities.MoveNext %>
		           <%   Wend %>

			</select></td>
            <td class="boxtitle">
			<input class="rh_button" name="display_city" type="submit" value="Display"></td>
            <td class="boxtitle">&nbsp;</td>
            <td class="boxtitle">&nbsp;</td>
            <td class="boxtitle">&nbsp;</td>
            
		  </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
          </tr>
          <tr class="profile_header">
            <td  class="boxtitle" style="height: 11px"><font size="2"><u>Location</u></font></td>
            <td  class="boxtitle" style="height: 11px"><font size="2"><u>Rate Code.</u></font></td>
            <td  class="boxtitle" style="height: 11px"><font size="2"><u>Car Type</u></font></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Begin</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">End</font></u></td>
            <td  class="boxtitle" style="height: 11px"><font size="2"><u>Early Charge</u></font></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Late Charge</font></u></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
          </tr>
          <% 	Dim intCount             	%>
          <% 	intCount = 0	            %>
          <% 	strClass = "profile_dark"	%>
          
          <% If adoRS.State = adStateOpen Then %>
          <% While (adoRS.EOF = False) %>
          <%   If strClass = "profile_dark" Then
          	     strClass = "profile_light"
          	   Else
          	     strClass = "profile_dark"
          	   End If
          %>
          <tr  class="<%=strClass %>" >
            <td class="boxtitle" ><font size="2"><%=adoRS.Fields("city_cd").Value %></font></td>
            <td class="boxtitle" ><font size="2"><%=adoRS.Fields("rate_cd").Value %></font></td>
            <td class="boxtitle" ><font size="2"><%=adoRS.Fields("car_type_cd").Value %></font></td>
            <td class="boxtitle" ><font size="2"><%=adoRS.Fields("begin_dt").Value %></font></td>
            <td class="boxtitle" ><font size="2"><%=adoRS.Fields("end_dt").Value %></font></td>
            <td class="style13"  ><font size="2">
            <% If CBool(adoRS.Fields("early_is_dollar").Value) Then %>
            <%=FormatCurrency(adoRS.Fields("early_amt").Value) %>
            <% Else %>
            <%=FormatPercent(adoRS.Fields("early_amt").Value / 100) %>
            <% End If %>
            </font></td>
            <td class="style13" ><font size="2">
            <% If CBool(adoRS.Fields("late_is_dollar").Value) Then %>
            <%=FormatCurrency(adoRS.Fields("late_amt").Value) %>
            <% Else %>
            <%=FormatPercent(adoRS.Fields("late_amt").Value / 100) %>
            <% End If %>
            </font></td>
            <td class="style14" ><a  href="early_late_charge_maint.asp?delete=true&return_chrg_id=<%=adoRS.Fields("return_chrg_id").Value %>">&nbsp;&nbsp;delete</a></td>
          </tr>
          <% 	intCount = intCount + 1	%>
          <%   adoRS.MoveNext         	%>
          <% Wend                     	%>
          <% End If 					%>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
          </tr>
          <tr >
            <td class="boxtitle" ><font size="2"><%=strCityCd%></font><input name="city_cd_also" type="hidden" size="5" value="<%=strCityCd%>"></td>
            <td class="boxtitle" ><input name="rate_cd" type="text" size="10"></td>
            <td class="boxtitle" ><select name="car_type_cd">
				   <%   While (adoRSCars.EOF = False) 
 		                  If adoRSCars.Fields("car_type_cd").Value = strCityCd Then %>
		                    <option selected ><%=adoRSCars.Fields("car_type_cd").Value %></option>		           
		           <%     Else %>	 
		                    <option ><%=adoRSCars.Fields("car_type_cd").Value %></option>
		           <%     End If %>
		 		   <%     adoRSCars.MoveNext %>
		           <%   Wend %>
			</select></td>
            <td class="boxtitle" >
			<input name="begin_dt" type="text" size="8"><img src="images/cal_button.gif" class="DatePicker" alt="Pick a begin date" height="20" width="32" onClick="getCalendarFor(document.fee_grid.begin_dt);return false" ></td>
            <td class="boxtitle" >
			<input name="end_dt" type="text" size="8"><img src="images/cal_button.gif" class="DatePicker" alt="Pick an end date" height="20" width="32" onClick="getCalendarFor(document.fee_grid.end_dt);return false" ></td>
            <td class="style13" >
			<input name="early_amt" type="text" size="5">
			<select name="early_is_dollar">
			<option value="True">$</option>
			<option value="False">%</option>
			</select>
			</td>
            <td class="style13" >
			<input name="late_amt" type="text" size="5">
			<select name="late_is_dollar">
			<option value="True">$</option>
			<option value="False">%</option>
			</select>
			</td>
            <td class="style12" >&nbsp;</td>
          </tr>
          </table>
		<input name="Add" type="submit" value="Add Charge"><br>
		<br>
		Total: <%=intCount %>
        </div>
        <p class="style11">&nbsp;</p>
<p class="style11">&nbsp; 
          </p>
  </FORM>

  <table border="0" style="width: 800px;" bordercolor="#FFFFFF" height="4" id="table1" align="center" class="style7">
    <tr>
      <td background="images/ruler.gif"></td>
      
    </tr>
  </table>
<p align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">© 
2002 - 2013 - All rights reserved<br>
<b>Rate-Highway, Inc.</b><br>
</font>18001 Cowan, 
    Suite F<br />
    Irvine, CA 92614<br />
    (949) 614-0751&nbsp;&nbsp; </font>
<p align="center">&nbsp;</p>
<p class="style12">u: <%=strUserId %></p>
<script language="JavaScript"type="text/JavaScript">
<!--
if (document.all) {
 document.writeln("<div id=\"PopUpCalendar\" style=\"position:absolute; left:0px; top:0px; z-index:7; width:200px; height:77px; overflow: visible; visibility: hidden; background-color: #FFFFFF; border: 1px none #000000\" onMouseOver=\"if(ppcTI){clearTimeout(ppcTI);ppcTI=false;}\" onMouseOut=\"ppcTI=setTimeout(\'hideCalendar()\',500)\">");
 document.writeln("<div id=\"monthSelector\" style=\"position:absolute; left:0px; top:0px; z-index:9; width:181px; height:27px; overflow: visible; visibility:inherit\">");}
else if (document.layers) {
 document.writeln("<layer id=\"PopUpCalendar\" pagex=\"0\" pagey=\"0\" width=\"200\" height=\"200\" z-index=\"100\" visibility=\"hide\" bgcolor=\"#FFFFFF\" onMouseOver=\"if(ppcTI){clearTimeout(ppcTI);ppcTI=false;}\" onMouseOut=\"ppcTI=setTimeout('hideCalendar()',500)\">");
 document.writeln("<layer id=\"monthSelector\" left=\"0\" top=\"0\" width=\"181\" height=\"27\" z-index=\"9\" visibility=\"inherit\">");}
else {
 document.writeln("<p><font color=\"#FF0000\"><b>Error ! The current browser is either too old or too modern (usind DOM document structure).</b></font></p>");}
 -->
</script>
<noscript><p><font color="#FF0000"><b>JavaScript is not activated !</b></font></p></noscript>
<table border="1" cellspacing="1" cellpadding="2" width="200" bordercolorlight="#000000" bordercolordark="#000000" vspace="0" hspace="0"><form name="ppcMonthList"><tr><td align="center" bgcolor="#CCCCCC"><a href="javascript:moveMonth('Back')" onMouseOver="window.status=' ';return true;"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000"><b> </b></font></a><font face="MS Sans Serif, sans-serif" size="1"> 
<select name="sItem" onMouseOut="if(ppcIE){window.event.cancelBubble = true;}" onChange="switchMonth(this.options[this.selectedIndex].value)" style="font-family: 'MS Sans Serif', sans-serif; font-size: 9pt"><option value="0" selected>2000 • January</option><option value="1">2000 • February</option><option value="2">2000 • March</option><option value="3">2000 • April</option><option value="4">2000 • May</option><option value="5">2000 • June</option><option value="6">2000 • July</option><option value="7">2000 • August</option><option value="8">2000 • September</option><option value="9">2000 • October</option><option value="10">2000 • November</option><option value="11">2000 • December</option><option value="0">2001 • January</option></select></font><a href="javascript:moveMonth('Forward')" onMouseOver="window.status=' ';return true;"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000"><b></b></font></a></td></tr></form></table>
<table border="1" cellspacing="1" cellpadding="2" bordercolorlight="#000000" bordercolordark="#000000" width="200" vspace="0" hspace="0"><tr align="center" bgcolor="#CCCCCC"><td width="20" bgcolor="#FFFFCC"><b><font face="MS Sans Serif, sans-serif" size="1">Su</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Mo</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Tu</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">We</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Th</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Fr</font></b></td><td width="20" bgcolor="#FFFFCC"><b><font face="MS Sans Serif, sans-serif" size="1">Sa</font></b></td></tr></table>
<script language="JavaScript" type="text/JavaScript">
<!--
if (document.all) {
 document.writeln("</div>");
 document.writeln("<div id=\"monthDays\" style=\"position:absolute; left:0px; top:52px; z-index:8; width:200px; height:17px; overflow: visible; visibility:inherit; background-color: #FFFFFF; border: 1px none #000000\"> </div></div>");}
else if (document.layers) {
 document.writeln("</layer>");
 document.writeln("<layer id=\"monthDays\" left=\"0\" top=\"52\" width=\"200\" height=\"17\" z-index=\"8\" bgcolor=\"#FFFFFF\" visibility=\"inherit\"> </layer></layer>");}
else {/*NOP*/}
-->
</script>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
<% Set adoRS = Nothing
   Set adoCmd = Nothing 
%>
<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<% Response.Expires = -1  
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache" 
   
   	on error resume next

   	Server.ScriptTimeout = 180

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_utilization_out_detail"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@day",       135, 1, 0, Request("date"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@org_id",      3, 1, 0, Request("org_id"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@car_class", 200, 1, 4, Request("car_type_cd"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",   200, 1, 6, Request("city_cd"))
		
	Set adoRS = adoCmd.Execute


	'Set adoRS = CreateObject("ADODB.Recordset")


	
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

<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; Utilization Settings - 
Reservation Detail</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
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
//-->
</script>

<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<style>
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
.style10 {
	text-align: right;
	border-style: solid;
	border-width: 0;
}
.style11 {
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
<p>&nbsp;&nbsp;&nbsp; <br>
&nbsp;<font size="2" face="Vendana, Arial, Helvetica, sans-serif">&nbsp;<b>[utilization settings 
- open contract detail] </b></font><br>
&nbsp;</p>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1114" cellspacing="0" height="4">
    <tr>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
    </tr>
  </table>
<p>&nbsp;<!-- 
	<p align="center">
		Peter - use this report for right now please =&gt;
	<a href="system_utilization_report.asp">utilization report</a></p>
	--><form method="post" name="display_utilization" >
<p class="style5">&nbsp; Current Open Contract and Reservation Detail<p class="style5">&nbsp;<div align="center">
        <table border="0" cellpadding="0" style="width: 645px;" bordercolor="#111111" class="style7" id="open">
          <tr>
           <td width="100%" class="boxtitle" colspan="8" style="height: 15px"><font size="2"><b>
           Directions:</b> There is nothing on this page to manipulate. This is 
			a detailed report of all the reservations for a specific day, 
			location and car type as selected on the previous page.</font><p>
           &nbsp;</p>
           </td>
           
          </tr>
          <tr>
           	<td class="boxtitle" >&nbsp;</td>
           	<td class="style10" >
			<font size="2">Date displayed:&nbsp;&nbsp; </font></td>
			<td class="boxtitle" colspan="2" >
			<input name="util_date" type="text" readonly="true" value="<%=FormatDateTime(Request("date"), 2) %>" size="10"></td>
            <td class="boxtitle" colspan="2">
			&nbsp;</td>
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
           <td  class="boxtitle"  style="height: 15px; background-color:yellow" colspan="2"><font size="2">Due back today or earlier</font></td>
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
            <td  class="boxtitle" style="height: 11px"><font size="2"><u>Capture Date*</u></font></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Contract</font></u></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Rented LOR</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Check Out</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Check Out</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Est. Return</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Class</font></u></td>
          </tr>
          <% intCount = 0 %>          
          <% If adoRS.State = adStateOpen Then %>
          <% While (adoRS.EOF = False) %>
          <tr>
           <% If adoRS.Fields("estimated_return").Value < Now Then %>
          
	            <td class="boxtitle" style="width: 14%" ><font size="2"><%=adoRS.Fields("capture_date").Value %></font></td>
	            <td colspan="2">
	            <input type="text" name="utilization" size="40" readonly value="<%=adoRS.Fields("res_number").Value %>" style="text-align: left; width: 160px; background-color:yellow;" ></td>
	            <td class="boxtitle" style="width: 14%">
	            <input type="text" name="on_rent" size="20" readonly value="<%=adoRS.Fields("booked_lor").Value %>" style="text-align: right; width: 80px; background-color:yellow;"></td>
	            <td class="boxtitle" style="width: 14%">
	            <input type="text" name="canceled" size="20" readonly value="<%=FormatDateTime(adoRS.Fields("actual_check_out_date").Value, 2) %>" style="text-align: right; width: 80px; background-color:yellow;" ></td>
	            <td class="boxtitle" style="width: 14%">
	            <input type="text" name="returned" size="20" readonly value="<%=FormatDateTime(adoRS.Fields("actual_check_out_time").Value, 4) %>" style="text-align: right; width: 80px; background-color:yellow;"></td>
	            <td class="boxtitle" style="width: 14%">
	           	<input type="text" name="total0" size="20" readonly value="<%=FormatDateTime(adoRS.Fields("estimated_return").Value, 2) %>" style="text-align: right; width: 80px; background-color:yellow;" ></td>
	            <td class="boxtitle" style="width: 14%">
	            <input type="text" name="total" size="20" readonly value="<%=adoRS.Fields("booked_car_class").Value %>" style="text-align: right; width: 80px; background-color:yellow;" ></td>

            <% Else %>
	            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("capture_date").Value %></font></td>
	            <td colspan="2">
	            <input type="text" name="utilization" readonly value="<%=adoRS.Fields("res_number").Value %>" style="text-align: left; width: 160px;" size="40"></td>
	            <td class="boxtitle" style="width: 14%">
	            <input type="text" name="on_rent" size="20" readonly value="<%=adoRS.Fields("booked_lor").Value %>" style="text-align: right; width: 80px;"></td>
	            <td class="boxtitle" style="width: 14%">
	            <input type="text" name="canceled" size="20" readonly value="<%=FormatDateTime(adoRS.Fields("actual_check_out_date").Value, 2) %>" style="text-align: right; width: 80px;">            </td>
	            <td class="boxtitle" style="width: 14%">
	            <input type="text" name="returned" size="20" readonly value="<%=FormatDateTime(adoRS.Fields("actual_check_out_time").Value, 4) %>" style="text-align: right; width: 80px;"></td>
	            <td class="boxtitle" style="width: 14%">
	           	<input type="text" name="total0" size="20" readonly value="<%=FormatDateTime(adoRS.Fields("estimated_return").Value, 2) %>" style="text-align: right; width: 80px;"></td>
	            <td class="boxtitle" style="width: 14%">
	            <input type="text" name="total" size="20" readonly value="<%=adoRS.Fields("booked_car_class").Value %>" style="text-align: right; width: 80px;"></td>

            <% End If %>
          </tr>
          <% 	intCount = intCount + 1	%>
          <%   adoRS.MoveNext         	%>
          <% Wend                     	%>
          <% End If 					%>
          </table>
		Total: <%=intCount %>
<br>
<br>
        <table border="0" cellpadding="0" style="width: 645px;" bordercolor="#111111" class="style7" id="open">
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
            <td  class="boxtitle" style="height: 11px"><font size="2"><u>Capture Date*</u></font></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">
			Reservation No.</font></u></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Booked LOR</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Booked Out</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Booked Out</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Est. Return</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Class</font></u></td>
          </tr>
          <% intCount = 0 %>          
          <% Set adoRS = adoRS.NextRecordset %>
          <% If adoRS.State = adStateOpen Then %>
          <% While (adoRS.EOF = False) %>
          <tr>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("capture_date").Value %></font></td>
            <td colspan="2">
            <input type="text" name="utilization" readonly value="<%=adoRS.Fields("res_number").Value %>" style="text-align: left; width: 160px;" size="40"></td>
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="on_rent" size="20" readonly value="<%=adoRS.Fields("booked_lor").Value %>" style="text-align: right; width: 80px;"></td>
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="canceled" size="20" readonly value="<%=FormatDateTime(adoRS.Fields("actual_check_out_date").Value, 2) %>" style="text-align: right; width: 80px;">            </td>
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="returned" size="20" readonly value="<%=FormatDateTime(adoRS.Fields("actual_check_out_time").Value, 4) %>" style="text-align: right; width: 80px;"></td>
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="total0" size="20" readonly value="<%=FormatDateTime(adoRS.Fields("estimated_return").Value, 2) %>" style="text-align: right; width: 80px;"></td>
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="total" size="20" readonly value="<%=adoRS.Fields("booked_car_class").Value %>" style="text-align: right; width: 80px;"></td>
          </tr>
          <% 	intCount = intCount + 1	%>
          <%   adoRS.MoveNext         	%>
          <% Wend                     	%>
          <% End If 					%>
          </table>
		Total: <%=intCount %>
        
          
        </div>
        <p class="style11">* Capture date shows the date and time the most 
		recent data for this reservation has been received.&nbsp;
        &nbsp;&nbsp; 
          </p>
        <p align="center">
		<a target="_blank" href="system_utilization_res_number_tracker.asp">
		reservation number transaction history</a></p>
  </FORM>

  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1114" cellspacing="0" height="4" id="table1">
    <tr>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
    </tr>
  </table>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
<% Set adoRS = Nothing
   Set adoCmd = Nothing 
%>
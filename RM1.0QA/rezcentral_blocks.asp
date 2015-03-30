<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<% Response.Expires = -1  
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache" 
   
   	on error resume next

   	Server.ScriptTimeout = 30

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strConn = Session("pro_con")
	
	If Request("delete") = "true" Then
		Set adoCmd = CreateObject("ADODB.Command")
		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_rate_change_queue_tsd_rezcentral_rate_block_delete"
		adoCmd.CommandType = adCmdStoredProc
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@BlockID",      3, 1, 0, Request("block_id"))
			
		adoCmd.Execute
		
	End If
	
	If Request("Branch") <> "" Then
	
		Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_rate_change_queue_tsd_rezcentral_rate_block_insert"
		adoCmd.CommandType = adCmdStoredProc
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",           3, 1,  0, strUserId)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@Branch", 		    200, 1,  6, Request("Branch"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@ClassCode", 	    200, 1,  4, Request("ClassCode"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@RateType", 		200, 1,  1, Request("RateType"))		
		adoCmd.Parameters.Append adoCmd.CreateParameter("@RateCode", 	    200, 1, 20, Request("RateCode"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@RatePlan",      	 17, 1,  0, Request("RatePlan"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@StartDate",       135, 1,  0, Request("StartDate"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@EndDate",         135, 1,  0, Request("EndDate"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@Expires",         135, 1,  0, Request("Expires"))

			
		Set adoRS = adoCmd.Execute

	
	End If
		
	Set adoCmd = CreateObject("ADODB.Command")
	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_rate_change_queue_tsd_rezcentral_rate_block_select"
	adoCmd.CommandType = adCmdStoredProc
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",      3, 1, 0, strUserId)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@one_way",     11, 1, 0, 0)
	Set adoRS = adoCmd.Execute

	'GET THE CITIES THIS USER CAN SEE
	Set adoCmd = CreateObject("ADODB.Command")
	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "user_city_select"
	adoCmd.CommandType = adCmdStoredProc
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",      3, 1, 0, strUserId)
	Set adoRS1 = adoCmd.Execute
	
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
<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; RezCentral blocks settings</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script language="JavaScript" type="text/JavaScript" src="inc/sitewide.js" ></script>
<script type="text/javascript" language="javascript" src="inc/pupdate.js"></script>
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
.style10 {
	text-align: right;
	border-style: solid;
	border-width: 0;
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
&nbsp;<img alt="RezCentral" src="images/rezcentral.jpg"  ><strong> 
</strong></p>
  <table border="0" style="width: 800px;" bordercolor="#FFFFFF"  align="center" class="style7">
    <tr>
      <td ><font size="2" face="Vendana, Arial, Helvetica, sans-serif">[<a href="rezcentral_tethering_20130715.asp">tether settings</a>]
[<a href="rezcentral_tethering_ow_20130715.asp">tethering one-way settings</a>] [<a href="system_queue_rezcentral.asp">queue 
status</a>] [<a href="rezcentral_update_status.asp">report status</a>] [<a href="RezCentralHeader.aspx?uid=<%=strUserId %>">Rate 
		Code Detail</a>] [<a href="https://rezcentral.tsdasp.net/WebRezClient/" target="_blank">login 
to RezCentral</a>] <b>[block settings] </b>
</font></td>
     
    </tr>
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
	--><form method="get" name="add_block" >
<p class="style5">&nbsp; Current Block Settings<p class="style5">&nbsp;<div align="center">
        <table border="0" cellpadding="0" style="width: 700px;" bordercolor="#111111" class="style7">
          <tr>
           <td width="100%" class="boxtitle" colspan="9" style="height: 15px"><font size="2"><b>
           Directions:</b> To create a new block for RezCentral updates, enter the 
			values in the fields at the bottom of the list. To delete a tether, 
			click the delete link to the right of the tether to delete. To 
			update, simply delete then recreate.<br>&nbsp;</font><br>
		   <font size="2"><strong>Note</strong>: Dates must be entered in 
		   MM/DD/YYYY format.</font><p>
           &nbsp;</p>
           </td>
           
          </tr>
          <tr>
           	<td class="boxtitle" >&nbsp;</td>
           	<td class="style10" >&nbsp;</td>
			<td class="boxtitle" colspan="2" >
			&nbsp;</td>
            <td class="boxtitle" colspan="2">
			&nbsp;</td>
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
           
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px"></td>
           <td  class="boxtitle"  style="height: 15px"></td>
           <td  class="boxtitle"  style="height: 15px"></td>
           <td  class="boxtitle"  style="height: 15px"> 
          	</td>
           <td  class="boxtitle"  style="height: 15px"></td>
           <td  class="boxtitle"  style="height: 15px"></td>
           <td  class="boxtitle"  style="height: 15px"></td>
           
           <td  class="boxtitle"  style="height: 15px"></td>
           
           <td  class="boxtitle"  style="height: 15px"></td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">
			&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
            <td  class="boxtitle" style="height: 11px"><font size="2"><u>Branch</u></font></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Class Code.</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Rate Type</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Rate Code</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Plan</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Start Date</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">End Date</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Expires</font></u></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
          </tr>
          <% 	Dim intCount 	%>
          <% 	intCount = 0	%>
          
          <% If adoRS.State = adStateOpen Then %>
          <% While (adoRS.EOF = False) %>
          <tr>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("Branch").Value %></font></td>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("ClassCode").Value %></font></td>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("RateType").Value %></font></td>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("RateCode").Value %></font></td>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("RatePlan").Value %></font></td>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("StartDate").Value %></font></td>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("EndDate").Value %></font></td>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("Expires").Value %></font></td>
            
  
            <td class="style14" style="height: 11px"><a  href="rezcentral_blocks.asp?delete=true&block_id=<%=adoRS.Fields("BlockId").Value %>">&nbsp;&nbsp;delete</a></td>
          </tr>
          <% 	intCount = intCount + 1	%>
          <%   adoRS.MoveNext         	%>
          <% Wend                     	%>
          <% End If 					%>
          <tr>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="style13" style="height: 11px">
			&nbsp;</td>
            <td class="style13" style="height: 11px">
			&nbsp;</td>
            <td class="style12" style="height: 11px">&nbsp;</td>
          </tr>
          <tr>
            <td class="boxtitle" style="width: 14%">
			<select name="Branch">
			<% while not adoRS1.EOF 
				Response.Write "<option value='" & adoRS1.Fields("city_cd") & "'>" & adoRS1.Fields("city_cd") & "</option>" & vbcrlf
				adoRS1.MoveNext
			wend %>
			</select>
            </td>
            <td class="boxtitle" style="width: 14%">
			<input name="ClassCode" type="text" size="4" value="All"></td>
            <td class="boxtitle" style="width: 14%">
			<input name="RateType" type="text" size="1" value="S"></td>
            <td class="boxtitle" style="width: 14%">
			<input name="RateCode" type="text" size="20" value="All"></td>
            <td class="boxtitle" style="width: 14%">
			<select name="RatePlan">
			<option selected="" value="10">Daily</option>
			<option value="20">Weekend</option>
			<option value="30">Weekly</option>
			<option value="40">Monthly</option>
			<option value="0">All</option>
			</select></td>
            <td class="boxtitle" style="width: 14%">
			<input name="StartDate" type="text" size="10" value="mm/dd/yyyy" style="width: 80px"></td>
            <td class="style13" style="height: 11px">
			<input name="EndDate"  type="text" size="10" value="mm/dd/yyyy" style="width: 80px"></td>
            <td class="style13" style="height: 11px">
			<input name="Expires"  type="text" size="10" value="mm/dd/yyyy" style="width: 80px">
            <!--
			<input name="expires" id="expires" type="text" value="<%=FormatDateTime(datUtilDate, 2) %>" size="8"><img src="images/cal_button.gif" class="DatePicker" alt="Pick a date to expire the rate block" height="20" width="32" onClick="getCalendarFor(document.display_blocks.expires);return false" >
			-->
			</td>
            <td class="style12" style="height: 11px">&nbsp;</td>
          </tr>
          </table>
		<input name="Add" type="submit" value="Add Block"><br>
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
<p class="style12">
u: <%=strUserId %><br>
b: <%=Request("Branch") %>
</p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
<% Set adoRS = Nothing
   Set adoCmd = Nothing 
%>
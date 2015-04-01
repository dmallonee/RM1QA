<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<% Response.Expires = -1  
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache" 
   
   	'on error resume next

   	Server.ScriptTimeout = 30

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strConn = Session("pro_con")
	
	If Request("delete") = "true" Then
		Set adoCmd = CreateObject("ADODB.Command")
		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_rate_change_queue_tsd_rezcentral_tether_delete"
		adoCmd.CommandType = adCmdStoredProc
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@tether_id",      3, 1, 0, Request("tether_id"))
			
		adoCmd.Execute
		
	End If
	
	If Request("Branch") <> "" Then
	
		Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_rate_change_queue_tsd_rezcentral_tether_insert"
		adoCmd.CommandType = adCmdStoredProc
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",           3, 1,  0, strUserId)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@Branch", 		    200, 1, 10, Request("Branch"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@DestBranch",      200, 1, 10, Null)
        if(Request("RateCode") = "") then
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@RateCode", 	 200, 1, 20, "****")
        else
    		adoCmd.Parameters.Append adoCmd.CreateParameter("@RateCode", 	 200, 1, 20, Request("RateCode"))
        end if
		adoCmd.Parameters.Append adoCmd.CreateParameter("@RatePlan",      	 17, 1,  0, Request("RatePlan"))
        if(Request("System") = "") then
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@System", 	    200, 1, 20, "****")
        else
 		    adoCmd.Parameters.Append adoCmd.CreateParameter("@System", 		200, 1, 25, Request("System"))
        end if
		adoCmd.Parameters.Append adoCmd.CreateParameter("@ClassCode", 	    200, 1,  4, Request("ClassCode"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@NewClassCode",    200, 1,  4, Request("NewClassCode"))
        if(Request("DiffAmt") = "") then
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@DiffAmt", 	  6, 1,  0, 0)
        else
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@DiffAmt", 	  6, 1,  0, Request("DiffAmt"))
        end if
		adoCmd.Parameters.Append adoCmd.CreateParameter("@DiffIsDollar",     11, 1,  0, True)
        if(Request("ExtraDayDiffAmt") = "") then
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@ExtraDayDiff", 	  6, 1,  0, 0)
        else
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@ExtraDayDiff", 	  6, 1,  0, Request("ExtraDayDiffAmt"))
        end if
		adoCmd.Parameters.Append adoCmd.CreateParameter("@ExtraDayIsDollar", 11, 1,  0, True)
			
		Set adoRS = adoCmd.Execute
	Else
		Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_rate_change_queue_tsd_rezcentral_tether_select"
		adoCmd.CommandType = adCmdStoredProc
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",      3, 1, 0, strUserId)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@one_way",     11, 1, 0, 0)
					
		Set adoRS = adoCmd.Execute
	End If
	
	
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
<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; RezCentral Tethering settings</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script language="JavaScript" type="text/JavaScript" src="inc/sitewide.js" ></script>
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
	font-size: small;
}
.style13 {
	text-align: right;
}
.style14 {
	font-size: small;
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
      <td align="center"><font size="2" face="Vendana, Arial, Helvetica, sans-serif"><b>[Tether Settings] </b>
[<a href="rezcentral_tethering_ow.asp">Tether One-Way Settings</a>] [<a href="system_queue_rezcentral.asp">Queue 
Status</a>] [<a href="rezcentral_update_status.asp">Report Status</a>] [<a href="RezCentralHeader.aspx?uid=<%=strUserId %>">Rate 
		Code Detail</a>] [<a href="https://rezcentral.tsdasp.net/WebRezClient/" target="_blank">Login 
to RezCentral</a>] [<a href="rezcentral_blocks.asp" target="_blank">Block Settings</a>] [<a href="rezcentral_fleet_adjustment.asp" target="_blank">Fleet Adjustment</a>]</font></td>
     
    </tr>
    <tr>
      <td background="images/ruler.gif" height="4"></td>
     
    </tr>
  </table>
<p>&nbsp;<!-- 
	<p align="center">
		Peter - use this report for right now please =&gt;
	<a href="system_utilization_report.asp">utilization report</a></p>
	--><form method="get" name="add_tether_setting" >
<p class="style5"><b>Current Tether Settings</b>
<div align="center">
        <table border="0" cellpadding="0" style="width: 700px;border-color=#ffffff #ffffff">
          <tr>
           <td width="100%" class="boxtitle" colspan="9" style="height: 15px"><font size="2"><b>
           Directions:</b> To create a new tether for RezCentral, enter the 
			values in the fields at the bottom of the list. To delete a tether, 
			click the delete link to the right of the tether. To 
			update or replace a tether, delete the target tether then recreate with the changed 
               or replacement values.</font>
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

          <tr class="profile_header">
            <td  class="boxtitle" style="height: 11px"><font size="2">Branch</font></td>
            <td  class="boxtitle" style="height: 11px"><font size="2">Rate Code.</font></td>
            <td  class="boxtitle" style="height: 11px"><font size="2">Plan</font></td>
            <td  class="boxtitle" style="height: 11px"><font size="2">System</font></td>
            <td  class="boxtitle" style="height: 11px"><font size="2">Base</font></td>
            <td  class="boxtitle" style="height: 11px"><font size="2">New</font></td>
            <td  class="boxtitle" style="height: 11px"><font size="2">Diff</font></td>
            <td  class="boxtitle" style="height: 11px"><font size="2">XDay Diff</font></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
          </tr>
          <% 	Dim intCount 	%>
          <% 	intCount = 0	%>
          <% If adoRS.State = adStateOpen Then %>
          <%    Dim ii  %>
          <% ii = 0 %>
          <% While (adoRS.EOF = False) %>
          <% ii = ii + 1 %>
          <% if adoRS.Fields("RatePlan") = 10 then rctext = "Daily" %>
          <% if adoRS.Fields("RatePlan") = 20 then rctext = "Weekend" %>
          <% if adoRS.Fields("RatePlan") = 30 then rctext = "Weekly" %>
          <% if adoRS.Fields("RatePlan") = 40 then rctext = "Monthly" %>
  		  <% If strClass <> "profile_dark" Then
        			strClass = "profile_dark" 'background-color:#B2BEC4
        		Else
        			strClass = "profile_light" 'background-color:#CFD7DB
        		End If
            %>
           <tr class="<%= strClass %>">
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("Branch").Value %></font></td>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("RateCode").Value %></font></td>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=rctext %></font></td>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("System").Value %></font></td>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("ClassCode").Value %></font></td>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS.Fields("NewClassCode").Value %></font></td>
            <td class="style13" style="height: 11px"><font size="2">
            <% If CBool(adoRS.Fields("DiffIsDollar").Value) Then %>
            <%=FormatCurrency(adoRS.Fields("DiffAmt").Value) %>
            <% Else %>
            <%=FormatPercent(adoRS.Fields("DiffAmt").Value) %>
            <% End If %>
            </font></td>
            <td class="style13" style="height: 11px"><font size="2">
            <% If CBool(adoRS.Fields("ExtraDayIsDollar").Value) Then %>
            <%=FormatCurrency(adoRS.Fields("ExtraDayDiff").Value) %>
            <% Else %>
            <%=FormatPercent(adoRS.Fields("ExtraDayDiff").Value) %>
            <% End If %>
            </font></td>
            <td class="style14" style="height: 11px"><a  href="rezcentral_tethering.asp?delete=true&tether_id=<%=adoRS.Fields("tether_id").Value %>">&nbsp;&nbsp;delete</a></td>
          </tr>
          <% 	intCount = intCount + 1	%>
          <%   adoRS.MoveNext         	%>
          <% Wend                     	%>
          <% End If 					%>
          <tr>
          <td height=10></td>
          </tr>
          <tr>
            <td class="boxtitle" style="width: 14%">
			<input name="Branch" type="text" size="5"></td>
            <td class="boxtitle" style="width: 14%">
			<input name="RateCode" type="text" size="10"></td>
            <td class="boxtitle" style="width: 14%"><select name="RatePlan">
			<option selected="" value="10">Daily</option>
			<option value="20">Weekend</option>
			<option value="30">Weekly</option>
			<option value="40">Monthly</option>
			</select></td>
            <td class="boxtitle" style="width: 14%">
			<input name="System" type="text" size="8"></td>
            <td class="boxtitle" style="width: 14%">
			<input name="ClassCode" type="text" size="5"></td>
            <td class="boxtitle" style="width: 14%">
			<input name="NewClassCode" type="text" size="5"></td>
            <td class="style13" style="height: 11px">
			<input name="DiffAmt" type="text" size="5"></td>
            <td class="style13" style="height: 11px">
			<input name="ExtraDayDiffAmt" type="text" size="5"></td>
            <td class="style12" style="height: 11px"><input name="Add" type="submit" value="Add New Tether"></td>
          </tr>
          </table>
  		<br>
		Total: <%=intCount %>
        </div>
        <p class="style11">&nbsp;</p>
  </FORM>

  <table border="0" style="width: 800px;" bordercolor="#FFFFFF" height="4" id="table1" align="center" class="style7">
    <tr>
      <td background="images/ruler.gif"></td>
      
    </tr>
    <tr>
      <td align="center"><font size="2" face="Vendana, Arial, Helvetica, sans-serif"><b>[Tether Settings] </b>
[<a href="rezcentral_tethering_ow.asp">Tether One-Way Settings</a>] [<a href="system_queue_rezcentral.asp">Queue 
Status</a>] [<a href="rezcentral_update_status.asp">Report Status</a>] [<a href="RezCentralHeader.aspx?uid=<%=strUserId %>">Rate 
		Code Detail</a>] [<a href="https://rezcentral.tsdasp.net/WebRezClient/" target="_blank">Login 
to RezCentral</a>] [<a href="rezcentral_blocks.asp" target="_blank">Block Settings</a>] [<a href="rezcentral_fleet_adjustment.asp" target="_blank">Fleet Adjustment</a>]</font></td>
     
    </tr>
  </table>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
<% Set adoRS = Nothing
   Set adoCmd = Nothing 
%>
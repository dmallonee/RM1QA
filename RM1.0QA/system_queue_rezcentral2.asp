<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<% Response.Expires = -1  
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache" 
   
   	'on error resume next
   	
   	Dim Updates(24, 2)
   	
   	Server.ScriptTimeout = 30

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_rate_change_queue_tsd_rezcentral_status"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",      3, 1, 0, strUserId)
		
	Set adoRSNotUploaded = adoCmd.Execute
	
	Dim strUploaded
	Dim strWaiting

	'intHour = 0

	While (adoRSNotUploaded.EOF = False) 
		Updates(adoRSNotUploaded.Fields("hour").Value, 0) = adoRSNotUploaded.Fields("not uploaded").Value
		strNotUploaded = strNotUploaded & "," & (adoRSNotUploaded.Fields("not uploaded").Value / 1000)
		strNotUploadedHour = strNotUploadedHour & "," & adoRSNotUploaded.Fields("hour").Value

		adoRSNotUploaded.MoveNext
	Wend


	Set adoRSUploaded = adoRSNotUploaded.NextRecordset
	
	'intHour = 0

	While (adoRSUploaded.EOF = False) 
		Updates(adoRSUploaded.Fields("hour").Value, 1) = adoRSUploaded.Fields("uploaded").Value
		strUploaded = strUploaded & "," & (adoRSUploaded.Fields("uploaded").Value / 1000)
		strUploadedHour = strUploadedHour & "," & adoRSUploaded.Fields("hour").Value

		adoRSUploaded.MoveNext

	Wend

	
	
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

<title>Rate-Monitor by Rate-Highway, Inc. | RezCentral Queue Status</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="refresh" content="300;url=system_queue_rezcentral.asp">
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
.style11 {
	text-align: center;
}
.style12 {
	text-align: right;
	border-style: solid;
	border-width: 0;
}
.style13 {
	font-size: x-small;
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
&nbsp;<img alt="RezCentral" src="images/rezcentral.jpg" ><strong> 
</strong></p>
<p><font size="2" face="Vendana, Arial, Helvetica, sans-serif">[<a href="rezcentral_tethering_20130715.asp">tethering settings</a>]<b> </b>
[<a href="rezcentral_tethering_ow_20130715.asp">tethering one-way settings</a>] [<strong>queue 
status</strong>] [<a href="rezcentral_update_status.asp">report status</a>]</font></p>
  <p>&nbsp;</p>
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
<p>&nbsp;
	<form method="post" name="display_utilization" >
	<p class="style5">&nbsp; Current Queue Levels - By Hour - Page 2 (<a href="system_queue_rezcentral.asp">Page 
	1</a>)<div align="center">
        <table border="0" cellpadding="0" style="width: 750px;" bordercolor="#111111" class="style7" name="contracts" id="contracts">
          <tr>
           <td width="100%" class="boxtitle" colspan="7" style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
 			<td class="style12" colspan="2" >&nbsp;</td>
			<td class="boxtitle" colspan="2" >
			&nbsp;</td>
            <td class="boxtitle" colspan="2">&nbsp;</td>
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
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style13"  style="height: 15px">Noon</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style13"  style="height: 15px">1 pm</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style11"  style="height: 15px"><img alt="hour 1" src="http://chart.apis.google.com/chart?cht=p3&chd=t:<%=Fix(Updates(12, 1)/1000)%>,<%=Fix(Updates(12,0)/1000)%>&chs=300x150&chco=ff8c00,ffebcd"></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style11"  style="height: 15px"><img alt="hour 2" src="http://chart.apis.google.com/chart?cht=p3&chd=t:<%=Fix(Updates(13,1)/1000)%>,<%=Fix(Updates(13,0)/1000)%>&chs=300x150&chco=ff8c00,ffebcd"></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="style13" style="height: 11px">Uploaded: <%=FormatNumber(Updates(12,1),0)%> Waiting: <%=FormatNumber(Updates(12,0),0)%></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="style13" style="height: 11px">Uploaded: <%=FormatNumber(Updates(13,1),0)%> Waiting: <%=FormatNumber(Updates(13,0),0)%></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
          </tr>
         
          <tr>
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
           <td  class="style13"  style="height: 15px">2 pm</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style13"  style="height: 15px">3 pm</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style11"  style="height: 15px"><img alt="hour 3" src="http://chart.apis.google.com/chart?cht=p3&chd=t:<%=Fix(Updates(14,1)/1000)%>,<%=Fix(Updates(14,0)/1000)%>&chs=300x150&chco=ff8c00,ffebcd"></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style11"  style="height: 15px"><img alt="hour 4" src="http://chart.apis.google.com/chart?cht=p3&chd=t:<%=Fix(Updates(15,1)/1000)%>,<%=Fix(Updates(15,0)/1000)%>&chs=300x150&chco=ff8c00,ffebcd"></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="style13" style="height: 11px">Uploaded: <%=FormatNumber(Updates(14,1),0)%> Waiting: <%=FormatNumber(Updates(14,0),0)%></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="style13" style="height: 11px">Uploaded: <%=FormatNumber(Updates(15,1),0)%> Waiting: <%=FormatNumber(Updates(15,0),0)%></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
          </tr>

          <tr>
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
           <td  class="style13"  style="height: 15px">4 pm</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style13"  style="height: 15px">5 pm</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style11"  style="height: 15px"><img alt="hour 3" src="http://chart.apis.google.com/chart?cht=p3&chd=t:<%=Fix(Updates(16,1)/1000)%>,<%=Fix(Updates(16,0)/1000)%>&chs=300x150&chco=ff8c00,ffebcd"></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style11"  style="height: 15px"><img alt="hour 4" src="http://chart.apis.google.com/chart?cht=p3&chd=t:<%=Fix(Updates(17,1)/1000)%>,<%=Fix(Updates(17,0)/1000)%>&chs=300x150&chco=ff8c00,ffebcd"></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="style13" style="height: 11px">Uploaded: <%=FormatNumber(Updates(16,1),0)%> Waiting: <%=FormatNumber(Updates(16,0),0)%></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="style13" style="height: 11px">Uploaded: <%=FormatNumber(Updates(17,1),0)%> Waiting: <%=FormatNumber(Updates(17,0),0)%></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
          </tr>

          <tr>
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
           <td  class="style13"  style="height: 15px">6 pm</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style13"  style="height: 15px">7 pm</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style11"  style="height: 15px"><img alt="hour 3" src="http://chart.apis.google.com/chart?cht=p3&chd=t:<%=Fix(Updates(18,1)/1000)%>,<%=Fix(Updates(18,0)/1000)%>&chs=300x150&chco=ff8c00,ffebcd"></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style11"  style="height: 15px"><img alt="hour 4" src="http://chart.apis.google.com/chart?cht=p3&chd=t:<%=Fix(Updates(19,1)/1000)%>,<%=Fix(Updates(19,0)/1000)%>&chs=300x150&chco=ff8c00,ffebcd"></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="style13" style="height: 11px">Uploaded: <%=FormatNumber(Updates(18,1),0)%> Waiting: <%=FormatNumber(Updates(18,0),0)%></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="style13" style="height: 11px">Uploaded: <%=FormatNumber(Updates(19,1),0)%> Waiting: <%=FormatNumber(Updates(19,0),0)%></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
          </tr>

          <tr>
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
           <td  class="style13"  style="height: 15px">8 pm</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style13"  style="height: 15px">9 pm</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style11"  style="height: 15px"><img alt="hour 3" src="http://chart.apis.google.com/chart?cht=p3&chd=t:<%=Fix(Updates(20,1)/1000)%>,<%=Fix(Updates(20,0)/1000)%>&chs=300x150&chco=ff8c00,ffebcd"></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style11"  style="height: 15px"><img alt="hour 4" src="http://chart.apis.google.com/chart?cht=p3&chd=t:<%=Fix(Updates(21,1)/1000)%>,<%=Fix(Updates(21,0)/1000)%>&chs=300x150&chco=ff8c00,ffebcd"></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="style13" style="height: 11px">Uploaded: <%=FormatNumber(Updates(20,1),0)%> Waiting: <%=FormatNumber(Updates(20,0),0)%></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="style13" style="height: 11px">Uploaded: <%=FormatNumber(Updates(21,1),0)%> Waiting: <%=FormatNumber(Updates(21,0),0)%></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
          </tr>

          <tr>
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
           <td  class="style13"  style="height: 15px">10 pm</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style13"  style="height: 15px">11 pm</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style11"  style="height: 15px"><img alt="hour 3" src="http://chart.apis.google.com/chart?cht=p3&chd=t:<%=Fix(Updates(22,1)/1000)%>,<%=Fix(Updates(22,0)/1000)%>&chs=300x150&chco=ff8c00,ffebcd"></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style11"  style="height: 15px"><img alt="hour 4" src="http://chart.apis.google.com/chart?cht=p3&chd=t:<%=Fix(Updates(12 + 11,1)/1000)%>,<%=Fix(Updates(12 + 11,0)/1000)%>&chs=300x150&chco=ff8c00,ffebcd"></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="style13" style="height: 11px">Uploaded: <%=FormatNumber(Updates(22,1),0)%> Waiting: <%=FormatNumber(Updates(22,0),0)%></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="style13" style="height: 11px">Uploaded: <%=FormatNumber(Updates(12 + 11,1),0)%> Waiting: <%=FormatNumber(Updates(12 + 11,0),0)%></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
          </tr>


          </table>
		

        </div>
        <p align="center">&nbsp;</p>
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
<% 
	strUploaded = RIGHT(strUploaded, LEN(strUploaded) - 1)
	strUploadedHour = RIGHT(strUploadedHour, LEN(strUploadedHour) - 1)

	'strNotUploaded = RIGHT(strNotUploaded, LEN(strNotUploaded) - 1)
	'strNotUploadedHour = RIGHT(strNotUploadedHour, LEN(strNotUploadedHour) - 1)


%>

<img alt="hour 4" src="http://chart.apis.google.com/chart?cht=lxy&chd=t:<%=strUploadedHour & "|" & strUploaded %>&chs=600x400&chds=0,23,0,100&chxt=x,y&chxl=0:|Midnight|Noon|11pm|1:|0|25,000|50,000|75,000|&chls=3,1,0""><br>
uploaded = <%=strNotUploaded %>
<br>uploaded hour = <%=strNotUploadedHour %>
<p>&nbsp;</p>
</body>
</html>
<% Set adoCmd = Nothing 
%>
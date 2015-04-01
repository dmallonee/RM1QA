<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" --> 
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 

<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   	Server.ScriptTimeout = 180

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "support_totals_by_org_weekly"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@org_id",   3, 1,  0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",  3, 1,  0, strUserId)
		
	Set adoRS = adoCmd.Execute

	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">

<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; System</title>
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
<style type="text/css" >
<!--
.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.style1 {
	border-collapse: collapse;
	border-style: solid;
	border-width: 1px;
}
.style2 {
	border-collapse: collapse;
}
.style3 {
	text-align: center;
	font-size: small;
	color: #0000FF;
}
.style5 {
	border-width: 0;
	height= "68" text-align:left;
	padding-left: 3;
	padding-right: 3;
	padding-top: 3;
	background-color: #CFD7DB;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10pt;
	vertical-align: bottom;
	text-align: right;
}
.style6 {
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
    <td background="images/h_tile.gif"><table width="100" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="images/h_system.gif" width="368" height="31"></td>
          <td><img src="images/h_right.gif" width="402" height="31"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p>&nbsp;&nbsp;&nbsp;<br>&nbsp;
<font size="2" face="Vendana, Arial, Helvetica, sans-serif">
&nbsp;<a href="javascript:not_enabled()">[custom city codes]</a>
&nbsp;<b>[system performance]</b>
&nbsp;<a href="system_proxy.asp">[proxy management]</a>
&nbsp;<a href="system_utilization.asp">[utilization settings]</a>
&nbsp;<a title="click to manage the utilization car groups" href="system_utilization_car_groups.asp">[utilization car groups]</a>
</font><br>&nbsp;</p>

  <table border="0" style="width: 800px;" bordercolor="#FFFFFF" height="4" class="style2" align="center">
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
  <table style="width: 800px;" bordercolor="#FFFFFF" id="profiles" class="style1" align="center">
    <tr>
      <td class="profile_header" height="45" style="width: 56px">Days Back</td>
      <td class="profile_header" height="45" style="width: 66px">Selected</td>
      <td class="profile_header" height="45" style="width: 178px">Date*</td>
      <td class="profile_header" height="45" style="width: 115px">Rates Requested</td>
      <td class="profile_header" height="45" style="width: 115px">Rates<br>Collected</td>
      <td class="profile_header" height="45" style="width: 635px">Notes</td>
    </tr>
<% Dim intCounter
   intCounter = 0                    %>
<% If adoRS.State = adStateOpen Then %>
<% While (adoRS.EOF = False) %>
    
    <tr>
      <td class="profile_light" height="20" style="width: 56px"><%=intCounter %></td>
      <td bgcolor="#FDC677" align="center" height="20" style="width: 66px">
      <input type="radio" value="V1" name="selected"></td>
      <td class="profile_light" height="20" style="width: 178px">
      <a href="system_performance_hourly.asp?days_back=<%=intCounter %>" ><%=adoRS.Fields("Date").Value %></a></td>
      <td class="style5" height="20" style="width: 115px"><%=FormatNumber(adoRS.Fields("requested").Value, 0) %></td>
      <td class="style5" height="20" style="width: 115px"><%=FormatNumber(adoRS.Fields("completed").Value, 0) %></td>
      <td class="profile_light" height="20" style="width: 635px">&nbsp;</td>
    </tr>


<% strDataValues = strDataValues & adoRS.Fields("requested").Value & "," %>
<% strDataValues2 = strDataValues2 & adoRS.Fields("completed").Value & "," %>
<% strDataLabels = strDataLabels & adoRS.Fields("Date").Value & "|" %>

<%   intCounter = intCounter + 1        %>
<%   adoRS.MoveNext        %>
<% Wend                     %>
<% End If %>
    
    </table>
  <table border="0" style="width: 800px;" bordercolor="#FFFFFF" height="4" class="style2" align="center">
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
  <p class="style6">*Click a date to view the hourly break-down of rates 
	collected for a given date.</p>
<p class="style6" style="height: 16px"><a href="system_performance_by_city.asp">
view with location break-down</a></p>
<p class="style6">
          <% If adoRS.State = adStateOpen Then %>
          <%
          
          strDataLabels = Left(strDataLabels, (Len(strDataLabels) - 1))
          strDataValues = Left(strDataValues, (Len(strDataValues) - 1))
          strDataValues2 = Left(strDataValues2, (Len(strDataValues2) - 1))

          %>
          <img src="http://chart.apis.google.com/chart?cht=ls&chd=t:<%=strDataValues %>|<%=strDataValues2 %>&chds=200000,500000&chco=FDC677,879AA2&chs=750x375&chdl=Requested|Completed&chxt=x,y&chxl=0:|<%=strDataLabels %>|1:|200,000|300,000|400,000|500,000|&chls=3,6,3|3,1,0" alt="Weekly Performace Graph" >

		  <% End If %>

</p>
<p class="style6">
          The above chart is not accurate - please disregard until development 
			is complete.</p>
<p class="style3"><strong>UNDER DEVELOPMENT</strong></p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
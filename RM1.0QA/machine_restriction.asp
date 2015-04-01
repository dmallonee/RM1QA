<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<% Response.Expires = -1  
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache" 
   
   Dim strCarClasses
   Dim strDataValues
   
   	on error resume next

   	Server.ScriptTimeout = 180

	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "machine_restriction_select"
	adoCmd.CommandType = adCmdStoredProc
		
	Set adoRS = adoCmd.Execute

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
<html xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<meta http-equiv="Content-Language" content="en-us">

<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; Machine Restriction</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script type="text/javascript" language="javascript" src="inc/sitewide.js" ></script>
<script type="text/javascript" language="javascript" src="inc/pupdate.js"></script>
<script type="text/JavaScript" language="JavaScript" >
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

function DisableTSDinfo()
{

	var s = document.getElementById("update_rezcentral");
	var v = s.options[s.selectedIndex].text;


	if (v == 'Yes')
	{
	document.utilization_update.tsd_customer_number.disabled = false;
	document.utilization_update.tsd_passcode.disabled = false;
	}
	else
	{
	document.utilization_update.tsd_customer_number.disabled = true;
	document.utilization_update.tsd_passcode.disabled = true;
	}

}

//-->
</script>



<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<style type="text/css" >
<!--
.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.style2 {
	font-size: x-small;
}
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
numeric_input {
	text-align: right;
}
}
.UtilGridValue {
	border: 0 solid #FFFFFF;
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
                  </font></div></td>
                    </tr>
                    <tr>
                      <td><img src="images/separator.gif" width="183" height="6" alt=""></td>
                    </tr>
                  </table>
                </td>
                <td><img src="images/user_tile.gif" width="7" height="31" alt=""></td>
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
        <td><img src="images/h_system.gif" width="368" height="31" alt=""></td>
        <td>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<p>&nbsp;&nbsp;&nbsp;<br>&nbsp;
<font size="2" face="Vendana, Arial, Helvetica, sans-serif">
&nbsp;</font><br>&nbsp;</p>
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
     <p>
<p class="style5">&nbsp; Machine Restriction / Assignment<div align="center">
        <table border="0" cellpadding="0" style="width: 600px;" bordercolor="#111111" class="style7">
          <tr>
           <td width="100%" class="boxtitle" colspan="5" style="height: 15px"><font size="2"><b>
           Directions:</b> To manage the machine restrictions settings manually please 
           use this page. Blah, blah, blah - the update ability is not available 
		   yet. Once you 
           are satisfied with your changes simple press the update button. If 
           you want to discard your changes and not save them, either navigate 
           away from this page or close the page.</font><p>
           <font size="2">&nbsp; </font>&nbsp;</p>
           </td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style10"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr class="profile_header">
            <td  class="boxtitle">Notes</td>
            <td  class="boxtitle">Machine Name</td>
            <td  class="boxtitle">Org Id</td>
            <td  class="boxtitle">Role Id</td>
            <td  class="boxtitle">IP Address</td>
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
            <td class="boxtitle" ><font size="2"><%=adoRS.Fields("client_userid").Value %></font></td>
            <td class="boxtitle" ><font size="2"><%=adoRS.Fields("machine_name").Value %></font></td>
            <td class="UtilGridValue" ><font size="2"><%=adoRS.Fields("org_id").Value %></font></td>
            <td class="UtilGridValue" ><font size="2"><%=adoRS.Fields("role_id").Value %></font></td>
            <td class="UtilGridValue" ><font size="2"><%=adoRS.Fields("ip").Value %></font></td>
          </tr>
          <% 	intCount = intCount + 1	%>
          <%   adoRS.MoveNext         	%>
          <% Wend                     	%>
          <% End If 					%>
          <tr>
          <td colspan="5" class="style2" >
         
          Org Id = 99 means that the machine will not be assigned jobs<br>Org Id 
		  = 44 means that this machine will only be assigned test jobs for 
		  development<br>Org Id = 00 means that this machine will only be 
		  assigned support jobs<br>Org Id = 01 means that this machine will be 
		  assigned any active customer's jobs</td>
          </tr>
          </table>
        </div>
        <p align="center">&nbsp;
        &nbsp;&nbsp; 
          </p>
        <p align="center">&nbsp;</p>
 

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
<%
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
  
<p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">© 
2002 - 2013 - All rights reserved<br>
<b>Rate-Highway, Inc.</b><br>
</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">18001 Cowan, 
    Suite F<br />
    Irvine, CA 92614<br />
    (949) 614-0751&nbsp; </font>

<p align="center">&nbsp;</p>
<p>&nbsp;</p>
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
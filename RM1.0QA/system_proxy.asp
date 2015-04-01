<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<% Response.Expires = -1  
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache" 

   Server.ScriptTimeout = 180

	Dim StrConn
	Dim adoCmd
	Dim strSQL
	Dim strProxyValue1
	Dim strProxyValue2
	Dim strProxyValue3
	Dim strProxyValue4
	Dim adoRS

	strProxyValue1 = Request.Form("Proxy1")
	strProxyValue2 = Request.Form("Proxy2")
	strProxyValue3 = Request.Form("Proxy3")
	strProxyValue4 = Request.Form("Proxy4")

	strConn = Session("pro_con")

 	If Len(strProxyValue1) > 0 Or Len(strProxyValue2) > 0 Or Len(strProxyValue3) > 0 Then 

		If Len(strProxyValue1) > 0 Then
		  	Set adoCmd = Server.CreateObject("ADODB.Command")

		  	adoCmd.ActiveConnection = strConn
		  	adoCmd.CommandText = "registry_data_update"
		  	adoCmd.CommandType = adCmdStoredProc
	  	
		  	adoCmd.Parameters.Refresh

	  		adoCmd.Parameters("@ValueName").Value = "ProxyServer"
	  		adoCmd.Parameters("@ValueData").Value = strProxyValue1
	  		adoCmd.Parameters("@ProxyGroup").Value = 1
  		
  			adoCmd.Execute

		End If

		If Len(strProxyValue2) > 0 Then
		  	Set adoCmd = Server.CreateObject("ADODB.Command")

		  	adoCmd.ActiveConnection = strConn
		  	adoCmd.CommandText = "registry_data_update"
		  	adoCmd.CommandType = adCmdStoredProc
	  	
		  	adoCmd.Parameters.Refresh

	  		adoCmd.Parameters("@ValueName").Value = "ProxyServer"
	  		adoCmd.Parameters("@ValueData").Value = strProxyValue2
	  		adoCmd.Parameters("@ProxyGroup").Value = 2
  		
  			adoCmd.Execute

		End If

		If Len(strProxyValue3) > 0 Then
		  	Set adoCmd = Server.CreateObject("ADODB.Command")

		  	adoCmd.ActiveConnection = strConn
		  	adoCmd.CommandText = "registry_data_update"
		  	adoCmd.CommandType = adCmdStoredProc
	  	
		  	adoCmd.Parameters.Refresh

	  		adoCmd.Parameters("@ValueName").Value = "ProxyServer"
	  		adoCmd.Parameters("@ValueData").Value = strProxyValue3
	  		adoCmd.Parameters("@ProxyGroup").Value = 3
  		
  			adoCmd.Execute

		End If

		If Len(strProxyValue4) > 3 Then
		
			If Left(strProxyValue4, 3) <> "Not" Then
			  	Set adoCmd = Server.CreateObject("ADODB.Command")

			  	adoCmd.ActiveConnection = strConn
			  	adoCmd.CommandText = "registry_data_update"
			  	adoCmd.CommandType = adCmdStoredProc
	  	
			  	adoCmd.Parameters.Refresh

	  			adoCmd.Parameters("@ValueName").Value = "ProxyServer"
		  		adoCmd.Parameters("@ValueData").Value = strProxyValue4
		  		adoCmd.Parameters("@ProxyGroup").Value = 4
  		
	  			adoCmd.Execute
	  			
	  		End If

		End If

  		
  		'Response.Redirect "menu.asp"
  	

	Else


	  	Set adoCmd = Server.CreateObject("ADODB.Command")

	  	adoCmd.ActiveConnection = strConn
	  	adoCmd.CommandText = "registry_data_select"
	  	adoCmd.CommandType = adCmdStoredProc
	  	
	  	adoCmd.Parameters.Refresh

  		adoCmd.Parameters("@ValueName").Value = "ProxyServer"
	
 		Set adoRS = adoCmd.Execute
 		
		strProxyValue1 = "Not Found"
 		strProxyValue2 = "Not Found"
 		strProxyValue3 = "Not Found"
 		strProxyValue4 = "Not Used"

 		
		While adoRS.EOF = False 
			Select Case adoRS.Fields("Proxy_Group_Id").Value
				Case 1 
					strProxyValue1 = adoRS.Fields("Value_Data").Value
					
				Case 2
					strProxyValue2 = adoRS.Fields("Value_Data").Value
				
				Case 3
					strProxyValue3 = adoRS.Fields("Value_Data").Value

				Case 4
					strProxyValue4 = adoRS.Fields("Value_Data").Value

			End Select
			
			adoRS.MoveNext 			
			
 		Wend
 
  	End If
  
%>    
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">

<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; System</title>
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
<table width="400" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="images/b_left.jpg" width="62" height="32"></td>
          <td><a href="search_profiles_car.asp" onMouseOver="MM_swapImage('s1','','images/b_search_pro_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_search_pro_of.gif" name="s1" border="0" id="s1" width="183" height="32"></a></td>
          <td><a href="search_queue_car.asp" onMouseOver="MM_swapImage('s2','','images/b_search_que_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_search_que_of.gif" name="s2" border="0" id="s2" width="97" height="32"></a></td>
          <td><a href="search_criteria_car.asp" onMouseOver="MM_swapImage('s3','','images/b_search_cri_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_search_cri_of.gif" name="s3" border="0" id="s3" width="103" height="32"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('ra','','images/b_rate_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_rate_of.gif" name="ra" border="0" id="ra" width="88" height="32"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('al','','images/b_alert_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_alert_of.gif" name="al" border="0" id="al" width="53" height="32"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('us','','images/b_user_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_user_of.gif" name="us" border="0" id="us" width="126" height="32"></a></td>
          <td><a href="system_proxy.asp" onMouseOver="MM_swapImage('sy','','images/b_system_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_system_of.gif" name="sy" border="0" id="sy" width="58" height="32"></a></td>
        </tr>
      </table>
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
<p>&nbsp;&nbsp;&nbsp; <br>
&nbsp;<font size="2" face="Vendana, Arial, Helvetica, sans-serif">&nbsp;<a href="javascript:not_enabled()">[custom 
city codes]</a>&nbsp; <b>[</b><a href="javascript:not_enabled()">system status</a><b>]</b>&nbsp; [<b>proxy 
management</b>] <a href="system_utilization.asp">[utilization settings]</a></font><br>
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
<p>&nbsp;
     <FORM METHOD="POST" NAME="proxy_value" action="system_proxy.asp"   > 
        <div align="center">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="500">
          <tr>
           <td width="100%" class="boxtitle" colspan="2">To change the proxy 
           server used by the search machines, enter a new value here in the 
           format xxx.xxx.xxx.xxx:xxxx, after changing the proxy remember to make sure the "auto-update" checkbox is checked on the RateAgent manager so that this will take effect.</td>
          </tr>
          <tr>
            <td width="50%" class="boxtitle">&nbsp;</td>
            <td width="50%" class="boxtitle">&nbsp;</td>
          </tr>
          <tr>
            <td width="50%" class="boxtitle">Proxy Group #1 Current Value:</td>
            <td width="50%" class="boxtitle"><%=strProxyValue1 %>&nbsp;</td>
          </tr>
          <tr>
            <td width="50%" class="boxtitle">New Value:</td>
            <td width="50%"><span class=explorer><input maxlength=50 name=proxy1 size="30" value='<%=strProxyValue1 %>'></span></td>
          </tr>
          <tr>
            <td width="50%" class="boxtitle">Proxy Group #2 Current Value:</td>
            <td width="50%" class="boxtitle"><%=strProxyValue2 %>&nbsp;</td>
          </tr>
          <tr>
            <td width="50%" class="boxtitle">New Value:</td>
            <td width="50%"><span class=explorer><input maxlength=50 name=proxy2 size="30" value='<%=strProxyValue2 %>'></span></td>
          </tr>
          <tr>
            <td width="50%" class="boxtitle">Proxy Group #3 Current Value:</td>
            <td width="50%" class="boxtitle"><%=strProxyValue3 %>&nbsp;</td>
          </tr>
          <tr>
            <td width="50%" class="boxtitle">New Value:</td>
            <td width="50%"><span class=explorer><input maxlength=50 name=proxy3 size="30" value='<%=strProxyValue3 %>'></span></td>
          </tr>
          <tr>
            <td width="50%" class="boxtitle">Proxy Group #4 Current Value:</td>
            <td width="50%" class="boxtitle"><%=strProxyValue4 %>&nbsp;</td>
          </tr>
          <tr>
            <td width="50%" class="boxtitle">New Value:</td>
            <td width="50%"><span class=explorer><input maxlength=50 name=proxy4 size="30" value='<%=strProxyValue4 %>'></span></td>
          </tr>
        </table>
        </div>
        <p>&nbsp;<p>&nbsp;<span class="boxtitle">&nbsp;</span><span class=explorer>
        </span>&nbsp;&nbsp; </p>
          <p align="center">
          <input type=submit value='Update' name=submit caption="Add To Database" >
        </p>
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
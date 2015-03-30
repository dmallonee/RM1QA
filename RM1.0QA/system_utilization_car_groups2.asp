<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<% 	Response.Expires = -1  
   	Response.cachecontrol="private" 
   	Response.AddHeader "pragma", "no-cache" 
   
   	on error resume next

   	Server.ScriptTimeout = 180

	Dim strSelected 

    strUserId =    Request.Cookies("rate-monitor.com")("user_id")
	strCityCd =    Request.Form("city_cd")
	strCarTypeCd = Request.Form("car_type_cd")
	
	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "user_city_select"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",  3, 1,  0, strUserId)
		
	Set adoRS = adoCmd.Execute

	adoCmd.CommandText = "car_type_select"
		
	Set adoRS1 = adoCmd.Execute
	Set adoRS2 = adoCmd.Execute
		
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
	
	strCarTypes = ""
  
	If (strCityCd <> "") And (strCarTypeCd <> "") Then
		Rem Get the data sources
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "utilization_car_group_select"
		adoCmd.CommandType = adCmdStoredProc

		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",      3, 1,  0, strUserId)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",     200, 1, 6, strCityCd)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd", 200, 1, 4, strCarTypeCd)

		Set adoRS3 = adoCmd.Execute
		
		While adoRS3.EOF = False
			strCarTypes = strCarTypes & adoRS3.Fields("car_type_cd") & ","
			adoRS3.MoveNext
		
		Wend
	
	End If

  
%>    
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">

<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; Utilization Settings</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script language="JavaScript" src="inc/sitewide.js" ></script>
<script language="JavaScript" src="inc/multiple_select_support2.js"></script>
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

function update_car_type_list() { 

	var list = new String("")
	
	for (i = 0; i < document.car_group_detail.selected_car_types.options.length;i++)
	{
		if (i == 0)
		{
	      list = document.car_group_detail.selected_car_types.options[i].value
		}
		else
		{
	      list = list + ',' + document.car_group_detail.selected_car_types.options[i].value
		}
	}

    document.car_group_detail.car_type_list.value = list 
    //document.search_criteria.selected_companies.options.text; 
     
} 


function removeText(what){
	var what=document.getElementById(what);
	whatChild=what.removeChild(what.childNodes[0]);
}

function replaceText(what,hlaska){
	removeText(what);
	var newText=document.createTextNode(hlaska);
	document.getElementById(what).appendChild(newText);
}

function update_city_cd() {
	replaceText('city_cd', document.getElementById('grp_city_cd').options[document.getElementById('grp_city_cd').selectedIndex].value);
}

function update_grp_car_type() {
	replaceText('car_type_cd', document.getElementById('grp_car_class').options[document.getElementById('grp_car_class').selectedIndex].value);
}

//-->
</script>

<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<style>
<!--
.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.style2 {
	border-collapse: collapse;
}
.style4 {
	border-style: solid;
	border-width: 0;
	text-align: center;
}
.style5 {
	border-style: solid;
	border-width: 0;
	text-align: right;
	font-size: small;
}
.style6 {
	text-align: center;
}
.style7 {
	border-style: solid;
	border-width: 0;
	text-align: left;
	font-size: small;
}
.style8 {
	border-style: solid;
	border-width: 0;
	text-align: left;
}
-->
</style>
<base target="_self">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')" >
<div class="style6">
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
<p>&nbsp;&nbsp;&nbsp; <br>
&nbsp;<font size="2" face="Vendana, Arial, Helvetica, sans-serif">&nbsp;<a href="javascript:not_enabled()">[custom 
city codes]</a>&nbsp;<a href="javascript:not_enabled()">[system status]</a>
<a href="system_proxy.asp">[proxy 
management]</a> <a href="system_utilization.asp">[utilization settings]</a><b> [utilization car groups]</b></font><br>
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
<p>&nbsp;

<FORM METHOD="post" NAME="car_group_detail" action="system_utilization_car_groups_insert.asp"   > 

        		&nbsp;
        		<input type="hidden" name="car_type_list" id="car_type_list" value="<%=strCarTypes %>">

<table border="0" cellpadding="0" bordercolor="#111111" width="500" align="center" class="style2">
                  <tr>
                    <td  class="style5">
                    <font size="2">City:&nbsp;</font></td>
                    <td  class="style4" style="width: 23px">
  			    

        			</td>
                    <td >
  			    <input name="city_cd" id="city_cd" value="<%=strCityCd %>" readonly ></td>
                  </tr>
                      	
                  <tr>
                    <td  class="style7">
                    &nbsp;</td>
                    <td  class="style4" style="width: 23px">
  			    &nbsp;</td>
                    <td  class="style8">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif"><a href="system_utilization_car_groups1.asp">select another 
					city/grp</a></font></td>
                  </tr>
                      	
                  <tr>
                    <td  class="style5">
                    <font size="2">Group Car Type:&nbsp;</font></td>
                    <td  class="style4" style="width: 23px">
  			    &nbsp;</td>
                    <td >

        		<input name="car_type_cd" id="car_type_cd" value="<%=strCarTypeCd %>" readonly ></td>
                  </tr>
                      	
                  <tr>
                    <td  class="style5">
                    &nbsp;</td>
                    <td  class="style4" style="width: 23px">
  			    &nbsp;</td>
                    <td  class="style4">
                    &nbsp;</td>
                  </tr>
                      	
                  <tr>
                    <td colspan="3" class="style4">
                      <font size="2">(use buttons to move between selected/unselected)
                      <br>
                      Unselected types&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Selected types&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font>
                    </td>
                  </tr>
                  <tr>
                    <td class="style4" >
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
                      <select name="unselected_car_types" size="5" style="width:200;"  multiple="multiple">
                   
           				<% While (adoRS2.EOF = False) %>
           					<% If InStr(1, strCarTypes, adoRS2.Fields("car_type_cd").Value) = 0 Then %>
		               			<option value="<%=adoRS2.Fields("car_type_cd").Value %>"><%=adoRS2.Fields("car_type_cd").Value %></option>
		               		<% Else   
		               		    strSelected = strSelected & "<option value=" & adoRS2.Fields("car_type_cd").Value & ">" & adoRS2.Fields("car_type_cd").Value & "</option>"  
							%>
		               		<% End If %>
 	    	   			<% adoRS2.MoveNext %>
               			<% Wend %>
                   
					  </select></font></td>
                    <td class="style4" style="width: 23px">
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
                    <a href="javascript:void(0)" onclick="moveDualList( document.car_group_detail.unselected_car_types, document.car_group_detail.selected_car_types, false );update_car_type_list();return false;">
                      <img border="0" src="images/move_right.GIF" width="24" height="22" alt="Add the highlighted car types"    ></a></font><br>
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
    
                  <a href="javascript:void(0)" onclick="moveDualList( document.car_group_detail.unselected_car_types, document.car_group_detail.selected_car_types, true );update_car_type_list();return false;">
                      <img border="0" src="images/move_right_all.GIF" width="24" height="22"  alt="All all the car types"  ></a></font><br>
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
    
                  <a href="javascript:void(0)" onclick="moveDualList( document.car_group_detail.selected_car_types, document.car_group_detail.unselected_car_types, false );update_car_type_list();return false;">
                      <img border="0" src="images/move_left.GIF" width="24" height="22"  alt="Remove the highlighted car types"  ></a></font><br>
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
    
                    <a href="javascript:void(0)" onclick="moveDualList( document.car_group_detail.selected_car_types, document.car_group_detail.unselected_car_types, true );update_car_type_list();return false;">
                      <img border="0" src="images/move_left_all.GIF" width="24" height="22"  alt="Remove all the car types"  ></a></font></td>                    
                    <td class="style4">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
                    <select name="selected_car_types" size="5" style="width:200;" multiple="multiple">
                    <%=strSelected %>
                 </select></font></td>
                  </tr>
                  <tr>
                    <td style="width: 163px">&nbsp;</td>
                    <td style="width: 23px">&nbsp;</td>                    
                    <td width="116">&nbsp;</td>
                  </tr>
                  <tr>
                    <td colspan="3" class="style4"><input type=submit value='Update' name=submit caption="Add To Database" class="rh_button" ></td>
                  </tr>
         </table>
        		<br>
			    <br>

        

        
        </FORM>

        <p>&nbsp;<p align="center">&nbsp;</p>

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
</div>
</body>
</html>
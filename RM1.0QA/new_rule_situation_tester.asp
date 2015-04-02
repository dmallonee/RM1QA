<%@ Language=VBScript %>
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
	Dim strUserId
	Dim intRuleId
	Dim strAlertDesc
	Dim datBeginDate
	Dim blnSuccess
	Dim blnIsDollar
	
	blnIsDollar = True

    'On Error Resume Next

	'intRuleId = Request.Form("rateruleid")	
	
	strConn = Session("pro_con")
	
	If Request.Form("situation_cd") > 0 Then
	
	    intSituationCd = Request.Form("situation_cd")
	    blnIsDollar = Request.Form("is_dollar")
    	
	    Rem Get the data sources
	    Set adoCmd = CreateObject("ADODB.Command")

	    adoCmd.ActiveConnection =  strConn
	    adoCmd.CommandText = "car_rate_rule_tester"
	    adoCmd.CommandType = 4


	    If IsNumeric(Request.Form("rt_amt")) = False Then
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@rt_amt",        6, 1, 0, 0)
	    Else
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@rt_amt",        6, 1, 0, Request.Form("rt_amt"))
		End If

	    If IsNumeric(Request.Form("situation_amt")) = False Then
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@situation_amt", 6, 1, 0, 0)
	    Else
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@situation_amt", 6, 1, 0, Request.Form("situation_amt"))
		End If

	    If IsNumeric(Request.Form("max_rt_amt")) = False Then
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@max_rt_amt",    6, 1, 0, 0)
	    Else
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@max_rt_amt",    6, 1, 0, Request.Form("max_rt_amt"))
		End If

	    If IsNumeric(Request.Form("min_rt_amt")) = False Then
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@min_rt_amt",    6, 1, 0, 0)
	    Else
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@min_rt_amt",    6, 1, 0, Request.Form("min_rt_amt"))
		End If

	    adoCmd.Parameters.Append adoCmd.CreateParameter("@is_dollar",    11, 1, 0, Request.Form("is_dollar"))
	    adoCmd.Parameters.Append adoCmd.CreateParameter("@situation_cd",  3, 1, 0, Request.Form("situation_cd"))
	    adoCmd.Parameters.Append adoCmd.CreateParameter("@success",      11, 3, 0, Null)


	    adoCmd.Execute

	    blnSuccess = adoCmd.Parameters("@success").Value

        Set adoCmd = Nothing

        If err.number <> 0 Then
	        pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	        response.write "<b>VBScript Errors Occured!<br>"
	        response.write "</b><br>"
	        response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	        response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	        response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	        response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	        response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
        End If


    End If

%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate rule situation tester</title>

<script language="JavaScript">
<!--
function disable(disableIt)
{
	if (disableIt) {
	document.situation_tester.qualifier.disabled = true;
	document.situation_tester.qualifier.selectedIndex = 0; 
	document.situation_tester.situation_amt.value = "";
	document.situation_tester.situation_amt.disabled = true;
	document.situation_tester.situation_amt.bgColor= "#D8DEE1";
	}
	else {
	document.situation_tester.qualifier.disabled = false;
	document.situation_tester.situation_amt.disabled = false;
	document.situation_tester.situation_amt.bgColor= "#000000";
	}
	// = "background-color:#D8DEE1"
}


function IsNumeric(sText)
{
   var ValidChars = "0123456789.";
   var IsNumber=true;
   var Char;

   //alert("sText.length " + sText.value.length);  

   for (i = 0; i < sText.value.length && IsNumber == true; i++) 
      { 
      Char = sText.value.charAt(i);
      //alert(Char); 
      if (ValidChars.indexOf(Char) == -1) 
         {
         IsNumber = false;
         alert("Please enter a valid numeric amount, no percent signs or dollar signs please");
		 sText.focus();
		 }
      }
      
      //alert("done");
   }


//-->
</script>
</head>

<body bgcolor="#D8DEE1" style="font-family: Verdana; font-size: 10pt">

<p>&nbsp;</p>
    <h2>
        Rate Rule Situation Tester</h2>
      <form method="POST" name="situation_tester" action="rule_situation_tester.asp">

  <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="717" id="table1" >
    <tr>
      <td width="160" height="22">&nbsp;</td>
      <td width="557" height="22" colspan="2">
      &nbsp;</td>
    </tr>
    <tr>
      <td width="160" height="22"><font size="2">Max competitive rate:</font></td>
      <td width="557" height="22" colspan="2">
        <p>
        <input type="text" name="max_rt_amt" size="20" style="text-align: right; width:75;"  value="<%=Request.Form("max_rt_amt") %>" ></p></td>
    </tr>
    <tr>
      <td width="160" height="22"><font size="2">Min competitive rate:</font></td>
      <td width="557" height="22" colspan="2">
      <input type="text" name="min_rt_amt" size="20"  style="text-align: right; width:75;"  value="<%=Request.Form("min_rt_amt") %>" id="Text1" onclick="return Text1_onclick()" ></td>
    </tr>
    <tr>
      <td width="160" height="22"><font size="2">Comparison rate:</font></td>
      <td width="557" height="22" colspan="2">
      <input type="text" name="rt_amt" size="20"  style="text-align: right; width:75;"  value="<%=Request.Form("rt_amt") %>" ><span
          style="font-size: 2px"></span></td>
    </tr>
    <tr>
      <td width="160" height="22">&nbsp;</td>
      <td width="557" height="22" colspan="2">
      &nbsp;</td>
    </tr>
    <tr>
<font color="#080000">
      <td width="160" height="22">
<font size="2">Situation:</font></td>
      <td width="557" height="22" colspan="2">
      <font size="2" color="#080000">
		<select size="1" name="situation" style="width:370; font-family:Verdana; font-size:10pt; height:24">
		<option selected value="0">None</option>
		<option value="1000">Alert me if comparison rate is:</option>
		<option value="3000">Alert me if comparison rate is not:</option>
		<option value="2000">Alert me if any competitive rate is:</option>
		</select></font></td>
</font>
    </tr>
    <tr>
<font color="#080000">
      <td width="160" height="23">
<font color="#080000">
      <font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; 
      Amount:</font></font></td>
      <td width="89" height="23">
      <input type="text" name="situation_amt" size="20"  onblur="IsNumeric(document.situation_tester.situation_amt)" style="text-align: right; width:75; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=Request.Form("situation_amt") %>"></td>
      
      
</font>
      <td width="468" height="23" align="left" style="text-align: left">
<font color="#080000">
      <% If blnIsDollar Then %>
      <input type="radio" value="11" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar1" onclick="disable(false)" checked="CHECKED"><font face="Verdana" size="2"><label for="is_dollar1">Dollar </label>
      <input type="radio" value="01" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar2" onclick="disable(false)" ><label for="is_dollar2">Percentage</label>
      <input type="radio" value="01" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar2" onclick="disable(true)"  ><label for="is_dollar2">Equal (Please leave amount blank)</label><br>
      <% Else %>
      <input type="radio" value="12" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar1"><label for="is_dollar1">Dollar </label>
      <input type="radio" value="02" name="is_dollar" style="font-family:Verdana; font-size:10pt"  id="is_dollar2"><label for="is_dollar2">Percentage</label>
      <input type="radio" value="01" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar2"  ><label for="is_dollar2">Equal (Please leave amount blank)</label><br>
      <% End If %></font></td>
      
      
    </tr>
    <tr>
      <td width="160" height="22">
      <font face="Verdana" size="2" color="#080000">&nbsp;&nbsp;&nbsp;Qualifier:</font></td>
      <td width="557" height="22" colspan="2">
      <select size="1" name="qualifier" style="width:225; font-family:Verdana; font-size:10pt; height:22; ">
		<option selected value="0" >   = (equal to) </option>
		<option value="3">  <> (not equal to)</option>
		<option value="1">  &gt; (greater than)</option>
		<option value="2">  < (less than)</option>
		<option value="4">  &gt;= (greater than or equal to)</option>
		<option value="5">  <= (less than or equal to)</option>
		</select></td>
    </tr>
    <tr>
      <td width="160" height="23">
<font color="#080000">
      <font face="Verdana" size="2" color="#080000">&nbsp;&nbsp;&nbsp; 
      </font>
      
      
</font>
    <font size="2" color="#080000">Detail</font><font face="Verdana" size="2" color="#080000">:</font></td>
      <td width="557" height="23" colspan="2">
      <font color="#080000">
      <select size="1" name="detail" style="width:370; font-family:Verdana; font-size:10pt; height:24">
      
     
      <% If intSituationCd = 0 Then	%>
	      <option selected value="0">Lowest competitor</option>
	  <% Else %>
	      <option value="0">Lowest competitor</option>
	  <% End If %>
 
      <% If intSituationCd = 1 Then	%>
	      <option selected value="1">Highest competitor</option>
	  <% Else %>
	      <option value="1">Highest competitor</option>
	  <% End If %>

      <% If intSituationCd = 2 Then	%>
	      <option selected value="2">All competitors</option>
	  <% Else %>
	      <option value="2">All competitors</option>
	  <% End If %>

      <% If intSituationCd = 3 Then	%>
	      <option selected value="3">Any competitors</option>
	  <% Else %>
	      <option value="3">Any competitors</option>
	  <% End If %>


      </select></font></td>
      
      
    </tr>
    <tr>
      <td width="160" height="23">
&nbsp;</td>
      <td width="557" height="23" colspan="2">
      &nbsp;</td>
      
      
    </tr>
    <tr>
      <td width="160" height="23">
&nbsp;</td>
      <td width="557" height="23" colspan="2">
      <input type="submit" value="Evalute Situation" name="submit"></td>
      
      
    </tr>
  </table>
  </form> 
    <br />
    <% If Request.Form("situation_cd") > 0 Then %>
    
    <% If IsNull(blnSuccess) Then %>
    <span style="font-size: 11pt;"><strong>ERROR</strong></span> --
    This situation is not functioning properly, please do not use it.<br />
    
    <% ElseIf blnSuccess Then %>
    <span style="font-size: 11pt; color: #009900"><strong>TRUE</strong></span> --
    The situation is true, the response will be performed, and if there is a follow-on rule for the true
    outcome of the situation, it will be performed.
    <% Else %>
    <br />
    <span style="font-size: 11pt; color: #cc0000"><strong>FALSE</strong></span> --
    The situation is false, the response will not be performed, and if there is a follow-on
    rule for
    the false outcome of the situation, it will be performed.<br />
    <% End If %>
    
    <% End If %>
<br><font size="1">rule id: <%= Request.Form("situation_cd") %></font>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
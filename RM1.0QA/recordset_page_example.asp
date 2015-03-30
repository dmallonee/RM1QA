<%
Const adOpenKeyset = 1
Const adLockReadOnly = 1
Dim strQuery   'to hold our query string

If Request("Action") = "New Query" Then
   Response.Redirect("ado_sample.htm")
   Response.End
End If

'Build up the query string
strQuery =
 BuildQuery()
If strQuery = "" Then
  Response.Redirect("ado_sample.htm")
  Response.End
End If

'Create a connection object to execute the query
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = "provider=msidxs"
objConn.Open
Set objRS = Server.CreateObject("ADODB.RecordSet")
objRS.Open Query, objConn, adOpenKeyset, adLockReadOnly
If objRS.EOF Then
  Response.Write("No records found!")
  Set objRS = Nothing
  objConn.Close
  Set objConn = Nothing
  Response.End
End If

'Set the page number - each page holds five records
objRS.PageSize = 5
Scroll = Request("Scroll")
If Scroll <> "" Then
  Page = mid(Scroll, 5)
  If Page < 1 Then Page = 1
Else
  Page = 1
End If
objRS.AbsolutePage = Page
%>


<HTML>
<HEAD><TITLE>Paged ADO Example</TITLE>

<SCRIPT LANGUAGE=VBScript RUNAT=Server>
Function BuildQuery()
  SQL = "SELECT Filename, Size, Vpath, Path, Write, Characterization FROM "
  If Request("Scope") = "" Then
    SQL = SQL & "SCOPE() "
  Else
    SQL = SQL & "SCOPE('"
    If Request("Depth") = "Shallow" Then
      SQL = SQL & "SHALLOW TRAVERSAL OF " & """" & Request("Scope")
      SQL = SQL & """" & "'" & ")"
    Else
      SQL = SQL & "DEEP TRAVERSAL OF " & """" & Request("Scope")
      SQL = SQL & """" & "'" & ")"
    End If
  End if
  If Request("WHERE") = "Content" Then
    SQL = SQL & " WHERE CONTAINS(" & "'" & Request("Criteria") & "'" & ") > 0"
    BuildQuery = SQL
  ElseIf Request("WHERE") = "Size" Then
    SQL = SQL & " WHERE " & Request("Where") & Request("Operator")
    SQL = SQL & Request("Criteria")
    BuildQuery = SQL
  Else
    SQL = SQL & " WHERE " & Request("Where") & Request("Operator")
    SQL = SQL & " '" & Request("Criteria") & "'"
    BuildQuery = SQL
  End If
End Function
</SCRIPT> 
</head>
<BODY>
<H3>Your query returned the following results:</H3>
<TABLE border="0" width="100%" height="66">

<% RowCount = objRS.PageSize %>
<% Do While Not objRS.EOF And RowCount > 0 %>

 <TR>
  <TD width="20%" align="right"><B>Virtual Path:</B></TD>
  <TD width="80%"><%= objRS("vPath")%></TD>
 </TR>
 <TR>
  <TD width="20%" align="right"><b><strong>Physical Path:</B></TD>
  <TD width="80%"><%= objRS("Path")%></TD>
 </TR>
 <TR>
  <TD width="20%" align="right"><B><strong>Filename:</B></TD>
  <TD width="80%"><%= objRS("Filename")%></TD>
 </TR>
 <TR>
  <TD width="20%" align="right"><B><strong>Size:</B></TD>
  <TD width="80%"><%= objRS("Size") & " bytes"%></TD>
 </TR>
 <TR>
  <TD width="20%" align="right"><B><strong>Last Modified:</B></tD>
  <TD width="80%"><%= objRS("Write")%></TD>
 </TR>
 <TR>
  <TD width="20%" align="right"><B><strong>Excerpt:</B></TD>
  <TD width="80%"><%= objRS("Characterization")%></TD>
 </TR>

<% RowCount = RowCount - 1 %>
<% objRS.MoveNext %>
<% Loop %>

</TABLE>

<% Set objRS = Nothing %>
<% objConn.Close %>
<% Set objConn = Nothing %>
<FORM METHOD="POST" ACTION="ADO_SAMPLE.ASP">
<INPUT TYPE="SUBMIT" NAME="ACTION" VALUE="New Query">
<INPUT TYPE="HIDDEN" NAME="Scope" VALUE="<%=Request("Scope")%>">
<INPUT TYPE="HIDDEN" NAME="Depth" VALUE="<%=Request("Depth")%>">
<INPUT TYPE="HIDDEN" NAME="Criteria" VALUE="<%=Request("Criteria")%>">
<INPUT TYPE="HIDDEN" NAME="Operator" VALUE="<%=Request("Operator")%>">
<INPUT TYPE="HIDDEN" NAME="Where" VALUE="<%=Request("Where")%>">

<% If Page > 1 Then %>
  <INPUT TYPE="SUBMIT" NAME="Scroll" VALUE="<%="Page " & Page - 1 %>">
<% End If %>
<% If RowCount = 0 Then %>
  <INPUT TYPE="SUBMIT" NAME="Scroll" VALUE="<%="Page " & Page + 1 %>">
<% End If %>

</FORM>
<!--#INCLUDE FILE="footer.asp"-->
</BODY>
</HTML>

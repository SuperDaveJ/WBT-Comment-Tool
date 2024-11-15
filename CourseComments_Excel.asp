<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Review Comments</title>
<!-- #include file="adovbs.inc" -->
<%
Session.Timeout = 60

'Set variables here
courseID = Request.QueryString("cID")

set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open(Session("strConnectSQL"))

set objCmd = Server.CreateObject("ADODB.Command")
set objRS = Server.CreateObject("ADODB.Recordset")
With objCmd
	.ActiveConnection = objConn
	.CommandType = adCmdStoredProc
	.CommandText = "usp_getReviewFeedbackByCourse"
	.Parameters.Append .CreateParameter("@courseID", adVarChar, adParamInput, 30)
	.Parameters("@courseID").Value = courseID
	Set objRS = .Execute()
End With

%>

</head>

<body class="oneColElsCtr">
<div id="container">
  <div id="mainContent">

<p style="text-align:center;font-weight:bold;">Comments for: <%=courseID%> course. <br /></p>

<%
'output to excel format for download
	Response.ContentType = "application/vnd.ms-excel"
	'********* file name code below does not work *************
	'fileName = "slim_m" & moduleID & "l" & lessonID & "_comments.xls"
	'Response.AddHeader "Content-Disposition", "filename=" & fileName
%>

<table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#333333">
    <tr>
    	<th width="50px">Module</th>
        <th width="50px">Lesson</th>
        <th width="50px">Page</th>
        <th width="50px">Category</th>
        <th>Comment</th>
        <th>Reviewer</th>
        <th>Date/Time</th>
    </tr>
<% If (NOT objRS.EOF) Then 
	DO WHILE NOT objRS.EOF
	If (objRS.Fields(5) <> "NA") Then
		userName = objRS.Fields(5)
	Else
		userName = "NA"
	End If
%>
    <tr>
      <td><%=objRS.Fields(0)%></td>
      <td><%=objRS.Fields(1)%></td>
      <td><%=objRS.Fields(2)%></td>
      <td><%=objRS.Fields(3)%></td>
      <td><%=objRS.Fields(4)%></td>
      <td><%=userName%></td>
      <td><%=objRS.Fields(6)%></td>
    </tr>
<% 
		objRS.MoveNext()
	LOOP
Else
%>
    <tr>
      <td colspan="5">No data found.</td>
    </tr>
<% End If %>

    </table>

<%

objRS.Close
Set objRS = Nothing
Set objCmd = Nothing
objConn.Close
Set objConn = Nothing

%>

<form>
	<div id="submit"> 
        <input type="button" value="Close" onclick="javascript:self.close();" title="Close" />
    </div>
</form>

	<!-- end #mainContent --></div>
<!-- end #container --></div>
</body>
</html>

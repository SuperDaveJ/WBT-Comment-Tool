<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Review Comments</title>
<link href="virtualPilot.css" rel="stylesheet" type="text/css" />
<!-- #include file="adovbs.inc" -->
<%
Session.Timeout = 60

'Set variables here
courseID = Request.QueryString("cID")
moduleID = Request.QueryString("mID")
thisPage = "moduleComments_excel.asp?cID=" & courseID & "&mID=" & moduleID

set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open(Session("strConnectSQL"))

set objCmd = Server.CreateObject("ADODB.Command")
set objRS = Server.CreateObject("ADODB.Recordset")
With objCmd
	.ActiveConnection = objConn
	.CommandType = adCmdStoredProc
	.CommandText = "usp_getReviewFeedbackByModule"
	.Parameters.Append .CreateParameter("@courseID", adVarChar, adParamInput, 30)
	.Parameters("@courseID").Value = courseID
	.Parameters.Append .CreateParameter("@moduleID", adVarChar, adParamInput, 50)
	.Parameters("@moduleID").Value = moduleID
	Set objRS = .Execute()
End With

%>

</head>

<body class="oneColElsCtr">
<div id="container">
  <div id="mainContent">

<div style="height:20px;">&nbsp;</div>
<h2> COMMENTS </h2>
<p style="text-align:center;font-weight:bold;">This is for: <%=courseID%>, Module <%=moduleID%>. </p>

<table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#333333">
    <tr>
        <th width="50px">Lesson</th>
        <th width="50px">Page Number</th>
        <th width="80px">Category</th>
        <th>Comment</th>
        <th>Reviewer</th>
        <th>Date/Time</th>
    </tr>
<% If (NOT objRS.EOF) Then 
	DO WHILE NOT objRS.EOF
	If (objRS.Fields(4) <> "NA") Then
		userName = objRS.Fields(4)
	Else
		userName = "NA"
	End If
%>
    <tr>
      <td><%=objRS.Fields(0)%></td>
      <td><%=objRS.Fields(1)%></td>
      <td><%=objRS.Fields(2)%></td>
      <td><%=objRS.Fields(3)%></td>
      <td><%=userName%></td>
      <td><%=objRS.Fields(5)%></td>
    </tr>
<% 
		objRS.MoveNext()
	LOOP
Else
%>
    <tr>
      <td colspan="6">No data found.</td>
    </tr>
<% End If %>

    </table>

<!--
<form action="<%=thisPage%>" method="get">
	<div id="send" style="text-align:center"> 
    	<input type="submit" value="Download this file" /> &nbsp;&nbsp;&nbsp;
        <input type="button" value="Close" onclick="javascript:self.close();" title="Close" />
    </div>
</form>
-->

<%

objRS.Close
Set objRS = Nothing
Set objCmd = Nothing
objConn.Close
Set objConn = Nothing
%>

    <p style="margin-top:2em; text-align:center; clear:both;">
        <a href="<%=thisPage%>" target="_blank">Download to Excel Format</a>
    </p>
    <p style="margin-top:2em; text-align:center; clear:both;">
        <a href="javascript:self.close()">Close this window</a>
    </p>
	<p>&nbsp;</p>

	<!-- end #mainContent --></div>
<!-- end #container --></div>
</body>
</html>

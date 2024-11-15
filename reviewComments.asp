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
userID = Request.QueryString("uID")
courseID = Request.QueryString("cID")
moduleID = Request.QueryString("mID")
lessonID = CInt(Request.QueryString("lID"))
pageID = Request.QueryString("pID")

If (moduleID <> "") Then
	pagePath = "<b>" & courseID & "</b>, moduleID: <b>" & moduleID & "</b>, Lesson: <b>" & lessonID & "</b>, Page: <b>" & pageID & "</b>."
Else
	pagePath = "<b>" & courseID & "</b>, Lesson: <b>" & lessonID & "</b>, Page: <b>" & pageID & "</b>."
End If

set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open(Session("strConnectSQL"))

set objCmd = Server.CreateObject("ADODB.Command")
set objRS = Server.CreateObject("ADODB.Recordset")
With objCmd
	.ActiveConnection = objConn
	.CommandType = adCmdStoredProc
	.CommandText = "usp_GetFeedbackForThisPage"
	.Parameters.Append .CreateParameter("@courseID", adVarChar, adParamInput, 30)
	.Parameters("@courseID").Value = courseID
	.Parameters.Append .CreateParameter("@moduleID", adVarChar, adParamInput, 50)
	If ( moduleID = "" ) Then
		.Parameters("@moduleID").Value = "NA"
	Else
		.Parameters("@moduleID").Value = moduleID
	End If
	.Parameters.Append .CreateParameter("@lessonID", adInteger, adParamInput)
	.Parameters("@lessonID").Value = lessonID
	.Parameters.Append .CreateParameter("@pageID", adVarChar, adParamInput, 30)
	.Parameters("@pageID").Value = pageID
	Set objRS = .Execute()
End With

%>

</head>

<body class="oneColElsCtr">
<div id="container">
  <div id="mainContent">

<div style="height:20px;">&nbsp;</div>
<h2> REVIEW COMMENTS </h2>
<p style="text-align:center">This is for: <%=pagePath%></p>

<table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#333333">
    <tr>
        <th width="100px">Category</th>
        <th>Comment</th>
        <th>Reviewer</th>
        <th>Date</th>
    </tr>
<% If (NOT objRS.EOF) Then 
	DO WHILE NOT objRS.EOF
	If (objRS.Fields(2) <> "NA") Then
		userName = objRS.Fields(2)
	Else
		userName = objRS.Fields(3)
	End If
%>
    <tr>
      <td><%=objRS.Fields(0)%></td>
      <td><%=objRS.Fields(1)%></td>
      <td><%=userName%></td>
      <td><%=objRS.Fields(4)%></td>
    </tr>
<% 
		objRS.MoveNext()
	LOOP
Else
%>
    <tr>
      <td colspan="4">No data found.</td>
    </tr>
<% End If %>

    </table>

<form>
	<div id="submit"> 
        <input type="button" value="Close" onclick="javascript:self.close();" title="Close" />
    </div>
</form>
<%
objRS.Close
Set objRS = Nothing
Set objCmd = Nothing
objConn.Close
Set objConn = Nothing
%>
	<!-- end #mainContent --></div>
<!-- end #container --></div>
</body>
</html>

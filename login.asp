<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Login</title>
<link href="virtualPilot.css" rel="stylesheet" type="text/css" />
<!-- #include file="adovbs.inc" -->
<%
 courseID = "1752"
 'strConn = "Provider=SQLOLEDB;Data Source=OKCOK06WBS01;Initial Catalog=PilotEvalTool;User ID=PrivateWebUserS;Password=SresUbeWetavirP;"
 strConn = Session("strConnectSQL")
 Session("cID") = courseID
%>
<style type="text/css">
<!--
-->
</style></head>

<body class="oneColElsCtr">
<%
strPostTarget = Trim(Request("URL"))
Dim objCmd, objConn, objRS
set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open(strConn)
set objCmd = Server.CreateObject("ADODB.Command")
set objRS = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM tblCourseList WHERE courseID='" & courseID & "'"
With objCmd
	.ActiveConnection = objConn
	.CommandType = adCmdText
	.CommandText = strSQL
	Set objRS = .Execute()
End With

If (NOT objRS.EOF) Then
	cTitle = objRS("courseTitle")
	Session("cTitle") = cTitle
	Session("cURL") = objRS("courseURL")
	Session("gPage") = objRS("gInfoFile")
End If

objRS.Close()
'Set objRS = Nothing

Temp = CStr(Request("Submit"))
Session("uID") = ""

errMsg = ""
If (Temp = "Submit") Then
	userID = LCase(request.form("userID"))
	password = LCase(request.form("pswd"))
	
	strSQL = "SELECT userID, password FROM tblUsers WHERE courseID='" & courseID & "' AND userID='" & userID & "'"
	objCmd.CommandText = strSQL
	Set objRS = objCmd.Execute()
	If (NOT objRS.EOF ) Then
		If ( password = objRS("password") ) Then
			Session("uID") = userID
			if (userID = "admin") then
				'if admin skip survey
				Response.Redirect(Session("cURL"))
			else
				Response.Redirect(Session("gPage"))
			end if
		Else
			errMsg = "Your password is incorrect. Please re-enter your correct password and try again."
		End If
	Else
		errMsg = "Your record not found. Please make sure that you entered your ID and password correctly."
	End If
End If
%>

<div id="container">
  <div id="mainContent">

	<div style="height:50px;">&nbsp;</div>
    <h2> Login </h2>
    <p style="text-align:center">Please log in to review the "<%=cTitle%>" course.</p>

	<form action="<%=strPostTarget%>" method="post" name="login">
    <table width="60%" align="center">
    <tr>
        <td width="120px" align="right">
            User Name:
        </td>
        <td>
            <input name="userID" class="textEntry" type="text" size="40" maxlength="50" />
        </td>
    </tr>
    <tr>
        <td align="right">
            Password:
        </td>
        <td>
            <input name="pswd" class="textEntry" type="password" size="40" maxlength="50" />
        </td>
    </tr>
    <tr>
        <td colspan="2" align="center">&nbsp;
            
        </td>
    </tr>
    <tr>
        <td colspan="2" align="center">
            <input name="submit" type="submit" value="Submit" />
        </td>
    </tr>
    <tr>
        <td colspan="2" align="center">&nbsp;
            
        </td>
    </tr>
    <tr>
        <td colspan="2" align="center">
            <%=errMsg%>
        </td>
    </tr>
    </table>
    </form>

	<!-- end #mainContent --></div>
<!-- end #container --></div>
</body>
</html>

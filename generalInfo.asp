<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>General Information</title>
<link href="virtualPilot.css" rel="stylesheet" type="text/css" />
<!-- #include file="adovbs.inc" -->
<%
 strConn = Session("strConnectSQL")
 courseID = Session("cID")
 uID = Session("uID")
 cTitle = Session("cTitle")
%>
<style type="text/css">
<!--
-->
</style></head>

<body class="oneColElsCtr">
<%
nItems = 10

strPostTarget = Trim(Request("URL"))
pgPostback = Request.Form("pgPostback")
pgRedirect = Session("cURL") & "?pilot=true&uID=" & uID & "&cID=" & courseID

' *** If postback
If (pgPostback = "yes") Then
	'get values for database
	Dim Temp
	Dim arrQId, arrQAns
	ReDim arrQId(nItems)
	ReDim arrQAns(nItems)

	arrQId(1) = "name"
	arrQAns(1) = Request.Form("name")
	arrQId(2) = "position"
	arrQAns(2) = Request.Form("position")
	arrQId(3) = "company"
	arrQAns(3) = Request.Form("company")
	arrQId(4) = "browser"
	Temp = Request.Form("browser")
	if (Temp = "") then
		Temp = 0
	end if
	arrQAns(4) = Temp
	arrQId(5) = "otherBrowser"
	arrQAns(5) = Request.Form("otherBrowser")
	
	For Itemp=6 to nItems
		Temp = Request.Form("q" & Itemp)
		if (Temp = "") then
			Temp = 0
		end if
		arrQId(Itemp) = "q" & Itemp
		arrQAns(Itemp) = Temp
	Next
End If

set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open(Session("strConnectSQL"))

' ***** Check if there are any user data on this page *****
set rsUserAnswer = Server.CreateObject("ADODB.Recordset")
set objCmd = Server.CreateObject("ADODB.Command")

With objCmd
	.ActiveConnection = objConn
	.CommandType = adCmdText
	.CommandText = "SELECT * FROM tblUserProfile WHERE courseID='" & courseID & "' AND userID='" & uID & "'"
End With

Set rsUserAnswer = objCmd.Execute()

If (NOT rsUserAnswer.EOF) Then
	rsUserAnswer.close
	Set rsUserAnswer = Nothing
	set objCmd = Nothing
	objConn.Close()
	Set objConn = nothing
	response.redirect(pgRedirect)
	Response.End
Else
	rsUserAnswer.close
	Set rsUserAnswer = Nothing
	If (pgPostback = "yes") Then
		' redirect
		manageUserData("insert")
		objConn.Close()
		Set objConn = nothing
		Response.Redirect(pgRedirect)
		response.end
	End If
End If

Sub manageUserData(whatToDo)
	'whatToDo is either insert or update
	set objCmd = Server.CreateObject("ADODB.Command")
	With objCmd
		.ActiveConnection = objConn
		.CommandText = "usp_ManageGeneralInfo"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter("@whatToDo", adVarChar, adParamInput, 15)
		.Parameters("@whatToDo").Value = whatToDo
		.Parameters.Append .CreateParameter("@courseID", adVarChar, adParamInput, 30)
		.Parameters("@courseID").Value = courseID
		.Parameters.Append .CreateParameter("@userID", adVarChar, adParamInput, 50)
		.Parameters("@userID").Value = Session("uID")
		.Parameters.Append .CreateParameter("@qID", adVarChar, adParamInput, 20)
		.Parameters.Append .CreateParameter("@qAnswer", adVarChar, adParamInput, 200)
	End With

	For Itemp=1 to nItems
		'assign parameter values
		objCmd.Parameters("@qID").Value = arrQId(Itemp)
		objCmd.Parameters("@qAnswer").Value = arrQAns(Itemp)
		objCmd.Execute()
	Next
	set objCmd = Nothing
End Sub

objConn.Close
set objConn = Nothing

%>

<div id="container">
  <div id="mainContent">

	<div style="height:20px;">&nbsp;</div>

	<form action="<%=strPostTarget%>" method="post" name="login">
    <h3>General Information</h3>
    
    <p>Thank you for participating in the virtual pilot of the Web-based training course, Planning for the Needs of Children in Disasters.  Your participation and the feedback you provide are crucial to the successful development of the course.</p>
    <p>The purposes of this pilot test are to:</p>
    <ul>
        <li>Determine the target audience’s overall satisfaction/dissatisfaction with the course.</li>
        <li>Determine if the course features are working as designed.</li>
        <li>Identify any problems with functionality or content in each lesson.</li>
        <li>Identify additional resources that should be included in the toolkit.</li>
    </ul>
    <p>Before you begin the course, please answer a few questions that will help us understand our audience.</p>
        <table style="width: 100%;" cellpadding="5">
            <tr>
                <td width="160px">
                    Your name:
                </td>
                <td>
                    <input name="name" type="text" size="80" />
                </td>
            </tr>
            <tr>
                <td>
                    Your current position:
                </td>
                <td>
                    <input name="position" type="text" size="80" />
                </td>
            </tr>
            <tr>
                <td>
                    Company/Organization:
                </td>
                <td>
                    <input name="company" type="text" size="80" />
                </td>
            </tr>
        </table>
        <br />
        <p>What browser are you using?</p>
        <table width="100%" cellpadding="0">
        <tr>
        	<td width="20px">&nbsp;</td>
        	<td width="100px">
        		<input name="browser" type="radio" value="IE 6" /> IE 6
            </td>
        	<td width="100px">
        		<input name="browser" type="radio" value="IE 7" /> IE 7
            </td>
        	<td width="100px">
        		<input name="browser" type="radio" value="IE 8" /> IE 8
            </td>
        	<td width="100px">
        		<input name="browser" type="radio" value="Other" /> Other
            </td>
        	<td>
        		<input name="otherBrowser" type="text" size="50" />
            </td>
		</tr>
       </table>     
       <p>To find this information in Internet Explorer, click "Help" in your browser window, and then click "About Internet Explorer."
       </p>
       
 	<br />
 	<div class="question">       
       <p>Describe the level at which you are currently involved in emergency operations:</p>
       <div class="distracter">
       		<input type="radio" name="q6" value="State government" /> State government
       		<br /><input type="radio" name="q6" value="Local or tribal government" /> Local or tribal government
       		<br /><input type="radio" name="q6" value="Emergency response (e.g., fire department, EMS)" /> Emergency response (e.g., fire department, EMS)
       		<br /><input type="radio" name="q6" value="Community organization (e.g., school, library)" /> Community organization (e.g., school, library)
       		<br /><input type="radio" name="q6" value="Private business (e.g., child care center, retail store)" /> Private business (e.g., child care center, retail store)
       		<br /><input type="radio" name="q6" value="Other" /> Other
        </div>
	</div>
       
 	<div class="question">       
       <p>What is your role in emergency planning <u>for your community</u>?</p>
       <div class="distracter">
       		<input type="radio" name="q7" value="Primary decision-maker" /> Primary decision-maker
       		<br /><input type="radio" name="q7" value="Actively involved in decision-making" /> Actively involved in decision-making
       		<br /><input type="radio" name="q7" value="Somewhat involved in decision-making" /> Somewhat involved in decision-making
       		<br /><input type="radio" name="q7" value="Included in drills and exercises" /> Included in drills and exercises
       		<br /><input type="radio" name="q7" value="Not involved in community planning efforts" /> Not involved in community planning efforts
        </div>
	</div>
       
 	<div class="question">       
       <p>What is your role in emergency planning <u>for your organization</u>?</p>
       <div class="distracter">
       		<input type="radio" name="q8" value="Primary decision-maker" /> Primary decision-maker
       		<br /><input type="radio" name="q8" value="Actively involved in decision-making" /> Actively involved in decision-making
       		<br /><input type="radio" name="q8" value="Somewhat involved in decision-making" /> Somewhat involved in decision-making
       		<br /><input type="radio" name="q8" value="Included in drills and exercises" /> Included in drills and exercises
       		<br /><input type="radio" name="q8" value="Not involved in organization planning efforts" /> Not involved in organization planning efforts
        </div>
	</div>
       
 	<div class="question">       
       <p>How long have you worked with children or in a field directly related to children’s needs?</p>
       <div class="distracter">
       		<input type="radio" name="q9" value="Never" /> Never
       		<br /><input type="radio" name="q9" value="Less than six months" /> Less than six months
       		<br /><input type="radio" name="q9" value="Six months to one year" /> Six months to one year
       		<br /><input type="radio" name="q9" value="One year to five years" /> One year to five years
       		<br /><input type="radio" name="q9" value="Five to ten years" /> Five to ten years
       		<br /><input type="radio" name="q9" value="Over ten years" /> Over ten years
        </div>
	</div>
       
 	<div class="question">       
       <p>How would you rate your community’s disaster planning efforts in regards to the needs of children?</p>
       <div class="distracter">
       		<input type="radio" name="q10" value="Excellent" /> Excellent
       		<br /><input type="radio" name="q10" value="Adequate" /> Adequate
       		<br /><input type="radio" name="q10" value="Improvement needed" /> Improvement needed
       		<br /><input type="radio" name="q10" value="Don’t know" /> Don’t know
        </div>
	</div>
       
	<div id="submit"> 
        <input name="submit" type="submit" value="Submit" />
        <input name="pgPostback" type="hidden" value="yes" />
    </div>
    </form>

	<!-- end #mainContent --></div>
<!-- end #container --></div>
</body>
</html>

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add Comment</title>
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
'reviewer is used for no login case.  That means no userID available.
reviewer = Request.QueryString("reviewer")

If (moduleID <> "") Then
	pagePath = "<b>" & courseID & "</b>, Module: <b>" & moduleID & "</b>, Lesson: <b>" & lessonID & "</b>, Page: <b>" & pageID & "</b>."
Else
	pagePath = "<b>" & courseID & "</b>, Lesson: <b>" & lessonID & "</b>, Page: <b>" & pageID & "</b>."
End If

strPostTarget = CStr(Request("URL")) & "?uID=" & userID & "&cID=" & courseID & "&mID=" & moduleID & "&lID=" & lessonID & "&pID=" & pageID
pgPostback = Request.Form("pgPostback")

If (pgPostback = "yes") Then
	Dim strComment, strCategory
	strComment = ""
	strCategory = ""
    If (userID = "NA") Then
		reviewerName = Trim(Request.Form("reviewerName"))
    End If
	strComment = Trim(Request.Form("comment"))
	If (Request.Form("text") = "on") Then
		strCategory = strCategory & "Text"
	End If
	If (Request.Form("graphic") = "on") Then
		if (strCategory = "") then
			strCategory = "Graphic"
		else
			strCategory = strCategory & ", Graphic"
		end if
	End If
	If (Request.Form("audio") = "on") Then
		if (strCategory = "") then
			strCategory = "Audio"
		else
			strCategory = strCategory & ", Audio"
		end if
	End If
	If (Request.Form("other") = "on") Then
		if (strCategory = "") then
			strCategory = "Other"
		else
			strCategory = strCategory & ", Other"
		end if
	End If
	
	set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open(Session("strConnectSQL"))
	
	set objCmd = Server.CreateObject("ADODB.Command")
	'****************** Additional Comment (one per lesson) ************************
	With objCmd
		.ActiveConnection = objConn
		.CommandText = "usp_InsertReviewFeedback"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter("@userID", adVarChar, adParamInput, 50)
		.Parameters.Append .CreateParameter("@reviewerName", adVarChar, adParamInput, 50)
		.Parameters("@userID").Value = userID
		If (userID = "NA") Then
			.Parameters("@reviewerName").Value = reviewerName
		Else
			.Parameters("@reviewerName").Value = "NA"
		End If
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
		.Parameters.Append .CreateParameter("@cat", adVarChar, adParamInput, 50)
		.Parameters("@cat").Value = strCategory
		.Parameters.Append .CreateParameter("@comment", adVarChar, adParamInput, -1)
		.Parameters("@comment").Value = strComment
		.Execute()
	End With
	
	Set objCmd = Nothing
	objConn.Close
	Set objConn = Nothing
	
	Response.End()
End If
%>

<script>
function checkForm() {
    alertMsg = "";
    
    if (document.addComment.reviewerName.value.length < 2) {        
        alertMsg += "Please enter your name in the Reviewer field.\n";
        document.addComment.comment.focus();
    }
	 
    if (document.addComment.comment.value.length < 3) {        
        alertMsg += "Please enter your comment in the Comment field.\n";
        document.addComment.comment.focus();
    }
	 
    if (!document.addComment.text.checked && !document.addComment.graphic.checked && !document.addComment.audio.checked && !document.addComment.other.checked) {        
        alertMsg += "Please select a category for your comment.\n";
    }

    if (alertMsg != "") { 
		alert(alertMsg); 
		return false; 
	} else return true;
}

function setFocus() {
	document.addComment.reviewerName.focus();
}

</script>

</head>

<body class="oneColElsCtr" onload="setFocus();" onunload="self.close();">

<div id="container">
  <div id="mainContent">

<div style="height:20px;">&nbsp;</div>
<h2> ADD COMMENT </h2>
<p>Module ID: <%=moduleID%></p>

<form name="addComment" action="<%=strPostTarget%>" method="post">
    <table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#333333">
        <tr>
        	<td>Adding comment for <%=pagePath%></td>
        </tr>
        <% If (userID = "NA") Then %>
        <tr>
        	<td>Reviewer: <input type="text" name="reviewerName" size="40" /></td>
        </tr>
        <% End If %>
        <tr>
        	<td>Please enter the following information:</td>
        </tr>
    </table><br>
     
    <table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#333333">
        <tr>
            <td width="100px">Comment:</td>
            <td><textarea name="comment" cols="65" rows="10"></textarea></td>
        </tr>
        <tr>
          <td>Category:</td>
          <td><div style="width:120px; text-align:left; float:left;"><input type=checkbox name="text">Text</div>
            <div style="width:120px; text-align:left; float:left;"><input type=checkbox name="graphic">Graphic</div>
            <div style="width:120px; text-align:left; float:left;"><input type=checkbox name="audio">Audio</div>
            <div style="width:120px; text-align:left; float:left;"><input type=checkbox name="other">Other</div>
          </td>
        </tr>
        </table>

	<div id="submit"> 
        <input name="submit" type="submit" value="Submit" onclick="return checkForm();" title="Submit" />
        <input name="pgPostback" type="hidden" value="yes" />
    </div>
    
    <p style="font-weight:bold;">Please fill in all fields!</p>
</form>

	<!-- end #mainContent --></div>
<!-- end #container --></div>
</body>
</html>

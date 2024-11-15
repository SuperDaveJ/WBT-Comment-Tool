<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Lesson Rating and Comments</title>
<link href="virtualPilot.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/JavaScript" src="radioButtonControl.js"></script>
<!-- #include file="adovbs.inc" -->
<%
Session.Timeout = 60

'Set variables here
userID = Request.QueryString("uID")
courseID = Request.QueryString("cID")
lessonID = CInt(Request.QueryString("lID"))

If (UserID = "") Then
	Response.write("You are not logged in or timed out.  Please log in again.")
	Response.End()
End If


strPostTarget = CStr(Request("URL")) & "?uID=" & userID & "&cID=" & courseID & "&lID=" & lessonID
pgPostback = Request.Form("pgPostback")

set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open(Session("strConnectSQL"))

' ***** Get Task IDs and Text from database *****
Dim arrQuestion
Dim nItems, courseURL, lessonTitle
set rsQuestion = Server.CreateObject("ADODB.Recordset")
set objRS = Server.CreateObject("ADODB.Recordset")

set objCmd = Server.CreateObject("ADODB.Command")
'Get course URL
With objCmd
	.ActiveConnection = objConn
	.CommandType = adCmdText
	.CommandText = "SELECT courseURL FROM tblCourseList WHERE courseID='" & courseID & "'"
	Set objRS = .Execute()
End With
If (NOT objRS.EOF) Then
	courseURL = objRS.Fields(0)
Else
	courseURL = "pageNotFound.htm"
End If
objRS.Close
'Get course or lesson title
If (lessonID = 999) Then
	objCmd.CommandText = "SELECT courseTitle FROM tblCourseList WHERE courseID='" & courseID & "'"
Else
	objCmd.CommandText = "SELECT lessonTitle FROM tblLessonList WHERE courseID='" & courseID & "' AND lessonID=" & lessonID
End If
Set objRS = objCmd.Execute()
If (NOT objRS.EOF) Then
	If (lessonID = 999) Then
		lessonTitle = objRS.Fields(0)
	Else
		lessonTitle = "Lesson " & lessonID & ": " & objRS.Fields(0)
	End If
Else
	lessonTitle = ""
End If
objRS.Close

Set objRS = Nothing

With objCmd
	.CommandType = adCmdStoredProc
	.CommandText = "usp_getQuestionText"
	.Parameters.Append .CreateParameter("@courseID", adVarChar, adParamInput, 30)
	.Parameters("@courseID").Value = courseID
	.Parameters.Append .CreateParameter("@lessonID", adInteger, adParamInput)
	.Parameters("@lessonID").Value = lessonID
End With
rsQuestion.CursorType = adOpenStatic
rsQuestion.CursorLocation = adUseServer
rsQuestion.LockType = adLockOptimistic
Set rsQuestion = objCmd.Execute()

arrQuestion = rsQuestion.Getrows
'The UBound function returns the largest subscript for the indicated dimension of an array.
'The actual number of items should be ONE more than the largest subscript.
nItems = uBound(arrQuestion,2) + 1
rsQuestion.Close
set rsQuestion = Nothing
Set objCmd = Nothing

' *** If postback
If (CStr(Request("pgPostback")) = "yes") Then
	'get values for database
	strAdditionalComment = Trim(Request.Form("addComment"))
	Dim Temp
	Dim arrRating, arrComment
	ReDim arrRating(nItems-1)
	ReDim arrComment(nItems-1)

	For Itemp=0 to nItems-1
		Temp = Request.Form("radio" & Itemp)
		if (Temp = "") then
			Temp = 0
		end if
		arrRating(Itemp) = Temp
		arrComment(Itemp) = Trim(Request.Form("comment" & Itemp))
	Next
End If

Sub manageUserData(whatToDo)
	'whatToDo is either insert or update
	set objCmd = Server.CreateObject("ADODB.Command")
	'****************** Additional Comment (one per lesson) ************************
	With objCmd
		.ActiveConnection = objConn
		.CommandText = "usp_manageAdditionalComments"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter("@whatToDo", adVarChar, adParamInput, 15)
		.Parameters("@whatToDo").Value = whatToDo
		.Parameters.Append .CreateParameter("@userID", adVarChar, adParamInput, 50)
		.Parameters("@userID").Value = userID
		.Parameters.Append .CreateParameter("@courseID", adVarChar, adParamInput, 30)
		.Parameters("@courseID").Value = courseID
		.Parameters.Append .CreateParameter("@lessonID", adInteger, adParamInput)
		.Parameters("@lessonID").Value = lessonID
		.Parameters.Append .CreateParameter("@comment", adVarChar, adParamInput, -1)
		.Parameters("@comment").Value = strAdditionalComment
		.Execute()
	End With

	'****************** Rating and Comment (one per question) ************************
	With objCmd
		.CommandText = "usp_manageLessonRating"
		.Parameters.Append .CreateParameter("@qID", adInteger, adParamInput)
		.Parameters.Append .CreateParameter("@rating", adTinyInt, adParamInput)
	End With

	For Itemp=0 to nItems-1
		'assign parameter values
		objCmd.Parameters("@qID").Value = arrQuestion(0,Itemp)
		objCmd.Parameters("@rating").Value = arrRating(Itemp)
		objCmd.Parameters("@comment").Value = arrComment(Itemp)
		objCmd.Execute()
	Next
	set objCmd = Nothing
End Sub

' ***** Check if there are any user data on this page *****
set rsUserRating = Server.CreateObject("ADODB.Recordset")
set rsAddComment = Server.CreateObject("ADODB.Recordset")
set objCmd = Server.CreateObject("ADODB.Command")

With objCmd
	.ActiveConnection = objConn
	.CommandType = adCmdStoredProc
	.CommandText = "usp_getUserRatings"
	.Parameters.Append .CreateParameter("@userID", adVarChar, adParamInput, 50)
	.Parameters("@userID").Value = userID
	.Parameters.Append .CreateParameter("@courseID", adVarChar, adParamInput, 30)
	.Parameters("@courseID").Value = courseID
	.Parameters.Append .CreateParameter("@lessonID", adInteger, adParamInput)
	.Parameters("@lessonID").Value = lessonID
	Set rsUserRating = .Execute()
	.CommandText = "usp_getAdditionalComment"
	Set rsAddComment = .Execute()
End With
Set objCmd = Nothing

' Get additional comment
If (NOT rsAddComment.EOF) Then
	strAddComment = rsAddComment.Fields(0)
Else
	strAddComment = ""
End If
rsAddComment.Close
Set rsAddComment = Nothing

' Get rating and comment for each question
Dim arrRatingFromDB, arrCommFromDB
ReDim arrRatingFromDB(nItems)
ReDim arrCommFromDB(nItems)

If (NOT rsUserRating.EOF) Then
	If (pgPostback = "yes") Then
		rsUserRating.close
		Set rsUserRating = Nothing
		manageUserData("update")
		objConn.Close()
		Set objConn = nothing
    	response.buffer = TRUE
    	response.clear
		'response.redirect(courseURL)
		Response.End
	Else
		i = 0
		Do While NOT rsUserRating.EOF
			arrRatingFromDB(i) = rsUserRating.Fields(1)
			arrCommFromDB(i) = rsUserRating.Fields(2)
			rsUserRating.MoveNext
			i = i + 1
		Loop
		rsUserRating.close
		Set rsUserRating = Nothing	
	End If
Else
	If (pgPostback = "yes") Then
		' redirect
		rsUserRating.close
		Set rsUserRating = Nothing
		manageUserData("insert")
		objConn.Close()
		Set objConn = nothing
		response.buffer = TRUE
		response.clear
		'Response.Redirect(courseURL)
		response.end
	Else
		rsUserRating.close
		Set rsUserRating = Nothing
		For i = 0 to nItems-1
			arrRatingFromDB(i) = 0
			arrCommFromDB(i) = ""
		Next
		strAddComment = ""
	End If
End If

objConn.Close
set objConn = Nothing

%>
</head>

<body class="oneColElsCtr" onunload="parent.close()">
<div id="container">
  <div id="mainContent">

<p style="font-weight: bold; text-align:center; margin-top:20px;"><%=lessonTitle%></p>

<form name="form1" method="post" action="<%=strPostTarget%>">
<table width="95%" cellpadding="3" border="1" style="border:1px solid #000000; border-collapse:collapse;">
  <tr>
    <th width="200px" scope="col">&nbsp;</th>
    <th scope="col" align="center">4<br />Highly satisfactory</th>
    <th scope="col" align="center">3<br />Satisfactory</th>
    <th scope="col" align="center">2<br />Somewhat satisfactory</th>
    <th scope="col" align="center">1<br />Not at all satisfactory</th>
    <th scope="col" align="center">Rating comment </th>
  </tr>

<!-- =============== Loop through Question Items ======================== -->
	<% For i = 0 to nItems-1 %> 
    <tr> <% if (i Mod 2 = 0) then%>
        <td width="200px" class="even">
        <% else %>
        <td width="200px" class="odd">
        <% end if %>
    
        <p class="scaletext"> 
            <input type="hidden" name="hiddenfield<%=(i)%>" value="<%=arrQuestion(0,i)%>" /><%=arrQuestion(1,i)%>
        </p>
      </td>
      
      <td width="40" align="center">
        <input type="radio" name="radio<%=(i)%>" value="4" <%If arrRatingFromDB(i)=4 Then%> checked <%End If%> >
      </td>
      <td width="40" align="center">
        <input type="radio" name="radio<%=(i)%>" value="3" <%If arrRatingFromDB(i)=3 Then%> checked <%End If%> >
      </td>
      <td width="40" align="center">
        <input type="radio" name="radio<%=(i)%>" value="2" <%If arrRatingFromDB(i)=2 Then%> checked <%End If%> >
      </td>
      <td width="40" align="center">
        <input type="radio" name="radio<%=(i)%>" value="1" <%If arrRatingFromDB(i)=1 Then%> checked <%End If%> >
      </td>
      <td width="40" align="center">
        <textarea name="comment<%=(i)%>" cols="30" rows="3"><%=arrCommFromDB(i)%></textarea>
      </td>
    
    </tr>
<% Next %> 
<!-- ============================== End of Loop ============================================== -->
	<tr>
    	<td>Additional comments:</td>
    	<td colspan="5">
        	<textarea name="addComment" cols="73" rows="5"><%=strAddComment%></textarea>
        </td>
    </tr>
</table>

	<div id="submit"> 
        <input name="submit" type="submit" value="Submit" onclick="return answeredAll(<%=nItems%>)" title="Submit" />
        <input name="pgPostback" type="hidden" value="yes" />
    </div>
</form>

	<!-- end #mainContent --></div>
<!-- end #container --></div>
</body>
</html>

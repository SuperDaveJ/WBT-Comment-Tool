<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<head>
<!-- #include file="adovbs.inc" -->
</head>
<%
Session.Timeout = 60

Dim strData, arrData
'q Data format: userID~courseID~lessonID~qID~correctAnswer~userAns1~userAns2~qStatus~userTry
strData = Request.QueryString("q")
arrData = Split(strData, "~")

strConnectSQL="Provider=SQLOLEDB;Data Source=Okcok06wbs01;Initial Catalog=PilotEvalTool;User ID=PrivateWebUserS;Password=SresUbeWetavirP;"

set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open(strConnectSQL)
set objCmd = Server.CreateObject("ADODB.Command")
With objCmd
	.ActiveConnection = objConn
	.CommandType = adCmdStoredProc
	.CommandText = "usp_ManageTestData"
	.Parameters.Append .CreateParameter("@userID", adVarChar, adParamInput, 50)
	.Parameters("@userID").Value = arrData(0)
	.Parameters.Append .CreateParameter("@courseID", adVarChar, adParamInput, 30)
	.Parameters("@courseID").Value = arrData(1)
	.Parameters.Append .CreateParameter("@lessonID", adInteger, adParamInput)
	.Parameters("@lessonID").Value = arrData(2)
	.Parameters.Append .CreateParameter("@qID", adVarChar, adParamInput, 20)
	.Parameters("@qID").Value = arrData(3)
	.Parameters.Append .CreateParameter("@ansCorrect", adVarChar, adParamInput, 30)
	.Parameters("@ansCorrect").Value = arrData(4)
	.Parameters.Append .CreateParameter("@ansUser1", adVarChar, adParamInput, 30)
	.Parameters("@ansUser1").Value = arrData(5)
	.Parameters.Append .CreateParameter("@ansUser2", adVarChar, adParamInput, 30)
	.Parameters("@ansUser2").Value = arrData(6)
	.Parameters.Append .CreateParameter("@qStatus", adTinyInt, adParamInput)
	.Parameters("@qStatus").Value = arrData(7)
	.Parameters.Append .CreateParameter("@tryUser", adTinyInt, adParamInput)
	.Parameters("@tryUser").Value = arrData(8)
	.Parameters.Append .CreateParameter("@dateAdded", adInteger, adParamInput)
	.Parameters("@dateAdded").Value = Now()
	.Execute()
End With

Set objCmd = Nothing
objConn.Close()
Set objConn = Nothing

'Response.Write(strData)
%>
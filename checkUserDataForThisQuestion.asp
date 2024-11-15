<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Session.Timeout = 60

Dim strData, arrData
'q Info format: userID~courseID~lessonID~qID
'strData = "Xihai Zhang~3~2~mc"

strData = Request.QueryString("q")
arrData = Split(strData, "~")

strConnectSQL="Provider=SQLOLEDB;Data Source=Okcok06wbs01;Initial Catalog=PilotEvalTool;User ID=PrivateWebUserS;Password=SresUbeWetavirP;"

set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open(strConnectSQL)
set objCmd = Server.CreateObject("ADODB.Command")
'adVarChar=200, adInteger=3, adParamInput = &H0001, adParamOutput = &H0002, adCmdStoredProc = &H0004
With objCmd
	.ActiveConnection = objConn
	.CommandType = &H0004
	.CommandText = "usp_GetTestData"
	.Parameters.Append .CreateParameter("@userID", 200, &H0001, 50)
	.Parameters("@userID").Value = arrData(0)
	.Parameters.Append .CreateParameter("@courseID", 200, &H0001, 30)
	.Parameters("@courseID").Value = arrData(1)
	.Parameters.Append .CreateParameter("@lessonID", 3, &H0001)
	.Parameters("@lessonID").Value = arrData(2)
	.Parameters.Append .CreateParameter("@qID", 200, &H0001, 20)
	.Parameters("@qID").Value = arrData(3)
	.parameters.Append .CreateParameter("@qCount", 3, &H0002)
	.Execute()
End With

If (objCmd.Parameters("@qCount") > 0) Then
	qDone = "yes"
Else
	qDone = "no"
End If

Set objCmd = Nothing
objConn.Close()
Set objConn = Nothing

Response.Write(qDone)
%>
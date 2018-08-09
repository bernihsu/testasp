<html>
<%
On Error Resume Next

strDate = "April 32, 2019"
d=CDate(strDate)

Response.Write "test string convert to date<br>"
Response.Write "The result is [" & d  & "]<br>"
Response.Write "IsDate [" & IsDate(strDate) & "]<br>"

' Genareal GUID for batch name
Set typeLib = Server.CreateObject("Scriptlet.TypeLib")
myGuid = typeLib.Guid
myGuid = Left(myGuid, 38)
Set typeLib = Nothing

dbAction = request.querystring("dbAction")
delKey = request.querystring("bufferName")
If dbAction = "delete" Then
	'delete all record in excelBufferHeaders
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open "provider=SQLOLEDB.1;persist security info=false;DATABASE=patDB;User ID=patentCN;password=ilovesquall0988;initial catalog=;Data Source=GAINIASERVER\SQLEXPRESS " 
	IF delKey = "" THEN
		commSQL = "DELETE  From [excelBufferHeaders]"
	Else
		commSQL = "DELETE  From [excelBufferHeaders] WHERE bufferName = '" & delKey &"'"
	End If
	conn.Execute(commSQL)
	Response.Redirect "testSQL.asp?dbAction=View"

ElseIf dbAction = "insert" Then

	'insert excelBufferHeaders
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open "provider=SQLOLEDB.1;persist security info=false;DATABASE=patDB;User ID=patentCN;password=ilovesquall0988;initial catalog=;Data Source=GAINIASERVER\SQLEXPRESS " 
	commSQL = "INSERT INTO [dbo].[excelBufferHeaders] ([bufferName],[recordType])VALUES ('" & myGuid & "',1)"
	conn.Execute(commSQL)

	Set bufferHeadercount = Server.CreateObject("ADODB.Recordset")
	bufferHeadercount.Open "Select count(*) AS count From excelBufferHeaders",conn,1,3
	nowcount=bufferHeadercount("count")
	bufferHeadercount.close

	Set bufferHeader = Server.CreateObject("ADODB.Recordset")
	bufferHeader.Open "Select * From excelBufferHeaders",conn,1,3	

	Response.Write "<table border=""1""><thead><tr>"
	Response.Write "</tr></thead><tbody>"
	IF NOT bufferHeader.EOF THEN
		WHILE NOT bufferHeader.eof
			Response.Write "<tr>"
			FOR EACH Field IN bufferHeader.Fields
				If IsNull(Field.value) Then
					showValue = "**"
				Else
					showValue = Field.value
				End If
				IF Field.Name = "bufferName" THEN
					Response.Write "<td><a href='testSQL.asp?dbAction=delete&bufferName=" & showValue & "'>" & showValue & "</a></td>"
				ELSE				
					Response.Write "<td>" & showValue & "</td>"
				End If
			NEXT
			Response.Write "</tr>"
			bufferHeader.movenext
		WEND
	END IF
	Response.Write "</tbody></table>"

Else
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open "provider=SQLOLEDB.1;persist security info=false;DATABASE=patDB;User ID=patentCN;password=ilovesquall0988;initial catalog=;Data Source=GAINIASERVER\SQLEXPRESS " 
	Set bufferHeader = Server.CreateObject("ADODB.Recordset")
	bufferHeader.Open "Select * From excelBufferHeaders",conn,1,3	

	Response.Write "<table border=""1""><thead><tr>"
	Response.Write "</tr></thead><tbody>"
	IF NOT bufferHeader.EOF THEN
		WHILE NOT bufferHeader.eof
			Response.Write "<tr>"
			FOR EACH Field IN bufferHeader.Fields
				If IsNull(Field.value) Then
					showValue = "**"
				Else
					showValue = Field.value
				End If
				IF Field.Name = "bufferName" THEN
					Response.Write "<td><a href='testSQL.asp?dbAction=delete&bufferName=" & showValue & "'>" & showValue & "</a></td>"
				ELSE				
					Response.Write "<td>" & showValue & "</td>"
				End If
			NEXT
			Response.Write "</tr>"
			bufferHeader.movenext
		WEND
	END IF
	Response.Write "</tbody></table>"	

End If

'Handle the error
If Err.Number <> 0 Then '發生Exception'
	Response.Write "Error Message: " & Err.Description  & "<BR>SQL Command :" & commSQL
	conn.close 
	'解除Exception狀態, 否則後面的語法如果還有發生Exception都會被忽略'
	On Error GoTo 0
Else '無error'
	Response.Write "SQL success!<BR>SQL Command :" & commSQL & "<br>目前資料筆數：" & nowcount
	conn.close
	'解除Exception狀態, 否則後面的語法如果還有發生Exception都會被忽略'
	On Error GoTo 0
End If

%>

</html>
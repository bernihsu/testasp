<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<%
On Error Resume Next

'init connection string
Set conn = Server.CreateObject("ADODB.Connection")
conn.open "provider=SQLOLEDB.1;persist security info=false;DATABASE=patDB;User ID=patentCN;password=ilovesquall0988;initial catalog=;Data Source=GAINIASERVER\SQLEXPRESS " 

Set oRS=Server.CreateObject("ADODB.recordset")
commSQL = "select * from dbo.excelBufferHeaders"
oRS.Open commSQL,conn

Response.Write "<div>Excel Buffer View<p>"
Response.Write "</div><hr>"
Response.Write "<table border=""1""><thead><tr>"
'show record name
For i = 0 To 8
	Response.Write "<th>" & oRS.Fields.Item(i).name & "</th>"
Next
Do until oRS.Eof
	'0~8	
	'show record value
	Response.Write "<tr>"	
	For i = 0 to 8
		prcessValue = "null"
		If (Not IsNull(oRS.Fields.Item(i).value)) Then prcessValue = oRS.Fields.Item(i).value
		If i = 0 then 
			link = "<a href=""www.google.com"">Go To</a>"
			If oRS.Fields.Item(1).value = 1 then  'patentData
				link = "<a href=""confirmPatentData.asp?bufferName=" & prcessValue & """>Go To</a>"
			Else 'judgment
				link = "<a href=""confirmJudgmentData.asp?bufferName=" & prcessValue & """>Go To</a>"			
			End If		
			Response.Write "<td>" & link & "</td>"		
		Else
			Response.Write "<td>" & prcessValue & "</td>"		
		End If
	Next
	Response.Write "</tr>"
	oRS.movenext
Loop	
Response.Write "</tbody></table>"

oRs.close
conn.close			


'Handle the error
If Err.Number <> 0 Then '發生Exception'
	Response.Write "Error Message: " & Err.Description  & "<BR>SQL Command :" & commSQL
	conn.close 
	'解除Exception狀態, 否則後面的語法如果還有發生Exception都會被忽略'
	On Error GoTo 0
Else '無error'
	conn.close	
	'解除Exception狀態, 否則後面的語法如果還有發生Exception都會被忽略'
	On Error GoTo 0
End If
%>

<hr />

<div>
excel Buffer:<%=nowcount%> records
</div>
<br/>

<html>
<script language=javascript>

</script>

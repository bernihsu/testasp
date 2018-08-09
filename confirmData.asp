<!--#INCLUDE FILE="upload2/clsUpload.asp"-->
<html>

<%
On Error Resume Next
Dim objUpload
Dim strFileName
Dim strPath

bufferName = request.querystring("bufferName")
recordType = request.querystring("type")
'read all excel buffer record via bufferName
Set conn = Server.CreateObject("ADODB.Connection")
set oRS=Server.CreateObject("ADODB.recordset")
conn.open "provider=SQLOLEDB.1;persist security info=false;DATABASE=patDB;User ID=patentCN;password=ilovesquall0988;initial catalog=;Data Source=GAINIASERVER\SQLEXPRESS " 

Set buffCount = Server.CreateObject("ADODB.Recordset")
buffCount.Open "Select count(*) AS count From [patDB].[dbo].[excelBuffers] WHERE [bufferName] = '" & bufferName & "'"
nowcount=buffCount("count")
buffCount.close

commSQL = "SELECT  *  FROM [patDB].[dbo].[excelBuffers] WHERE [bufferName] = '" & bufferName & "'"

oRS.Open commSQL,conn
'if(isnull(oRs.Fields.Item(1).value)) then
'1~26
'27~45
Response.Write "<table border=""1""><thead><tr>"
'show record name
IF recordType = "1" THEN 
	FOR i = 0 to 25
		Response.Write "<th>" & oRS.Fields.Item(i).name & "</th>"
	NEXT	
Else
	FOR i = 0 to 3
		Response.Write "<th>" & oRS.Fields.Item(i).name & "</th>"
	NEXT
	FOR i = 26 to 44
		Response.Write "<th>" & oRS.Fields.Item(i).name & "</th>"
	NEXT
End If
'show record value
Response.Write "</tr></thead><tbody>"
do until oRS.eof
Response.Write "<tr>"
	IF recordType = "1" THEN 
		FOR i = 0 to 25
			Response.Write "<td>" & oRS.Fields.Item(i).value & "</td>"
		NEXT	
	Else
		FOR i = 0 to 3
			Response.Write "<td>" & oRS.Fields.Item(i).value & "</td>"
		NEXT	
		FOR i = 26 to 44
			Response.Write "<td>" & oRS.Fields.Item(i).value & "</td>"
		NEXT		
	End If
	Response.Write "</tr>"
oRS.movenext
loop
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
	Response.Write "SQL success!<BR>SQL Command :" & commSQL & "<br>目前資料筆數：" & recordCount
	conn.close
	'解除Exception狀態, 否則後面的語法如果還有發生Exception都會被忽略'
	On Error GoTo 0
End If
%>

<hr />

<div>此檔案共有<%=lineNo%>筆資料</div>
<br />


<script language=javascript>
    //alert('上傳成功');
	/*
   setTimeout(function () {

       location.href = "uploadXls.asp?state=success"
   }, 8000); */
</script>

<!--#INCLUDE FILE="upload2/clsUpload.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body>
<%
On Error Resume Next

bufferName = request.querystring("bufferName")
UserAction = request.querystring("Action")
'init connection string
Set conn = Server.CreateObject("ADODB.Connection")
conn.open "provider=SQLOLEDB.1;persist security info=false;DATABASE=patDB;User ID=patentCN;password=ilovesquall0988;initial catalog=;Data Source=GAINIASERVER\SQLEXPRESS " 

IF (bufferName <> "") AND (UserAction = "allInsert") THEN

	set rs=server.createobject("adodb.recordset")
	commSQL = "exec excelBufferTojudgmentDatas @bufferName = '" & bufferName & "'"
	conn.Execute (commSQL)	

ElseIF (bufferName <> "") AND (UserAction = "allDelete") THEN
	
	commSQL = "delete from dbo.excelBuffers where [bufferName] = '" & bufferName & "'" & _
		      "delete from dbo.excelBufferHeaders where [bufferName] = '" & bufferName & "'"
	conn.Execute commSQL
		
	
ElseIF (bufferName <> "") AND (UserAction = "") THEN
	'
	'read all excel buffer record via bufferName
	Set buffCount = Server.CreateObject("ADODB.Recordset")
	buffCount.Open "Select count(*) AS count From [patDB].[dbo].[excelBuffers] WHERE [recordType] = 2 and [bufferName] = '" & bufferName & "'" ,conn,1,3
	nowcount=buffCount("count")
	buffCount.close

	Set oRS=Server.CreateObject("ADODB.recordset")
	commSQL = "SELECT  *  FROM [patDB].[dbo].[excelBuffers] WHERE [recordType] = 2 and [bufferName] = '" & bufferName & "'"
	oRS.Open commSQL,conn
	'if(isnull(oRs.Fields.Item(1).value)) then
	'1~26
	'27~45
	Response.Write "<div>此次判決資料上傳Excel暫存檔名為:[" & bufferName & "]<p>"
	Response.Write "※藍色資料行且第一欄寫「準備新增」將會嚐試新增至資料庫<br>"	
	Response.Write "※黃色資料行不會新增至資料庫，請先改正其欄位內容錯誤。<br>"			
	Response.Write "<button onClick=""RecordAction('bufferName=" & bufferName & "&Action=allInsert','insert')"">確定新增</button>"
	Response.Write "<button onClick=""RecordAction('bufferName=" & bufferName & "&Action=allDelete','delete')"">全部刪除</button>"
	Response.Write "</div><hr>"
	Response.Write "<table border=""1""><thead><tr>"
	'show record name
	Response.Write "<th>" & oRS.Fields.Item(3).name & "</th>"	
	FOR i = 29 to 44
		Response.Write "<th>" & oRS.Fields.Item(i).name & "</th>"
	NEXT	
	FOR i = 27 to 28
		Response.Write "<th>" & oRS.Fields.Item(i).name & "</th>"
	NEXT			

	'show record value
	Response.Write "</tr></thead><tbody>"
	do until oRS.eof

	if oRS.Fields.Item(3).value = 0 then 
		Response.Write "<tr bgcolor=""#ffff33"">" '如果此行有錯誤則為紅色
	else 
		Response.Write "<tr bgcolor=""#ccffff"">" 
	end if	
	    '[missFKey],judgmentDate,
		if oRS.Fields.Item(3).value = 0 then 'error flag and check content of error
			ErrContent = oRS.Fields.Item("errorFields").value
			showErrMsg = ""
			if InStr(ErrContent,"exist") > 0 then
				showErrMsg = "主鍵重複"			
			end if
			if (InStr(ErrContent,"missFKey") > 0)then	
				showErrMsg = "<br>查無專利號"
			end if
			if showErrMsg = "" then showErrMsg = "內容錯誤"
			showErrMsg = "<td>" & showErrMsg & "</td>"
			Response.Write showErrMsg
		else  
			Response.Write "<td>準備新增</td>"
		end if
		
		FOR i = 29 to 44
			prcessValue = "null"
			IF (Not IsNull(oRS.Fields.Item(i).value)) Then prcessValue = oRS.Fields.Item(i).value
			
			if InStr(oRS.Fields.Item("errorFields").value,oRS.Fields.Item(i).Name) > 0 then	
				Response.Write "<td><B><S>" & prcessValue & "</S></B></td>"			
			else
				Response.Write "<td>" & prcessValue & "</td>"
			end if
		NEXT	
		FOR i = 27 to 28
			prcessValue = "null"
			IF (Not IsNull(oRS.Fields.Item(i).value)) Then prcessValue = oRS.Fields.Item(i).value
			if InStr(oRS.Fields.Item("errorFields").value,oRS.Fields.Item(i).Name) > 0 then	
				Response.Write "<td><B><S>" & prcessValue & "</S></B></td>"			
			else
				Response.Write "<td>" & prcessValue & "</td>"
			end if
		NEXT			
		Response.Write "</tr>"
	oRS.movenext
	loop
	Response.Write "</tbody></table>"

	oRs.close
	conn.close		
END IF

'Handle the error
If Err.Number <> 0 Then '發生Exception'
	Response.Write "Error Message: " & Err.Description  & "<BR>SQL Command :" & commSQL
	conn.close 
	'解除Exception狀態, 否則後面的語法如果還有發生Exception都會被忽略'
	On Error GoTo 0
Else '無error'
	conn.close	
	IF (bufferName <> "") AND (UserAction = "allInsert") THEN
		Response.Write "SQL success!<BR>SQL Command :" & commSQL 
		Response.Redirect "uploadXls.asp?state=success&UserAction=allInsert&bufferName=" & bufferName	
	ElseIF (bufferName <> "") AND (UserAction = "allDelete") THEN
		Response.Write "SQL success!<BR>SQL Command :" & commSQL & "<br>"
		Response.Redirect "uploadXls.asp?state=success&UserAction=allDelete&bufferName=" & bufferName
	ELSE
		Response.Write "SQL success!<BR>SQL Command :" & commSQL & "<br>目前資料筆數：" & recordCount	
	END IF	

	'解除Exception狀態, 否則後面的語法如果還有發生Exception都會被忽略'
	On Error GoTo 0
End If
%>

<hr />

<div>
此檔案共有<%=nowcount%>筆資料
</div>
<br/>
</body>
<html>
<script language=javascript>
function RecordAction(theURL,doAction){
	//alert("doing... \nconfirmPatentData.asp?" + doAction);
	var showMsg = '判斷不出來';
	if (doAction == 'delete'){showMsg = "確定刪除\n此次上傳Excel資料";}
	else if (doAction == 'insert'){showMsg = "確定新增\n此次上傳Excel資料";}
	
	var r=confirm(showMsg)
	if (r == true) {
		location.href='confirmJudgmentData.asp?' + theURL;
	}

}
</script>

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
	getInsertCount = 10
	
	set rs=server.createobject("adodb.recordset")
	commSQL = "exec excelBufferTopatentDatas @bufferName = '" & bufferName & "'"
	conn.Execute (commSQL)	
	
ElseIF (bufferName <> "") AND (UserAction = "allDelete") THEN
	
	commSQL = "delete from dbo.excelBuffers where [bufferName] = '" & bufferName & "'" & _
		      "delete from dbo.excelBufferHeaders where [bufferName] = '" & bufferName & "'"
	conn.Execute commSQL
		
	
ElseIF (bufferName <> "") AND (UserAction = "") THEN
	'
	'read all excel buffer record via bufferName
	Set buffCount = Server.CreateObject("ADODB.Recordset")
	buffCount.Open "Select count(*) AS count From [patDB].[dbo].[excelBuffers] WHERE [recordType] = 1 and [recordOK] = 1 and [bufferName] = '" & bufferName & "'" ,conn,1,3
	nowcount=buffCount("count")
	buffCount.close

	Set oRS=Server.CreateObject("ADODB.recordset")
	commSQL = "SELECT  *  FROM [patDB].[dbo].[excelBuffers] WHERE [recordType] = 1 and [bufferName] = '" & bufferName & "'"
	oRS.Open commSQL,conn
	'if(isnull(oRs.Fields.Item(1).value)) then
	'1~26
	'27~45
	Response.Write "<div>�����M�Q��ƤW��Excel�Ȧs�ɦW��:[" & bufferName & "]<p>"
	Response.Write "���Ŧ��Ʀ�B�Ĥ@��g�u�ǳƷs�W�v�N�|�|�շs�W�ܸ�Ʈw<br>"	
	Response.Write "�������Ʀ椣�|�s�W�ܸ�Ʈw�A�Х��勵����줺�e���~�C<br>"		
	if nowcount > 0 then
		Response.Write "<button onClick=""RecordAction('bufferName=" & bufferName & "&Action=allInsert','insert')"">�T�w�s�W</button>"
	end if
	Response.Write "<button onClick=""RecordAction('bufferName=" & bufferName & "&Action=allDelete','delete')"">�����R��</button>"
	Response.Write "</div><hr>"
	Response.Write "<table border=""1""><thead><tr>"
	'show record name

	FOR i = 3 to 25
		Response.Write "<th>" & oRS.Fields.Item(i).name & "</th>"
	NEXT	

	'show record value
	Response.Write "</tr></thead><tbody>"
	do until oRS.eof
	
	if oRS.Fields.Item(3).value = 0 then 
		Response.Write "<tr bgcolor=""#ffff33"">" '�p�G���榳���~�h������
	else 
		Response.Write "<tr bgcolor=""#ccffff"">" 
	end if

		FOR i = 3 to 25
			prcessValue = "null"
			IF (Not IsNull(oRS.Fields.Item(i).value)) Then prcessValue = oRS.Fields.Item(i).value
			IF (i = 3) then 'make a delete link
				if oRS.Fields.Item(3).value = 0 then
					'InStr
					if InStr(oRS.Fields.Item("errorFields").value,"exist") > 0 then	
						Response.Write "<td>�D�䭫��</td>"
					else 
						Response.Write "<td>���e����</td>" 
					end if
				else 
					Response.Write "<td>�ǳƷs�W</td>"
				end if
			Else'�P�_��즳���D					
				if InStr(oRS.Fields.Item("errorFields").value,oRS.Fields.Item(i).Name) > 0 then	
					Response.Write "<td><B><S>" & prcessValue & "</S></B></td>"
				else
					Response.Write "<td>" & prcessValue & "</td>"
				end if
			End if 
		NEXT	
		Response.Write "</tr>"
	oRS.movenext
	loop
	Response.Write "</tbody></table>"

	oRs.close
	conn.close		
END IF

'Handle the error
If Err.Number <> 0 Then '�o��Exception'
	Response.Write "Error Message: " & Err.Description  & "<BR>SQL Command :" & commSQL
	conn.close 
	'�Ѱ�Exception���A, �_�h�᭱���y�k�p�G�٦��o��Exception���|�Q����'
	On Error GoTo 0
Else '�Lerror'
	conn.close
	
	IF (bufferName <> "") AND (UserAction = "allInsert") THEN
		Response.Write "SQL success!<BR>SQL Command :" & commSQL 
		Response.Redirect "uploadXls.asp?state=success&UserAction=allInsert&bufferName=" & bufferName
	ElseIF (bufferName <> "") AND (UserAction = "allDelete") THEN
		Response.Write "SQL success!<BR>SQL Command :" & commSQL & "<br>"
		esponse.Redirect "uploadXls.asp?state=success&UserAction=allDelete&bufferName=" & bufferName
	ELSE
		Response.Write "SQL success!<BR>SQL Command :" & commSQL & "<br>�ثe��Ƶ��ơG" & recordCount	
	END IF
	
	'�Ѱ�Exception���A, �_�h�᭱���y�k�p�G�٦��o��Exception���|�Q����'
	On Error GoTo 0
End If
%>

<hr />

<div>
���ɮצ@��<%=nowcount%>�����
</div>
<br/>
</body>
<html>
<script language=javascript>
function RecordAction(theURL,doAction){
	//alert("doing... \nconfirmPatentData.asp?" + doAction);
	var showMsg = '�P�_���X��';
	if (doAction == 'delete'){showMsg = "�T�w�R��\n�����W��Excel���";}
	else if (doAction == 'insert'){showMsg = "�T�w�s�W\n�����W��Excel���";}
	
	var r=confirm(showMsg)
	if (r == true) {
		location.href='confirmPatentData.asp?' + theURL;
	}

}
</script>

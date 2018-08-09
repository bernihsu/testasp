<!--#INCLUDE FILE="upload2/clsUpload.asp"-->
<html>

<%
On Error Resume Next
Dim objUpload
Dim strFileName
Dim strPath
Dim dbType 

' Instantiate Upload Class
Set objUpload = New clsUpload

' 1: patentData
' 2: judgmentData
dbType = objUpload.Fields("apply").Value 

if(dbType="1") then
    strFileName = "patentData"
elseif (dbType="2") then
    strFileName = "judgmentData"
else 
    Response.Redirect "uploadXls.asp?status=error"
end if
 
strPath = Server.MapPath("upload") & "\tmp_" & strFileName & ".xls"

' Save the binary data to the file system
objUpload("File1").SaveAs strPath

' Release upload object from memory
Set objUpload = Nothing

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & strPath

' genareal GUID for batch name and insert excelBufferHeadaers
Set typeLib = Server.CreateObject("Scriptlet.TypeLib")
myGuid = typeLib.Guid
myGuid = Left(myGuid, 38)
Set typeLib = Nothing

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "provider=SQLOLEDB.1;persist security info=false;DATABASE=patDB;User ID=patentCN;password=ilovesquall0988;initial catalog=;Data Source=GAINIASERVER\SQLEXPRESS " 
commSQL = "INSERT INTO [dbo].[excelBufferHeaders] ([bufferName],[recordType],[createUserId],[updateUserId])VALUES ('" & myGuid & "'," & dbType & ",'berni','berni')"
conn.Execute(commSQL)

lineNo = 1 'Counting total record
errorCount = 0 'Counting total error record
okCount = 0 'Counting total right record
if(dbType="1")then
    ' patentData 專利資料
	' read excel content and insert into excelBuffer
	Set oRS = oConn.Execute("select * from [工作表1$]")
    do until oRS.EOF
		'init SQLcomm
		commSQL = "insert into [dbo].[excelBuffers] (bufferName,bufferLine,recordType,recordOK,patentNo,applyDate,publishNo,publishDate,inventionName,IPCCategory,applicant,inventor,priorityrightNo,priorityrightDate,summary,patentClaim,keyword,designPatentLOC,designPatentDescribe,attorney,agent,nationOfApplicant,FTCategory,UCCategory,ECLACategory,createUserId,updateUserId,errorFields) VALUES ("

		'Start valid columnes
		errorField = ""
		recodOk = 1
		if IsNull(oRs.Fields.Item(0).value) then 'patentNo
			recodOk = 0 
			errorField = errorField + "[0],"
		else 
			'check there is not in patentData		errorField = ""
			Set patentCount = Server.CreateObject("ADODB.Recordset")
			patentCount.OPEN "select count(patentNo) as myCount from patentDatas where patentNo = '" & oRs.Fields.Item(0).value & "'" ,conn,1,3		
			existCount = 0
			existCount = patentCount("myCount")
			patentCount.close
			set patentCount = nothing					
			If existCount > 0 then
				recodOk = 0 
				errorField = errorField + "[exist],"
			end if				
		end if
		if (Not IsDate(oRs.Fields.Item(1).value)) then 'applyDate		
			recodOk = 0
			errorField = errorField & oRs.Fields.Item(1).name &","
		end if
		if (Not IsDate(oRs.Fields.Item(3).value)) then 'publishDate					
			recodOk = 0
			errorField = errorField & oRs.Fields.Item(3).name &","
		end if
		if (Not IsDate(oRs.Fields.Item(9).value)) then 'priorityrightDate
			recodOk = 0
			errorField = errorField & oRs.Fields.Item(9).name & ","
		end if
		if recodOk = 0 then 
			errorCount = errorCount + 1
		else 
			okCount = okCount + 1
		end if
		'End valid columnes
		
		commSQL = commSQL & "'" & myGuid & "'," & lineNo & "," & dbType & "," & recodOk
		
		for i=0 to 20
		
			if(isnull(oRs.Fields.Item(i).value)) then
				commSQL = commSQL & ",null "
			else
				commSQL = commSQL & ",'" & replace(oRs.Fields.Item(i).value, "'","''") & "'"				
			end if

		next 	
		commSQL = commSQL & ",'berni','berni','" & errorField & "')"		
		conn.Execute(commSQL)		
		lineNo = lineNo + 1
		
	oRS.movenext
	loop
	
	commSQL = "update [dbo].[excelBufferHeaders] set recordCount = " & lineNo & ",recordErrorCount = " & errorCount & ",recordOkCount = " & okCount & _
	          ",updateAt = GETDATE() WHERE [bufferName] = '" & myGuid & "'"
	conn.Execute(commSQL)		

oRs.close
oConn.close
conn.close		
	
	
elseif(dbType="2") then
    ' judgmentData 判決資料
	' read excel content and insert into excelBuffer
	Set oRS = oConn.Execute("select * from [工作表1$]")
    do until oRS.EOF
		'init SQLcomm
		commSQL = "insert into [dbo].[excelBuffers](bufferName,bufferLine,recordType,recordOK,caseNo,patentno2,court,trialLevel,trialDate,judgmentDate,plaintiff1,plaintiffAgent1,plaintiff2,plaintiffAgent2,defendant1,defendantAgent1,defendant2,defendantAgent2,[2ndAppeal],[1stRequiredAmount],[1stCompensation],judgmentResult,presidingJudge,createUserId,updateUserId,errorFields) VALUES ("


		'Stard valid columnes
		errorField = ""		
		recodOk = 1		
		if isnull(oRs.Fields.Item(0).value) then 'caseNo
			recodOk = 0 
			errorField = errorField & oRs.Fields.Item(0).name &","	
		else
			'check there is not in caseNo
			Set judgmentCount = Server.CreateObject("ADODB.Recordset")
			judgmentCount.OPEN "select count(judgmentID) as myCount from dbo.judgmentDatas where caseNo = '" & oRs.Fields.Item(0).value & "'" ,conn,1,3		
			existCount = 0
			existCount = judgmentCount("myCount")
			judgmentCount.close
			set judgmentCount = nothing					
			If existCount > 0 then
				recodOk = 0 
				errorField = errorField + "[exist],"				
			end if	
		end if
		if(isnull(oRs.Fields.Item(1).value)) then 'patentno	
			recodOk = 0 
			errorField = errorField & oRs.Fields.Item(1).name &","	
		else 
			'check there is not in patentData		errorField = ""
			Set patentCount = Server.CreateObject("ADODB.Recordset")
			patentComm = "SELECT COUNT(*) as myCount from dbo.patentDatas where patentNo = '" & oRs.Fields.Item(1).value & "'"
			patentCount.OPEN patentComm ,conn,1,3		
			existCount = 0
			existCount = patentCount("myCount")
			patentCount.close
			set patentCount = nothing					
			If existCount = 0 then
				recodOk = 0 
				errorField = errorField + "[missFKey],"				
			end if
		end if
		if (Not IsDate(oRs.Fields.Item(4).value)) then 'trialDate								
			recodOk = 0'
			errorField = errorField & oRs.Fields.Item(4).name &","									
		end if		
		if (Not IsDate(oRs.Fields.Item(5).value)) then 'judgmentDate							
			recodOk = 0
			errorField = errorField & oRs.Fields.Item(5).name &","												
		end if
		if (Not IsNumeric(oRs.Fields.Item(15).value)) then '1stRequiredAmount
			recodOk = 0
			errorField = errorField & oRs.Fields.Item(15).name &","															
		end if
	    if (Not IsNumeric(oRs.Fields.Item(16).value)) then '1stCompensation
			recodOk = 0
			errorField = errorField & oRs.Fields.Item(16).name &","															
		end if
	

		if recodOk = 0 then 
			errorCount = errorCount + 1
		else 
			okCount = okCount + 1
		end if
		'End valid columnes		
		commSQL = commSQL & "'" & myGuid & "'," & lineNo & "," & dbType & "," & recodOk
		
		for i=0 to 18
		
			if(isnull(oRs.Fields.Item(i).value)) then
				commSQL = commSQL & ",null "
			else
				commSQL = commSQL & ",'" & replace(oRs.Fields.Item(i).value, "'","''") & "'"				
			end if

		next 	
		commSQL = commSQL & ",'berni','berni','" & errorField & "')"			
		conn.Execute(commSQL)		
		lineNo = lineNo + 1
		
	oRS.movenext
	loop
	
	commSQL = "update [dbo].[excelBufferHeaders] set recordCount = " & lineNo & ",recordErrorCount = " & errorCount & ",recordOkCount = " & okCount & _
	          ",updateAt = GETDATE() WHERE [bufferName] = '" & myGuid & "'"
	conn.Execute(commSQL)		

oRs.close
oConn.close
conn.close		
end if

'Handle the error
If Err.Number <> 0 Then '發生Exception'
	Response.Write "Error Message: " & Err.Description  
	Response.Write "Error Number: " & Err.Number & "<BR>SQL Command :" & commSQL
	conn.close 
	'解除Exception狀態, 否則後面的語法如果還有發生Exception都會被忽略'
	On Error GoTo 0
Else '無error'
	Response.Write "SQL success!<BR>SQL Command :" & commSQL & "<br>目前資料筆數：" & recordCount
	
	if(dbType="1")then
		Response.Redirect "confirmPatentData.asp?bufferName=" & myGuid
	elseif (dbType="2") then
		Response.Redirect "confirmJudgmentData.asp?bufferName=" & myGuid
	else 
		Response.Redirect "uploadXls.asp?status=error"
	end if
	'解除Exception狀態, 否則後面的語法如果還有發生Exception都會被忽略'
	
	On Error GoTo 0
End If
%>

<hr />

<div>處理資料種類為<%=dbType%><br>匯入完成，共匯入<%=lineNo%>筆</div>
<br />
<div><a href="confirmPatentData.asp?bufferName=<%=myGuid%>">轉至匯入資料頁</a></div>  
<div><a href="confirmJudgmentData.asp?bufferName=<%=myGuid%>">轉至匯入資料頁</a></div>  
<div><a href="uploadXls.asp?state=success">回到匯入頁面</a>，切勿重新整理頁面 (8秒後自動返回)</div>  

<script language=javascript>
    //alert('上傳成功');
	
   setTimeout(function () {

       location.href = "uploadXls.asp?state=success"
   }
</script>

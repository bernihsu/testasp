<!--#INCLUDE FILE="upload2/clsUpload.asp"-->
<html>


<%
Dim objUpload
Dim strFileName
Dim strPath
Dim dbType
Dim total

' Instantiate Upload Class
Set objUpload = New clsUpload

' Grab the file name
'response.write(objUpload.Fields("apply").FileExt)
'response.write(objUpload.Fields("apply").Value)
'1:�M�Q���patentData
'2:�P�M���judgmentData
dbType = objUpload.Fields("apply").Value 

if(dbType="1") then
    strFileName = "patentData"
elseif (dbType="2") then
    strFileName = "judgmentData"
else 
    Response.Redirect "uploadExcel.asp?status=error"
end if
 
strPath = Server.MapPath("upload") & "\tmp_" & strFileName & ".xls"

' Save the binary data to the file system
objUpload("File1").SaveAs strPath

' Release upload object from memory
Set objUpload = Nothing



Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.Recordset")
'oConn.Open "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};" & "DBQ=" & strPath
oConn.Open "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & strPath


if(dbType="1")then

Set oRS = oConn.Execute("select * from [�u�@��1$]")

	Response.Write "<table border=""1""><thead><tr>"
	FOR EACH Column IN oRS.Fields
		Response.Write "<th>" & Column.Name & "</th>"
	NEXT
	Response.Write "</tr></thead><tbody>"
	IF NOT oRS.EOF THEN
		WHILE NOT oRS.eof
			Response.Write "<tr>"
			FOR EACH Field IN oRS.Fields
				Response.Write "<td>" & Field.value & "</td>"

			NEXT
			Response.Write "</tr>"
			total = total + 1			
			oRS.movenext
		WEND
	END IF
	Response.Write "</tbody></table>"
oRs.close
oConn.close

elseif(dbType="2") then

end if


%>

<hr />

<div>�פJ�����A�@�פJ<%=total%>��</div>
<br />
<div><a href="uploadExcel.asp?state=success">�^��פJ����</a>�A���ŭ��s��z���� (8���۰ʪ�^)</div>  

<script language=javascript>
    //alert('�W�Ǧ��\');
   setTimeout(function () {

       location.href = "uploadXls.asp?state=success"
   }, 8000);
</script>

<html>
<head>
<title>Excel export bulk data to SQL Server.</title>
</head>
<body>
<%
Dim exceltype 
exceltype = request.querystring("type")

Dim strPath,myFieldValue
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.Recordset")

if NOT IsNull(exceltype) and exceltype = "xlsx" then
	strPath = Server.MapPath("Upload") & "\excel.xlsx"
	oConn.Open "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};" & "DBQ=" & strPath
else
	strPath = Server.MapPath("Upload") & "\excel_f2.xls"
	oConn.Open "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & strPath
end if 


Set oRS = oConn.Execute("select * from [sheet2$]")

	Response.Write "<table border=""1""><thead><tr>"
	FOR EACH Column IN oRS.Fields
		Response.Write "<th>" & Column.Name & "</th>"
	NEXT
	Response.Write "</tr></thead><tbody>"
	IF NOT oRS.EOF THEN
		WHILE NOT oRS.eof
			Response.Write "<tr>"
			FOR EACH Field IN oRS.Fields
				'myFieldValue = "&nbsp;"
				myFieldValue = "--"
				IF Not IsNull(Field.value) AND Field.Value <> "" THEN myFieldValue = Field.value
				Response.Write "<td>" & myFieldValue & "</td>"
			NEXT
			Response.Write "</tr>"
			oRS.movenext
		WEND
	END IF
	Response.Write "</tbody></table>"
oRs.close
oConn.close
%>
</body>
</html>

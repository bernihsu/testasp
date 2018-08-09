<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>匯入專利資料</title>
</head>

<body>
<%
'nothing
state = request.querystring("state")
bufferName = request.querystring("bufferName")
UserAction = request.querystring("UserAction")
showMsg = ""
If UserAction = "allInsert" Then
	showMsg = "新增 excelBuffer:[" & bufferName & "]"
Else	
	showMsg = "刪除 excelBuffer:[" & bufferName & "]"
End If
	
If state = "success" Then 
	showMsg = showMsg + " 成功"
Else
	showMsg = showMsg + " 失敗"
End If	
If state <> "" Then
	Response.Write showMsg & "<hr>"
End If
	Response.Write "<a href=bufferView.asp>瀏覽excelBuffers</a><hr>"

%>
<div align="center">
 
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" >
    <tr>
      <td width="100%" align="left" valign="top">
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" height="100%">
        <tr>
          <td width="100" align="center" valign="top">
          </td>
          <td align="center" valign="top">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" >
            <tr>
              <td width="100%" align="center">
                <H1>匯入Excel資料</H1>

                    <FORM method="post" encType="multipart/form-data" name="form1" action="fileToDatabase.asp">    
                        <p>上傳類型：<select id="apply" name="apply"><option value="0">請選擇</option><option value="1">專利資料</option><option value="2">判決資料</option></select></p>
	                    <p><INPUT type="File" name="file1" id="file1"></p>
	                    <INPUT type="Submit" value="匯入">
                    </FORM>

              </td>
            </tr>
          </table>
          <div align="center">
            <center>
            <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="95%" >
              <tr>
                <td width="100%" align="center">
                 </td>
              </tr>
            </table>
            </center>
          </div>
          </td>
        </tr>
      </table>
      </td>
    </tr>
  </table>

</div>

</body>

</html>





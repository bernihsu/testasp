<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�פJ�M�Q���</title>
</head>

<body>
<%
'nothing
state = request.querystring("state")
bufferName = request.querystring("bufferName")
UserAction = request.querystring("UserAction")
showMsg = ""
If UserAction = "allInsert" Then
	showMsg = "�s�W excelBuffer:[" & bufferName & "]"
Else	
	showMsg = "�R�� excelBuffer:[" & bufferName & "]"
End If
	
If state = "success" Then 
	showMsg = showMsg + " ���\"
Else
	showMsg = showMsg + " ����"
End If	
If state <> "" Then
	Response.Write showMsg & "<hr>"
End If
	Response.Write "<a href=bufferView.asp>�s��excelBuffers</a><hr>"

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
                <H1>�פJExcel���</H1>

                    <FORM method="post" encType="multipart/form-data" name="form1" action="fileToDatabase.asp">    
                        <p>�W�������G<select id="apply" name="apply"><option value="0">�п��</option><option value="1">�M�Q���</option><option value="2">�P�M���</option></select></p>
	                    <p><INPUT type="File" name="file1" id="file1"></p>
	                    <INPUT type="Submit" value="�פJ">
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





<!--#include file="function.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>匯入專利資料</title>
</head>

<body>

<div align="center">
 
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" >
    <tr>
      <td width="100%" align="left" valign="top">
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" height="100%">
        <tr>
          <td width="100" align="center" valign="top">
          <!--#include file="lv.asp"-->
          </td>
          <td align="center" valign="top">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" >
            <tr>
              <td width="100%" align="center">
                <H1>匯入專利資料</H1>

                    <FORM method="post" encType="multipart/form-data" name="form1" action="ToFileSystem.asp">    
                        <p>上傳類型：<select id="apply" name="apply"><option value="0">請選擇</option><option value="1">判決</option><option value="2">交易</option><option value="3">冠亞鑑價</option></select></p>
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





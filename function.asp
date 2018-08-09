<%
  If Session("login") <> "y" then 
  Response.Redirect "default.asp" 
  Response.End
  End If
  Response.Buffer=true
  Response.Expires = 0
  Response.AddHeader "Pragma", "No-Cache" 
  Response.Write "<meta http-equiv=pragma content=no-cache>"
  Response.Write " <script LANGUAGE=Javascript SRC=FieldTools.js></script>"
  Response.Write " <script LANGUAGE=Javascript>"
  Response.Write Chr(13) & "function lockNum(strObj,sFlag) {"
  Response.Write Chr(13) & "        if (typeof(sFlag) == 'undefined') {"
  Response.Write Chr(13) & "            var sFlag = ''"
  Response.Write Chr(13) & "        }"
  Response.Write Chr(13) & "        if (event.keyCode < ""48"" || event.keyCode > ""57"") {"
  Response.Write Chr(13) & "            if (sFlag == '' || event.keyCode != ""46"" || strObj.value.indexOf(""."") != -1) {"
  Response.Write Chr(13) & "                event.keyCode=""0"";"
  Response.Write Chr(13) & "            }"
  Response.Write Chr(13) & "        } else {"
  Response.Write Chr(13) & "            if (strObj.value.length == 0 && event.keyCode==""47"" ) {"
  Response.Write Chr(13) & "                event.keyCode=""0"";"
  Response.Write Chr(13) & "            }"
  Response.Write Chr(13) & "        }"
  Response.Write Chr(13) & "    }"
  Response.Write "</script>"  
  Set conn = Server.CreateObject("ADODB.Connection")
'  DBPath = Server.MapPath("./db/db.mdb")
'  conn.open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
'  'Set conn = Server.CreateObject("ADODB.Connection")
'  'conn.Open application("conn")
'conn.open "Provider=SQLOLEDB.1;Server=GAINIASERVER;UID=WorkingHour;PWD=aluba"
conn.open "provider=SQLOLEDB.1;persist security info=false;DATABASE=patDB;User ID=patentCN;password=ilovesquall0988;initial catalog=;Data Source=GAINIASERVER\SQLEXPRESS " 
  Response.Write "<link href='style.css' rel='stylesheet' type='text/css'>"
%>
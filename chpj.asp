<!--#include file="function.asp"-->
<!--#include file="include/bs_function.asp"-->

<%  
'讀取參數值
ToPage= cint(Request("ToPage"))


pjno	  = trim(Request("pjno"))		      '專利代號
patentNo  = trim(Request("patentNo"))        '專利號
pjna    = trim(Request("pjna"))              '專利名稱
country =  trim(Request("country"))          '國別
patentType = trim(Request("patentType"))     '專利類別
IPC = trim(Request("ipc"))                   'IPC
patentApplydate = trim(Request("patentApplydate"))       '專利申請日
applydate_2 = trim(Request("applydate_2"))       '專利申請日
apply = trim(Request("apply"))               '申請類型
  
if pjno="" then	pjno=trim(request.form("pjno"))
if patentNo="" then	patentNo=trim(request.form("patentNo"))
if pjna="" then	pjna=trim(request.form("pjna"))
if country="" then	country=trim(request.form("country"))
if patentType="" then	patentType=trim(request.form("patentType"))
if IPC="" then	IPC=trim(request.form("IPC"))
if patentApplydate="" then	patentApplydate=trim(request.form("patentApplydate"))
if applydate_2="" then	applydate_2=trim(request.form("applydate_2"))
if apply="" then	apply=trim(request.form("apply"))

SQLSTR = "SELECT A.* FROM proj A"
condition = ""


    if pjno <> "" then
        condition = condition + " AND A.pjno LIKE '%" & pjno & "%'"
    end if

    if patentNo <> "" then
        condition = condition + " AND A.patentNo LIKE '%" & patentNo & "%'"
    end if

    if pjna <> "" then
        condition = condition + " AND A.pjna LIKE '%" & pjna & "%'"
    end if


    if country <> "" then
        condition = condition + " AND A.country = '" & country & "'"
    end if

    if patentType <> "" then
        condition = condition + " AND A.patentType = '" & patentType & "'"
    end if

    if IPC <> "" then
        condition = condition + " AND A.IPC LIKE '%" & IPC & "%'"
    end if


    if patentApplydate <> "" then
        condition = condition + " AND A.patentApplydate = '" & patentApplydate & "'"
    end if


    '案件種類

    '判決
    if apply="1" then
        SQLSTR = "SELECT A.* FROM [patDB].[dbo].[proj] A LEFT JOIN [patDB].[dbo].[adjudication] B ON a.pjno = b.pjno "
        condition = conditioin + " AND A.pjno LIKE 'A%'"

        adjudgeCourt = trim(Request.form("adjudgeCourt"))             '裁決法院
        acceptDate = trim(Request.form("acceptDate"))                 '受理日期
        adjudgeDate = trim(Request.form("adjudgeDate"))               '判決日期
        plaintiff = trim(Request.form("plaintiff"))                   '原告
        defendant = trim(Request.form("defendant"))                   '被告
        requestPrice  = trim(Request.form("requestPrice"))            '原告要求金額
        adjudgePrice = trim(Request.form("adjudgePrice"))             '判賠金額
        usPatents = trim(Request.form("usPatents"))                   '引證前案數
        foreignPatents = trim(Request.form("foreignPatents"))         '引證前案數
        citeCount = trim(Request.form("citeCount"))                   '引證文獻數
        claim  = trim(Request.form("claim"))                          '總項數
        idClaim = trim(Request.form("idClaim"))                       '獨立項總項數
        citedCount = trim(Request.form("citedCount"))                 '備引證總數
        globalEqPatents = trim(Request.form("globalEqPatents"))       '全球專利家族數量

        if adjudgeCourt <> "" then condition = condition + " AND B.adjudgeCourt LIKE '%" + adjudgeCourt + "%'"
        if acceptDate <> "" then condition = condition + " AND B.acceptDate = '" + acceptDate + "'"
        if adjudgeDate <> "" then condition = condition + " AND B.adjudgeDate = '%" + adjudgeDate + "%'"
        if plaintiff <> "" then condition = condition + " AND B.plaintiff LIKE '%" + plaintiff + "%'"
        if defendant <> "" then condition = condition + " AND B.defendant LIKE '%" + defendant + "%'"
        if requestPrice <> "" then condition = condition + " AND B.requestPrice = " + requestPrice 
        if adjudgePrice <> "" then condition = condition + " AND B.adjudgePrice = " + adjudgePrice 
        if usPatents <> "" then condition = condition + " AND B.usPatents = " + usPatents 
        if foreignPatents <> "" then condition = condition + " AND B.foreignPatents = " + foreignPatents
        if citeCount <> "" then condition = condition + " AND B.citeCount =" + citeCount 
        if claim <> "" then condition = condition + " AND B.claim =" + claim 
        if idClaim <> "" then condition = condition + " AND B.idClaim =" + idClaim
        if citedCount <> "" then condition = condition + " AND B.citedCount = " + citedCount
        if globalEqPatents <> "" then condition = condition + " AND B.globalEqPatents =" + globalEqPatents
      
    end if      


      

    '交易
    if apply="2" then
        SQLSTR = "SELECT A.* FROM [patDB].[dbo].[proj] A LEFT JOIN [patDB].[dbo].[TRANSACTION] B ON a.pjno = b.pjno "
        condition = condition + " AND A.pjno LIKE 'B%'"

        buyer = trim(Request.form("buyer"))                    '買方
        seller = trim(Request.form("seller"))                   '賣方
        transType = trim(Request.form("transType"))                '交易性質
        dealDate = trim(Request.form("dealDate"))                 '成交日期
        dealPrice = trim(Request.form("dealPrice"))                '成交金額
        dealPatentCount = trim(Request.form("dealPatentCount"))          '專利件數
        avgPatentPrice = trim(Request.form("avgPatentPrice"))           '一件專利均價
        caseSource = trim(Request.form("caseSource"))               '來源說明
        techInno = trim(Request.form("techInno"))                 '技術創新
        techCompetence = trim(Request.form("techCompetence_1"))           '技術競爭
        commercialization = trim(Request.form("commercialization_1"))        '商品化
        rddegree = trim(Request.form("rddegree"))                 '研發程度
        protectionExtent = trim(Request.form("protectionExtent_1"))         '保護範圍
        designAround = trim(Request.form("designAround_1"))             '迴避設計
        infringeIdentify = trim(Request.form("infringeIdentify_1"))         '侵權鑑定
        remainYear = trim(Request.form("remainYear_1"))               '剩餘年分
        inIndustryExpand = trim(Request.form("inIndustryExpand_1"))         '本業寬廣
        outIndustryExpand = trim(Request.form("outIndustryExpand_1"))        '異業寬廣
        infringement = trim(Request.form("infringement_1"))             '是否侵權
        techName  = trim(Request.form("techName_1"))        '技術名稱
        techField = trim(Request.form("techField_1"))       '技術領域

        if buyer <> "" then condition = condition + " AND B.buyer LIKE '%" + buyer + "%'"
        if seller <> "" then condition = condition + " AND B.seller LIKE '%" + seller + "%'"
        if transType <> "" then condition = condition + " AND B.transType LIKE '%" + transType + "%'"
        if dealDate <> "" then condition = condition + " AND B.dealDate = '" + dealDate + "'"
        if dealPrice <> "" then condition = condition + " AND B.dealPrice = " + dealPrice 
        if dealPatentCount <> "" then condition = condition + " AND B.dealPatentCount = " + dealPatentCount 
        if avgPatentPrice <> "" then condition = condition + " AND B.avgPatentPrice = " + avgPatentPrice 
        if caseSource <> "" then condition = condition + " AND B.caseSource LIKE '%" + caseSource + "%'"
        if techInno <> "" then condition = condition + " AND B.techInno = " + techInno 
        if techCompetence <> "" then condition = condition + " AND B.techCompetence = " + techCompetence 
        if commercialization <> "" then condition = condition + " AND B.commercialization = " + commercialization 
        if rddegree <> "" then condition = condition + " AND B.rddegree = " + rddegree 
        if protectionExtent <> "" then condition = condition + " AND B.protectionExtent = " + protectionExtent 
        if designAround <> "" then condition = condition + " AND B.designAround = " + designAround 
        if infringeIdentify <> "" then condition = condition + " AND B.infringeIdentify = " + infringeIdentify 
        if remainYear <> "" then condition = condition + " AND B.remainYear = " + remainYear 
        if inIndustryExpand <> "" then condition = condition + " AND B.inIndustryExpand = " + inIndustryExpand 
        if outIndustryExpand <> "" then condition = condition + " AND B.outIndustryExpand = " + outIndustryExpand 
        if infringement <> "" then condition = condition + " AND B.infringement = " + infringement 
        if techName <> "" then condition = condition + "AND b.techName LIKE '%" + techName + "%'"  '技術名稱
        if techField <> "" then condition = condition + "AND b.techField LIKE '%" + techField + "%'"  '技術名稱


    end if

    '冠亞鑑價
    if apply="3" then
        SQLSTR = "SELECT A.* FROM [patDB].[dbo].[proj] A LEFT JOIN [patDB].[dbo].[valuation] B ON a.pjno = b.pjno "
        condition = condition + " AND A.pjno LIKE 'C%'"     
        
        consignor = trim(Request.form("consignor"))                     '委託方
        consignorid = trim(Request.form("consignorid"))                     '委託方
        valuatePrice = trim(Request.form("valuatePrice"))               '鑑價金額
        currencyid = trim(Request.form("currencyid"))                   '幣別代碼
        valuateDate = trim(Request.form("valuateDate"))                 '鑑價日期
        techInfo = trim(Request.form("techInfo"))                       '技術創新
        techCompetence = trim(Request.form("techCompetence"))           '技術競爭
        commercialization = trim(Request.form("commercialization"))     '商品化
        redegree = trim(Request.form("redegree"))                       '研發程度
        protectionExtent = trim(Request.form("protectionExtent"))       '保護範圍
        designAround = trim(Request.form("designAround"))               '迴避設計
        infringeIdentify = trim(Request.form("infringeIdentify"))       '侵權鑑定
        remainYear = trim(Request.form("remainYear"))                   '剩餘年分
        inIndustryExpand = trim(Request.form("inIndustryExpand"))       '本業寬廣
        outIndustryExpand = trim(Request.form("outIndustryExpand"))     '異業寬廣
        infringement = trim(Request.form("infringement"))               '是否侵權
        totalScore = trim(Request.form("totalScore"))                   '總評分
        techName  = trim(Request.form("techName"))        '技術名稱
        techField = trim(Request.form("techField"))       '技術領域


        if consignor <> "" then condition = condition + " AND B.consignor LIKE '%" + consignor + "%'"
        if consignorid <> "" then condition = condition + " AND B.consignorid LIKE '%" + consignorid + "%'"
        if valuatePrice <> "" then condition = condition + " AND B.valuatePrice  = " + valuatePrice 
        if currencyid <> "" then condition = condition + " AND B.currencyid = '%" + currencyid + "%'"
        if valuateDate <> "" then condition = condition + " AND B.valuateDate = '" + valuateDate + "'"
        if techInfo <> "" then condition = condition + " AND B.techInfo  = " + techInfo 
        if techCompetence <> "" then condition = condition + " AND B.techCompetence = " + techCompetence 
        if commercialization <> "" then condition = condition + " AND B.commercialization  = " + commercialization 
        if redegree <> "" then condition = condition + " AND B.redegree = " + redegree 
        if protectionExtent <> "" then condition = condition + " AND B.protectionExtent = " + protectionExtent 
        if designAround <> "" then condition = condition + " AND B.designAround = " + designAround 
        if infringeIdentify <> "" then condition = condition + " AND B.infringeIdentify = " + infringeIdentify 
        if remainYear <> "" then condition = condition + " AND B.remainYear = " + remainYear 
        if inIndustryExpand <> "" then condition = condition + " AND B.inIndustryExpand = " + inIndustryExpand 
        if outIndustryExpand <> "" then condition = condition + " AND B.outIndustryExpand = " + outIndustryExpand 
        if infringement <> "" then condition = condition + " AND B.infringement = " + infringement 
        if totalScore <> "" then condition = condition + " AND B.totalScore  = " + totalScore
        if techName <> "" then condition = condition + "AND b.techName LIKE '%" + techName + "%'"   '技術名稱
        if techField <> "" then condition = condition + "AND b.techField = '%" + techField + "%'"  '技術名稱

       
    end if


    if condition = "" then 
	    SQLSTR=SQLSTR+" ORDER BY A.pjno"
    else
	    SQLSTR = SQLSTR + " WHERE 1=1 " + condition +" ORDER BY A.pjno"
    end if

    'response.write(SQLSTR)


Set rs = Server.CreateObject("ADODB.Recordset")

'SQLSTR = "SELECT * FROM PROJ"
rs.Open SQLSTR,conn,3,1


 '資料筆數設定 
  Count=Rs.RecordCount
 
 '頁數判斷 
  If Count<>0 then
    If Count/20 > (Count\20) then
       TotalPage=(Count\20)+1
    Else 
       TotalPage=(Count\20)
    End If
    Rs.MoveFirst
    If ToPage="" or ToPage=Empty then
    ToPage=1
    End If    
    Rs.Move (ToPage-1)*20
  Else
    TotalPage=1
  End If
%>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script src="lib/jquery-1.11.2.min.js"></script>
<title>專案管理</title>
</head>

<body>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" >
    <tr>
      <td width="100%" align="center" valign="top">
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" >
          <tr>
            <td width="100%">　</td>
          </tr>
        </table>
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" height="100%">
        <tr>
          <td width="100" align="center" valign="top">
          <!--#include file="lv.asp"-->
          </td>
          <td align="center" valign="top">
               <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="95%" >
                   <form>
				   
			   
<input type="hidden" name="pjno" value="<%=pjno%>">
<input type="hidden" name="patentNo" value="<%=patentNo%>">
<input type="hidden" name="pjna" value="<%=pjna%>">
<input type="hidden" name="country" value="<%=country%>">
<input type="hidden" name="patentType" value="<%=patentType%>">
<input type="hidden" name="patentApplydate" value="<%=patentApplydate%>">
<input type="hidden" name="applydate_2" value="<%=applydate_2%>">
<input type="hidden" name="apply" value="<%=apply%>">
<input type="hidden" name="IPC" value="<%=IPC%>">


                    <tr>
                      <td width="100%" align="left">第 
                      <% Response.Write "<select name=ToPage onchange=this.form.submit() >"
                         For i=1 to TotalPage 
                         If i<>ToPage then 
                         Response.Write "<option value="""& i & """ >" & i & "</optin>"
                         Else
                         Response.Write "<option value="""& i & """ selected >" & i & "</optin>"
                         End If 
                         Next 
                         Response.Write "</select>"
                      %>頁，總共<%=Count %> 筆，總共<%=TotalPage %>頁，每頁20筆</td>
                    </tr>
                   </form>
                  </table>
                  <table border="0" cellspacing="3" style="border-collapse: collapse" width="95%" bgcolor="E7F2FF">
                    <tr>
                     
                      <td width="6%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>專利代號</b></font>
                      </td>
					  <td width="15%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>專利號</b></font>
                      </td>
                      <td width="15%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>專利名稱</b></font>
                      </td>
                      <td width="10%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>國別</b></font>
                      </td>                      
                      <td width="8%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>專利類別</b></font>
                      </td>
                      <td width="8%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>IPC</b></font>
                      </td>
                      <td width="5%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>專利申請日</b></font>
                      </td>
                      <td width="5%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>編輯</b></font>
                      </td>
                    </tr>
                    <% i=0
                       If (not Rs.Eof) then
					   					   
                       Do while not Rs.Eof and i<20
					   
					   i=i+1
                       If (i mod 2)=1 Then
                        Response.Write "<tr bgcolor=#FFFFFF>"
                       Else
                        Response.Write "<tr>"
                       end if
                                              
                       Response.Write "<td align=left>" & Rs("pjno")       & "</td>"
                       Response.Write "<td align=left>" & Rs("patentNo")       & "</td>"
                       Response.Write "<td align=left>" & Rs("pjna")       & "</td>"                       
                       
                       if rs("country")<>"" then
                       Set idno = Server.CreateObject("ADODB.Recordset")
		                idno.Open "Select * From country where countryid='" & Rs("country") & "'" ,conn,1,1
                        Response.Write "<td align=center>" & idno("countryno") & "-" &idno("countryname") & "</td>"
                       else
                        Response.Write "<td align=center>-</td>"
                       end if
                       Response.Write "<td align=center>" & Rs("patentType")       & "</td>"
                       Response.Write "<td align=center>" & Rs("IPC")       & "</td>"                       
                       Response.Write "<td>" & Rs("patentApplydate")       & "</td>"

                       if session("idno")=47 then
                         Response.Write "<td align=center><a href=chpj2.asp?pjno=" & rs("pjno") & ">查看修改</a> | <a href=""#"" class=""delpj"" value="""&rs("pjno")&""">刪除</a> </td>"                         
                       else
                         Response.Write "<td align=center><a href=chpj2.asp?pjno=" & rs("pjno") & ">查看修改</a></td>" 
                       end if

                       Response.Write "</tr>"
                       Rs.MoveNext
                       Loop	
                       
                       
                       Else
                       Response.Write "<tr bgcolor=#FFFFFF>"
                       Response.Write "<td align=center colspan=11 bgcolor=#66ccff>查無資料</td>"
                               
                       Response.Write "</tr>"
                       End If

                     %>
                  </table>
                  </center>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
  </center>
</div>

</body>
<script>





    $(function () {
        $(".delPj").on("click", function () {
            var that = $(this);
            val = that.attr("value");
            if (window.confirm('確定要刪除')) {
                delProj(val, that);
            }
        });

        function delProj(pjno, obj) {

            var setting = {
                url: "deldata.asp",
                type: "get",
                dataType: "html",
                data: { action: "delPj", pjno: pjno }
            },
            def = $.ajax(setting);

            def.done(function (data) {
                if(data.substr(data.length - 3, 3)==200){
                obj.parent().parent().remove();
                }
            }).fail(function (data) {
                console.log("ajax is fail");
            }).always(function (data) {
                //$.unblockUI();
                console.log("always");
            })
        }

    });
</script>
</html>
<%
rs.close
%>
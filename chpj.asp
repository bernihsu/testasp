<!--#include file="function.asp"-->
<!--#include file="include/bs_function.asp"-->

<%  
'Ū���Ѽƭ�
ToPage= cint(Request("ToPage"))


pjno	  = trim(Request("pjno"))		      '�M�Q�N��
patentNo  = trim(Request("patentNo"))        '�M�Q��
pjna    = trim(Request("pjna"))              '�M�Q�W��
country =  trim(Request("country"))          '��O
patentType = trim(Request("patentType"))     '�M�Q���O
IPC = trim(Request("ipc"))                   'IPC
patentApplydate = trim(Request("patentApplydate"))       '�M�Q�ӽФ�
applydate_2 = trim(Request("applydate_2"))       '�M�Q�ӽФ�
apply = trim(Request("apply"))               '�ӽ�����
  
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


    '�ץ����

    '�P�M
    if apply="1" then
        SQLSTR = "SELECT A.* FROM [patDB].[dbo].[proj] A LEFT JOIN [patDB].[dbo].[adjudication] B ON a.pjno = b.pjno "
        condition = conditioin + " AND A.pjno LIKE 'A%'"

        adjudgeCourt = trim(Request.form("adjudgeCourt"))             '���M�k�|
        acceptDate = trim(Request.form("acceptDate"))                 '���z���
        adjudgeDate = trim(Request.form("adjudgeDate"))               '�P�M���
        plaintiff = trim(Request.form("plaintiff"))                   '��i
        defendant = trim(Request.form("defendant"))                   '�Q�i
        requestPrice  = trim(Request.form("requestPrice"))            '��i�n�D���B
        adjudgePrice = trim(Request.form("adjudgePrice"))             '�P�ߪ��B
        usPatents = trim(Request.form("usPatents"))                   '���ҫe�׼�
        foreignPatents = trim(Request.form("foreignPatents"))         '���ҫe�׼�
        citeCount = trim(Request.form("citeCount"))                   '���Ҥ��m��
        claim  = trim(Request.form("claim"))                          '�`����
        idClaim = trim(Request.form("idClaim"))                       '�W�߶��`����
        citedCount = trim(Request.form("citedCount"))                 '�Ƥ����`��
        globalEqPatents = trim(Request.form("globalEqPatents"))       '���y�M�Q�a�ڼƶq

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


      

    '���
    if apply="2" then
        SQLSTR = "SELECT A.* FROM [patDB].[dbo].[proj] A LEFT JOIN [patDB].[dbo].[TRANSACTION] B ON a.pjno = b.pjno "
        condition = condition + " AND A.pjno LIKE 'B%'"

        buyer = trim(Request.form("buyer"))                    '�R��
        seller = trim(Request.form("seller"))                   '���
        transType = trim(Request.form("transType"))                '����ʽ�
        dealDate = trim(Request.form("dealDate"))                 '������
        dealPrice = trim(Request.form("dealPrice"))                '������B
        dealPatentCount = trim(Request.form("dealPatentCount"))          '�M�Q���
        avgPatentPrice = trim(Request.form("avgPatentPrice"))           '�@��M�Q����
        caseSource = trim(Request.form("caseSource"))               '�ӷ�����
        techInno = trim(Request.form("techInno"))                 '�޳N�зs
        techCompetence = trim(Request.form("techCompetence_1"))           '�޳N�v��
        commercialization = trim(Request.form("commercialization_1"))        '�ӫ~��
        rddegree = trim(Request.form("rddegree"))                 '��o�{��
        protectionExtent = trim(Request.form("protectionExtent_1"))         '�O�@�d��
        designAround = trim(Request.form("designAround_1"))             '�j�׳]�p
        infringeIdentify = trim(Request.form("infringeIdentify_1"))         '�I�vŲ�w
        remainYear = trim(Request.form("remainYear_1"))               '�Ѿl�~��
        inIndustryExpand = trim(Request.form("inIndustryExpand_1"))         '���~�e�s
        outIndustryExpand = trim(Request.form("outIndustryExpand_1"))        '���~�e�s
        infringement = trim(Request.form("infringement_1"))             '�O�_�I�v
        techName  = trim(Request.form("techName_1"))        '�޳N�W��
        techField = trim(Request.form("techField_1"))       '�޳N���

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
        if techName <> "" then condition = condition + "AND b.techName LIKE '%" + techName + "%'"  '�޳N�W��
        if techField <> "" then condition = condition + "AND b.techField LIKE '%" + techField + "%'"  '�޳N�W��


    end if

    '�a��Ų��
    if apply="3" then
        SQLSTR = "SELECT A.* FROM [patDB].[dbo].[proj] A LEFT JOIN [patDB].[dbo].[valuation] B ON a.pjno = b.pjno "
        condition = condition + " AND A.pjno LIKE 'C%'"     
        
        consignor = trim(Request.form("consignor"))                     '�e�U��
        consignorid = trim(Request.form("consignorid"))                     '�e�U��
        valuatePrice = trim(Request.form("valuatePrice"))               'Ų�����B
        currencyid = trim(Request.form("currencyid"))                   '���O�N�X
        valuateDate = trim(Request.form("valuateDate"))                 'Ų�����
        techInfo = trim(Request.form("techInfo"))                       '�޳N�зs
        techCompetence = trim(Request.form("techCompetence"))           '�޳N�v��
        commercialization = trim(Request.form("commercialization"))     '�ӫ~��
        redegree = trim(Request.form("redegree"))                       '��o�{��
        protectionExtent = trim(Request.form("protectionExtent"))       '�O�@�d��
        designAround = trim(Request.form("designAround"))               '�j�׳]�p
        infringeIdentify = trim(Request.form("infringeIdentify"))       '�I�vŲ�w
        remainYear = trim(Request.form("remainYear"))                   '�Ѿl�~��
        inIndustryExpand = trim(Request.form("inIndustryExpand"))       '���~�e�s
        outIndustryExpand = trim(Request.form("outIndustryExpand"))     '���~�e�s
        infringement = trim(Request.form("infringement"))               '�O�_�I�v
        totalScore = trim(Request.form("totalScore"))                   '�`����
        techName  = trim(Request.form("techName"))        '�޳N�W��
        techField = trim(Request.form("techField"))       '�޳N���


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
        if techName <> "" then condition = condition + "AND b.techName LIKE '%" + techName + "%'"   '�޳N�W��
        if techField <> "" then condition = condition + "AND b.techField = '%" + techField + "%'"  '�޳N�W��

       
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


 '��Ƶ��Ƴ]�w 
  Count=Rs.RecordCount
 
 '���ƧP�_ 
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
<title>�M�׺޲z</title>
</head>

<body>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" >
    <tr>
      <td width="100%" align="center" valign="top">
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" >
          <tr>
            <td width="100%">�@</td>
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
                      <td width="100%" align="left">�� 
                      <% Response.Write "<select name=ToPage onchange=this.form.submit() >"
                         For i=1 to TotalPage 
                         If i<>ToPage then 
                         Response.Write "<option value="""& i & """ >" & i & "</optin>"
                         Else
                         Response.Write "<option value="""& i & """ selected >" & i & "</optin>"
                         End If 
                         Next 
                         Response.Write "</select>"
                      %>���A�`�@<%=Count %> ���A�`�@<%=TotalPage %>���A�C��20��</td>
                    </tr>
                   </form>
                  </table>
                  <table border="0" cellspacing="3" style="border-collapse: collapse" width="95%" bgcolor="E7F2FF">
                    <tr>
                     
                      <td width="6%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>�M�Q�N��</b></font>
                      </td>
					  <td width="15%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>�M�Q��</b></font>
                      </td>
                      <td width="15%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>�M�Q�W��</b></font>
                      </td>
                      <td width="10%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>��O</b></font>
                      </td>                      
                      <td width="8%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>�M�Q���O</b></font>
                      </td>
                      <td width="8%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>IPC</b></font>
                      </td>
                      <td width="5%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>�M�Q�ӽФ�</b></font>
                      </td>
                      <td width="5%" align="center" bgcolor="#0066FF">
                        <font color="#FFFFFF"><b>�s��</b></font>
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
                         Response.Write "<td align=center><a href=chpj2.asp?pjno=" & rs("pjno") & ">�d�ݭק�</a> | <a href=""#"" class=""delpj"" value="""&rs("pjno")&""">�R��</a> </td>"                         
                       else
                         Response.Write "<td align=center><a href=chpj2.asp?pjno=" & rs("pjno") & ">�d�ݭק�</a></td>" 
                       end if

                       Response.Write "</tr>"
                       Rs.MoveNext
                       Loop	
                       
                       
                       Else
                       Response.Write "<tr bgcolor=#FFFFFF>"
                       Response.Write "<td align=center colspan=11 bgcolor=#66ccff>�d�L���</td>"
                               
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
            if (window.confirm('�T�w�n�R��')) {
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
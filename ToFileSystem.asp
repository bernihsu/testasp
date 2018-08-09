<!--#include file="function.asp"-->
<!--#include file="include/bs_function.asp"-->
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
dbType = objUpload.Fields("apply").Value

if(dbType="1") then
    strFileName = "adjudication"
elseif (dbType="2") then
    strFileName = "transaction"
elseif (dbType="3") then
    strFileName = "valuation"
else 
    Response.Redirect "uploadExcel.asp?status=error"
end if

'strFileName = objUpload.Fields("File1").FileName

strPath = Server.MapPath("Upload2") & "\uploads\tmp_" & strFileName & ".xls"

' Save the binary data to the file system
objUpload("File1").SaveAs strPath

' Release upload object from memory
Set objUpload = Nothing


if(dbType="1")then
    'adjudication 判決    
    Set custcount = Server.CreateObject("ADODB.Recordset")
    custcount.Open "Select count(*) AS count From proj where pjno LIKE 'A%'",conn,1,3

    nowcount=custcount("count")

    custcount.close

    Set oConn = Server.CreateObject("ADODB.Connection")
    Set oST = Server.CreateObject("ADODB.Stream")

    'fn1表示上傳的excel檔
    fpath=Server.MapPath("upload2/uploads/tmp_adjudication.xls")
    oConn.Open "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & fpath

    Set oRS = oConn.Execute("select * from [工作表1$]")
    
        oRS.Save oST, 1
        oRS.Close

        Set oRS.Source = oST
        oRS.Open
    
    do until ors.EOF
      if(not oRs.Fields.Item(0).value="") then

          i = nowcount+1    
          nowcount = right("00000" & i,5)
          newpjno= "A" + nowcount

        country =  oRs.Fields.Item(3).value
        if country <>"" then
            Set idno = Server.CreateObject("ADODB.Recordset")
            idno.Open "Select * From country WHERE countryno='" & country & "' OR countryname='" & country & "'" ,conn,1,1
            
            If (not idno.Eof) then
                countryid = idno("countryid")
            else
                countryid = ""
            end if
        else
            countryid = ""
        end if

         if(isnull(oRs.Fields.Item(1).value)) then 
            patentno = ""
         else
            patentno = replace(oRs.Fields.Item(1).value, "'","''")
         end if

         if(isnull(oRs.Fields.Item(2).value)) then 
            pjna = ""
         else
            pjna = replace(oRs.Fields.Item(2).value, "'","''") 
         end if

         if(isnull(oRs.Fields.Item(4).value)) then 
            patentType = ""
         else
            patentType = replace(oRs.Fields.Item(4).value, "'","''")
         end if

         if(isnull(oRs.Fields.Item(5).value)) then 
            IPC = ""
         else
            IPC = replace(oRs.Fields.Item(5).value, "'","''")
         end if

         if(isnull(oRs.Fields.Item(6).value)) then 
            patentApplyNo = ""
         else
            patentApplyNo = replace(oRs.Fields.Item(6).value, "'","''")
         end if

       
         'response.write(patentno& "-"  & oRs.Fields.Item(6).value  & oRs.Fields.Item(7).value & "-" & len(oRs.Fields.Item(7).value)&"<br>")

         if(isnull(oRs.Fields.Item(7).value)) then 
            patentApplydate = "null"
         else
         

            if(len(oRs.Fields.Item(7).value)>10) then 
            
               patentApplydate = split(oRs.Fields.Item(7).value, " ")               
               
            else
                patentApplydate = "'" & oRs.Fields.Item(7).value & "'"
            end if
            
            
         end if        
        

          commSQL = "INSERT INTO [patdb].[dbo].[proj] (pjno,patentno,pjna,country,patentType,IPC,patentApplyNo,patentApplydate) VALUES ('" & newpjno & "','" & patentno & "','" & pjna & "','" & countryid & "','" & patentType & "','" & IPC & "','" & patentApplyNo & "'," & patentApplydate & ")"
      
          adjSQL = "INSERT INTO [patdb].[dbo].[adjudication](pjno,adjudgeCourt,acceptDate,adjudgeDate,plaintiff,defendant,requestPrice,adjudgePrice,currencyid,usPatents,foreignPatents,citeCount,claim,idClaim,citedCount,globalEqPatents,note) VALUES "
                                       
          adjSQL = adjSQL & "('" & newpjno & "'"
      
          for i=8 to 23

            if(isnull(oRs.Fields.Item(i).value)) then
                adjSQL = adjSQL & ",null "
            else
                adjSQL = adjSQL & ",'" & replace(oRs.Fields.Item(i).value, "'","''") & "'"
                
            end if

            'adjSQL = adjSQL & ",'" & oRs.Fields.Item(i).value & "'"
          next 

          adjSQL = adjSQL &  ")"
          total = total + 1
          conn.Execute(commSQL)
          conn.Execute(adjSQL)

          
      response.write("專利號:" & patentno & "專利名稱:" & pjna & "專利類別:" & patentType & "IPC:" & IPC & "<br>")
'      response.writE(adjSQL & "<br>")
      end if
    oRs.MoveNext
    loop

    oRs.close
    oConn.close
elseif(dbType="2") then
'response.write("交易")
    'transaction 交易
    Set custcount = Server.CreateObject("ADODB.Recordset")
    custcount.Open "Select count(*) AS count From proj where pjno LIKE 'B%'",conn,1,3

    nowcount=custcount("count")

    custcount.close

    Set oConn = Server.CreateObject("ADODB.Connection")
    Set oST = Server.CreateObject("ADODB.Stream")

    'fn1表示上傳的excel檔
    fpath=Server.MapPath("upload2/uploads/tmp_transaction.xls")
    oConn.Open "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & fpath

    Set oRS = oConn.Execute("select * from [工作表1$]")
    
        oRS.Save oST, 1
        oRS.Close

        Set oRS.Source = oST
        oRS.Open


    do until ors.EOF
      
      if(not oRs.Fields.Item(0).value="") then

          i = nowcount+1    
          nowcount = right("00000" & i,5)
          newpjno= "B" + nowcount

          country =  oRs.Fields.Item(4).value
          if country <>"" then
            Set idno = Server.CreateObject("ADODB.Recordset")
            idno.Open "Select * From country where countryname='" & country & "' OR countryno='" & country & "'" ,conn,1,1
            If (not idno.Eof) then
                countryid = idno("countryid")
            else
                countryid = ""
            end if
          else
            countryid = ""
          end if

          if(isnull(oRs.Fields.Item(1).value)) then 
            patentno = ""
          else
            patentno = replace(oRs.Fields.Item(1).value, "'","''")
          end if
            
          if(isnull(oRs.Fields.Item(2).value)) then 
            pjna = ""
          else
            pjna = replace(oRs.Fields.Item(2).value, "'","''")
          end if

          if(isnull(oRs.Fields.Item(5).value)) then 
            patentType = ""
          else
            patentType = replace(oRs.Fields.Item(5).value, "'","''")
          end if

          if(isnull(oRs.Fields.Item(6).value)) then 
            IPC = ""
          else
            IPC = replace(oRs.Fields.Item(6).value, "'","''")
          end if

          if(isnull(oRs.Fields.Item(7).value)) then 
            patentApplyNo = ""
          else
            patentApplyNo = replace(oRs.Fields.Item(7).value, "'","''")
          end if

          if (isnull(oRs.Fields.Item(9).value)) then
            patentApplydate = "null"
          else
            patentApplydate = "'" & oRs.Fields.Item(9).value & "'"
          end if

          if(isnull(oRs.Fields.Item(3).value)) then 
            techname = ""
          else
            techname = replace(oRs.Fields.Item(3).value, "'","''")
          end if

          if(isnull(oRs.Fields.Item(7).value)) then 
            techField = ""
          else
            techField = replace(oRs.Fields.Item(7).value, "'","''")
          end if

          commSQL = "INSERT INTO [patdb].[dbo].[proj] (pjno,patentno,pjna,country,patentType,IPC,patentApplyNo,patentApplydate) VALUES ('" & newpjno & "','" & patentno & "','" & pjna & "','" & countryid & "','" & patentType & "','" & IPC & "','" & patentApplyNo & "'," & patentApplydate & ")"
      
          adjSQL = "INSERT INTO [patdb].[dbo].[transaction](pjno,techname,techField,buyer,seller,transType,dealDate,dealPrice,currencyid,dealPatentCount,avgPatentPrice,caseSource,techInno,techCompetence,commercialization,rddegree,protectionExtent,designAround,infringeIdentify,remainYear,inIndustryExpand,outIndustryExpand,infringement,note) VALUES "
                                       
          adjSQL = adjSQL & "('" & newpjno & "','" & techname & "','" & techField & "'"
      
          for i=10 to 30

            if( oRs.Fields.Item(i).value<>"") then
            
                adjSQL = adjSQL & ",'" & replace(oRs.Fields.Item(i).value, "'","''") & "'"
            else
                adjSQL = adjSQL & ",'" & oRs.Fields.Item(i).value & "'"
                
            end if
           
          
          next 

          adjSQL = adjSQL &  ")"

          total = total + 1
          conn.Execute(commSQL)
          conn.Execute(adjSQL)


      response.write("專利號:" & patentno & "專利名稱:" & pjna & "專利類別:" & patentType & "IPC:" & IPC & "<br>")
      'response.writE(commSQL & "<br>")
      'response.writE(adjSQL & "<br>")
      end if
    oRs.MoveNext
    loop

    oRs.close
    oConn.close 
elseif(dbType="3") then
    'valuation 鑑價
    Set custcount = Server.CreateObject("ADODB.Recordset")
    custcount.Open "Select count(*) AS count From proj where pjno LIKE 'C%'",conn,1,3

    nowcount=custcount("count")

    custcount.close

    Set oConn = Server.CreateObject("ADODB.Connection")
    Set oST = Server.CreateObject("ADODB.Stream")

    'fn1表示上傳的excel檔
    fpath=Server.MapPath("upload2/uploads/tmp_valuation.xls")
    oConn.Open "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & fpath

    Set oRS = oConn.Execute("select * from [工作表1$]")
    
        oRS.Save oST, 1
        oRS.Close

        Set oRS.Source = oST
        oRS.Open


    do until ors.EOF
      if(not oRs.Fields.Item(0).value="") then

          i = nowcount+1    
          nowcount = right("00000" & i,5)
          newpjno= "C" + nowcount

          country =  oRs.Fields.Item(4).value
          if country <>"" then
            Set idno = Server.CreateObject("ADODB.Recordset")
            idno.Open "Select * From country where countryname='" & country & "' OR countryno='" & country & "'" ,conn,1,1
            If (not idno.Eof) then
                countryid = idno("countryid")
            else
                countryid = ""
            end if
          else
            countryid = ""
          end if

          if(isnull(oRs.Fields.Item(1).value)) then 
            patentno = ""
         else
            patentno = replace(oRs.Fields.Item(1).value, "'","''")
         end if

         if(isnull(oRs.Fields.Item(2).value)) then 
            pjna = ""
         else
            pjna = replace(oRs.Fields.Item(2).value, "'","''") 
         end if

         if(isnull(oRs.Fields.Item(5).value)) then 
            patentType = ""
         else
            patentType = replace(oRs.Fields.Item(5).value, "'","''")
         end if

         if(isnull(oRs.Fields.Item(6).value)) then 
            IPC = ""
         else
            IPC = replace(oRs.Fields.Item(6).value, "'","''")
         end if

         if(isnull(oRs.Fields.Item(7).value)) then 
            patentApplyNo = ""
         else
            patentApplyNo = replace(oRs.Fields.Item(7).value, "'","''")
         end if

         if(isnull(oRs.Fields.Item(9).value)) then 
            patentApplydate = "null"
         else
            patentApplydate = "'" & oRs.Fields.Item(9).value & "'"
         end if        
        
         
          if(isnull(oRs.Fields.Item(3).value)) then 
            techname = ""
          else
            techname = replace(oRs.Fields.Item(3).value, "'","''")
          end if

          if(isnull(oRs.Fields.Item(7).value)) then 
            techField = ""
          else
            techField = replace(oRs.Fields.Item(7).value, "'","''")
          end if   

          commSQL = "INSERT INTO [patdb].[dbo].[proj] (pjno,patentno,pjna,country,patentType,IPC,patentApplyNo,patentApplydate) VALUES ('" & newpjno & "','" & patentno & "','" & pjna & "','" & countryid & "','" & patentType & "','" & IPC & "','" & patentApplyNo & "'," & patentApplydate & ")"
      
          adjSQL = "INSERT INTO [patdb].[dbo].[valuation](pjno,techName,techField,consignor,consignorid,valuatePrice,currencyid,valuateDate,techInfo,techCompetence,commercialization,redegree,protectionExtent,designAround,infringeIdentify,remainYear,inIndustryExpand,outIndustryExpand,infringement,totalScore,note) VALUES "                                       

          adjSQL = adjSQL & "('" & newpjno & "','" & techName & "','" & techField & "'"
      
          for i=10 to 27
            if (isnull(oRs.Fields.Item(i).value)) then
                adjSQL = adjSQL & ", null "
            else
                adjSQL = adjSQL & ",'" & replace(oRs.Fields.Item(i).value, "'","''") & "'"
            end if
          next 

          adjSQL = adjSQL &  ")"

      
      'response.writE(commSQL & "<br>")
      'response.writE(adjSQL & "<br>")
      total = total + 1
          conn.Execute(commSQL)
          conn.Execute(adjSQL)
        response.write("專利號:<b>" & patentno & "</b> 專利名稱:<b>" & pjna & "</b>專利類別:<b>" & patentType & "</b> IPC:<b>" & IPC & "</b><br>")    

      end if
    oRs.MoveNext
    loop

    oRs.close
    oConn.close 
end if



%>

<hr />

<div>匯入完成，共匯入<%=total%>筆</div>
<br />
<div><a href="uploadExcel.asp?state=success">回到匯入頁面</a>，切勿重新整理頁面 (8秒後自動返回)</div>  

<script language=javascript>
    //alert('上傳成功');
   setTimeout(function () {

       location.href = "uploadExcel.asp?state=success"
   }, 8000);
</script>

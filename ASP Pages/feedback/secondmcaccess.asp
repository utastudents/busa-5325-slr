<!--#include file="adovbs.inc"-->

<!doctype html>
<html >
<head>

<title>My Online Survey</title>

</head>

<%
    Dim cid
    dim ans
    dim f, fs,fname

   'Check for correct variable lengths and type
    cid = Request.Form("cid")
    ans =  Request.Form("ans")
    
   
     if len(ans) <>0 then
        select case ans
         case "1", "2", "3", "4", "5"
         case else
           ans="0"
        end select
        ans = chr(64 + ans)
     else
        ans="missing"
        response.write("Please go back and choose an answer. It is missing at the moment. <br> ")
     end if

    if isnumeric(cid)=false then 
      response.write("Please go back and enter a computer number. It is missing or not a number at the moment. <br> ")
      cid="missing"
    else

       if (csng(cid)-cint(cid) <> 0 ) or (cint(cid)>72) or (cint(cid) < 1)  then 
          response.write( "<br> You entered " & cid & " for your computer number. <br>")
          response.write( "The computer number should be an integer from 1 to 72. <br>")
          response.write( " Please go back and enter a computer number with the correct value. <br> <br>")
          cid = "missing"
       end if
    end if


    if (ans<>"missing") and (cid <> "missing") then 
         Response.write("Your entered a computer number of " & cid & " and an answer of " & ans & "<br> <br>"  ) 
    end if

   'Send responses to text files
 

  
%>

  <% 'from update wiht include . asp file

      '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
      'VARIABLE DECLARATION
      '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

      Dim objConn
      Dim strPath
      Dim strConn
      Dim objRecSet
      
      Dim arrType
      Dim strSQL
      Dim rsTest
    '  
      call sbOpen("answers")
     
      
     'sub  UpdateDatabase(strTableName, strColumnName,  intRowId, strSavedValue, strKeyName)
      call UpdateDB("McAnswers", "mcanswer", cid, ans, "ID") 
   
      'call sbClose()

     '================== Subs start here =================================="

      sub sbOpen(strDataBaseName)
           Set objConn = CreateObject("ADODB.Connection") 
           strPath = Server.MapPath(strDataBaseName & ".accdb")
           strConn = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & strPath & ";"
           objConn.Open strConn  
           set objRecSet = createobject("ADODB.Recordset")
      end sub  

     sub UpdateDB(strTableName, strColumnName, intRowID, strSavedValue, strKeyName) 
      
        'sub UpdateDB(strAccessName, strTableName, strColumnName, intRowID, strSavedValue, strKeyName) 
        'call sbOpen(straccessname)
        'strTableName is the name of the table
        'strColumnName is the name of the column
        'intRowId gives the row of the table 
        'strSavedValue is value to be stored
        'strKeyName is the name of the Primary Key variable

        dim strQuery
        dim strPlayer

        

        strQuery = "select * " & strtablename & " where " & strkeyname &  " = " & intRowId  
        prhtml(strquery & "<br>")
         
          'objRecSet.CursorLocation=adUseServer
          'objRecSet.CursorType=adOpenStatic
          'objRecSet.LockType=adLockBatchOptimistic
          'objRecSet.open strQuery, objConn

         'objRecSet.Close


        prhtml("UPDATE " & strtablename & " Set " & strcolumnname & " = '" & strsavedValue & "' where " & strkeyname &  " = " & intRowId & ";") 
        strQuery = "Update " & strtablename & " Set " & strcolumnname & " = '" & strsavedValue & "' where " & strkeyname &  " = " & intRowId & ";"
          objRecSet.CursorLocation=adUseServer
          objRecSet.CursorType=adOpenStatic
          objRecSet.LockType=adLockBatchOptimistic
          objRecSet.open strQuery, objConn

         
           response.write "<br>" & objrecset.state
        '  objRecSet.Close
        ' Set objRecSet = Nothing
        ' objConn.Close
        ' Set objConn = Nothing
      
       'shell
     end sub

     sub sbClose()
       objRecSet.Close
       Set objRecSet = Nothing
       objConn.Close
       Set objConn = Nothing
     end sub

     sub Prhtml(strTemp)
      response.write strTemp
     end sub


  %>
  

Thank You <br>

</body>

</html>
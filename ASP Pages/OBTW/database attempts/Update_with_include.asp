<!--#include file="adovbs.inc"-->

<!doctype html>
<html>
<head>
<meta charset="utf-8">

<title>Test</title>
</head>
<body>
<% 
     response.write " Current Time is " & time & "<br>"
     'the following needs to be replaced under in head area
     '<META HTTP-EQUIV="refresh" CONTENT="5">
     
 
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
      
      call sbOpen("obtw")


          'UpdateDatabase(strTableName, intPlayerNumber,  intRowId, intSavedValue, strKeyName)
      call updateDatabase(   "Played",          1,            3,         1,          "ID"     )

      


       call sbClose()

      '=======================  'subroutines start here '=====================
     
      %>



     <p><!--#include file=sub_updateDatabase.inc--></p>

    <%  
      
     'other subs will go here     
      

    %>
</body>

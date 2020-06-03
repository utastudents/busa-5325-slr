<!--#include file="adovbs.inc"-->

<!doctype html>
<html>
<head>
<meta charset="utf-8">
<META HTTP-EQUIV="refresh" CONTENT="5">
<title>Test</title>
</head>
<body>
<% 
     response.write " Current Time is " & time & "<br>"
      
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
      
      call sbOpen()

      ' call UpdateDatabase(intPlayerNumber, strTableName, intID, intSavedValue, strVariableName, strKeyVariableName)
        call UpdateDatabase(      3,           "Played",       2,        1,          "Player3",         "id")
      
      call sbClose()

      sub sbOpen()
           Set objConn = CreateObject("ADODB.Connection") 
           strPath = Server.MapPath("obtw.accdb")
           strConn = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & strPath & ";"
           objConn.Open strConn  
           set objRecSet = createobject("ADODB.Recordset")
      end sub

      
     
     sub UpdateDatabase(intPlayerNumber, strTableName, intID, intSavedValue, strVariableName, strKeyVariableName)
                dim strQuery
        strQuery = "Select * from " & strTableName & " where " & strKeyVariableName & " =" & intID
        objRecSet.CursorLocation=adUseServer
        objRecSet.CursorType=adOpenStatic
        objRecSet.LockType=adLockBatchOptimistic
        objRecSet.open strQuery, objConn

        response.write strVariableName & " has the value " & objRecSet("Player" & intPlayerNumber) & "<br/>"
        'objrecset.fields("player" & intPLayerNumber) = 1
        'objRecSet.update "player" & intPlayerNumber, 1

        objRecSet.Close
        strQuery = "UPDATE " & strTableName & " Set " & strVariableName & " = '" & intSavedValue & "' where " & strKeyVariableName & " = " & intID & ";"
        objRecSet.CursorLocation=adUseServer
        objRecSet.CursorType=adOpenStatic
        objRecSet.LockType=adLockBatchOptimistic
        objRecSet.open strQuery, objConn

        strQuery = "Select * from " & strTableName & " where " & strKeyVariableName & " =" & intID
        objRecSet.CursorLocation=adUseServer
        objRecSet.CursorType=adOpenStatic
        objRecSet.LockType=adLockBatchOptimistic
        objRecSet.open strQuery, objConn
        response.write strVariableName & " has the value " & objRecSet("Player" & intPlayerNumber)

     end sub


     sub sbClose()
       objRecSet.Close
       Set objRecSet = Nothing
       objConn.Close
       Set objConn = Nothing
     end sub


%>
</body>

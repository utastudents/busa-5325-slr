<%    
     sub sbOpen(strDataBaseName)
           Set objConn = CreateObject("ADODB.Connection") 
           strPath = Server.MapPath(strDataBaseName & ".accdb")
           strConn = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & strPath & ";"
           objConn.Open strConn  
           set objRecSet = createobject("ADODB.Recordset")
      end sub  
     
     sub UpdateDatabase(strTableName, intPlayerNumber,  intRowId, intSavedValue, strKeyName)
        'strTableName is the name of the table
        'intPlayerNumber gives you the suffix to identify the word "Player#" column. THis contains the plays the person has made
        'intRowId gives the row of the table 
        'intSavedValue is value to be stored
        'strKeyName is the name of the Primary Key variable

        dim strQuery
        dim strPlayer
      
        strQuery = "Select * from " & strTableName & " where " & strKeyName & "=" & intRowId
        objRecSet.CursorLocation=adUseServer
        objRecSet.CursorType=adOpenStatic
        objRecSet.LockType=adLockBatchOptimistic
        objRecSet.open strQuery, objConn

        Prhtml("Player " & intPlayerNumber & " has the value " & objRecSet("Player3") & "<br/>")


        objRecSet.Close
       
         strQuery = "UPDATE Played Set player" & intPlayerNumber & " = '" & intSavedValue & "' where id = " & intRowId & ";"
        
        objRecSet.CursorLocation=adUseServer
        objRecSet.CursorType=adOpenStatic
        objRecSet.LockType=adLockBatchOptimistic
        objRecSet.open strQuery, objConn

        
        strQuery = "Select * from " & strTableName & " where " & strKeyName & "=" & intRowId 

        'prhtml(strquery3)

        objRecSet.CursorLocation=adUseServer
        objRecSet.CursorType=adOpenStatic
        objRecSet.LockType=adLockBatchOptimistic
        objRecSet.open strQuery, objConn
        
         Prhtml("Player " & intPlayerNumber & " has the value " & objRecSet("Player3") & "<br/>")
        
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

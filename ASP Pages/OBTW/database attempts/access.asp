<% Option Explicit %>
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Test</title>
</head>
<body>
Inside the BODY<br/>
<%
response.write "Inside the ASP<br/>"

      
      response.write "Beginning job now..."
      
      '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
      'VARIABLE DECLARATION
      '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

      Dim objConn
      Dim strPath
      Dim strConn
      Dim strSQL
      Dim arrType
      Dim r, c
      Dim rsTest
      
      call sbOpen()

      sub sbOpen()
           Set objConn = CreateObject("ADODB.Connection") 
           strPath = Server.MapPath("garagemike.accdb")
           response.write "Database Connection Path: " & strPath & "</br>"
           strConn = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & strPath & ";"
           objConn.Open strConn  
           response.write "database open..."
           response.write "getting records..."
      end sub

                objConn.Close
    Set objConn = Nothing

%>
</body>
</html>


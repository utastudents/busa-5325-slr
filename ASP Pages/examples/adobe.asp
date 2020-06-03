
<job id="RH_job_1">

<script language-"VBScript">
response.write "Hello World"

      Option Explicit
      Wscript.Echo "Beginning job now..."
      
      '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
      ‘VARIABLE DECLARATION
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
           strPath = "https://wweb.uta.edu/faculty/eakin/asps/examples/garagemike.accdb"
           strConn = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & strPath & ";"
           objConn.Open strConn  
           <% response.write "database open..." %>
           Wscript.Echo "getting records..."
      end sub

</script>

</job>
<body>
  <%response.write "Hello" %>
</body>
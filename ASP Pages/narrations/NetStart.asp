
<html>


<body>


<h1> Narrated Excel Examples Using Netscape </h1>
<h2> Click on any Option Button and Press the Enter Key </h2>
<p>Best seen with a 1024x768 or higher screen resolution</p>

<Form action="page0Net.asp" method="post">
<TAble border=2>
<tr>
   <TD> Title            </TD>
</tr>
    <%
    Dim myConn
    Dim ConnString
    dim strPrefixAndNumber

    Set myConn=Server.CreateObject("ADODB.Connection")
    Set rsNar8=Server.CreateObject("ADODB.Recordset")

    myConn.Open"Driver={Microsoft Access Driver (*.mdb)};dbq=" & server.mappath("nar8.mdb")

    Set rsNar8=myConn.Execute("Select * from SlideList")


    Do while not rsNar8.EOF %>

   <TR>
     <% strPrefixAndNumber = rsnar8("Prefix") & "-" & rsnar8("NumberSlides") %>

     <TD>   <INPUT TYPE=RADIO NAME=GO VALUE=<%=strPrefixAndNumber%>> <%=rsnar8("Title")%> </TD>
    
   
   </tr>
       <%rsNar8.MoveNext%>
   <%Loop%>  
   
   <Tr>
     <td>
       
       <INPUT TYPE=submit VALUE="GO">
       <Input type=reset value="Cancel"> 
     </td>
  </tr>
</table>

</FORM>



</body>
</html>

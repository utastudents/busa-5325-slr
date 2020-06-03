
<html>


<body>


<Form action="trypage.asp" method="post">
<TAble border=2>

    <%

    dim i as integer
    
    i=1

    loop until i=5 %>

   <TR>
     

     <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="Joe"> </TD>
    
   
    </tr>
       
   <%
     i=i+1
     loop
      %>  
   
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

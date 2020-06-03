
<html>


<body>

<h2> Finding Normal Probabilities for Tables That Read 0 to a Z-value</h2>
<Form action="tableLookup.asp" method="post">
<% Randomize %>
<input Name="GivenZscore" value="<%=round(rnd()*3,2)%>" Type = "Hidden">
<input Name="GivenZscore2" value="<%=round(rnd()*3,2)%>" Type = "Hidden">
 
<INPUT TYPE=submit VALUE="GO">  <Input type=reset value="Cancel"> 
Choose and click option button and then click Go
<TAble border=2>
<tr>
   
</tr>
  
<TR>     <TD>   Find the probability that a standard normal variable is </TD>
<TR>     <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="A"> greater than Z1 </TD>
<TR>     <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="B"> less than Z1  </TD>
<TR>     <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="C"> inside the range Z1 to Z2   </TD>
<TR>     <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="D"> outside the range Z1 to Z2   </TD>
     
     
   
       
      
     </td>
  </tr>
</table>

</FORM>



</body>
</html>

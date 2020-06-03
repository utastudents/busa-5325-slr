
<html>


<body>

<h2> Finding Normal Probabilities for Tables That Read 0 to a Z-value</h2>
<Form action="cumnor.asp" method="post">
<% Randomize %>
<input Name="GivenZscore" value="<%=round(rnd()*3,2)%>" Type = "Hidden">
<input Name="GivenZscore2" value="<%=round(rnd()*3,2)%>" Type = "Hidden">
 
<INPUT TYPE=submit VALUE="GO">  <Input type=reset value="Cancel"> 
Choose and click option button and then click Go
<TAble border=2>
<tr>
   
</tr>
  
<TR>     <TD align=center>    <img src="gAs.gif">  </TD>
         <TD align=center>    <img src="gBs.gif">  </TD>
         <TD align=center>    <img src="gCs.gif">  </TD>

<TR>     <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="A"> Pr(0 &lt Z &lt Z1 )     </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="B"> Pr( -Z1 &lt Z &lt Z1 )  </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="C"> Pr( Z1 &lt Z &lt Z2 )   </TD>

<TR>     <TD align=center>    <img src="gDs.gif">  </TD>
         <TD align=center>   <img src="gEs.gif">  </TD>
         <TD align=center>   <img  src="gFs.gif">  </TD>

<TR>     <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="D"> Pr( Z1 &lt Z &lt Z2 )   </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="E"> Pr( Z &gt Z1 )    </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="F"> Pr( Z &lt -Z1 or Z &gt Z1 )   </TD>
        

<TR>     <TD align=center>   <img  src="gGs.gif">  </TD>
         <TD align=center>   <img  src="gHs.gif">  </TD>
         <TD align=center>   <img  src="gIs.gif">  </TD>

<TR>     <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="G"> Pr( Z &lt Z1 or Z &gt Z2 )   </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="H"> Pr( Z &lt Z1)                 </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="I"> Pr( Z &lt Z1 or Z &gt Z2 )   </TD>
     
     
   
       
      
     </td>
  </tr>
</table>

</FORM>



</body>
</html>

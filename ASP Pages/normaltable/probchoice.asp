
<html>


<body>

<h1> Normal Table Look-up </h1>
<Form action="cumnor.asp" method="post">
<input Name="GivenZscore" value="<%=round(rnd()*3,2)%>" Type = "Hidden">
 
<INPUT TYPE=submit VALUE="GO">  <Input type=reset value="Cancel"> 
Choose option button and then click Go
<TAble border=2>
<tr>
   
</tr>
  

<TR>     <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="A"> Pr(0 &lt Z &lt Z1 )         <img src="gAs.gif">  </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="B"> Pr( -Z1 &lt Z &lt Z1 )      <img src="gBs.gif">  </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="C"> Pr( Z1 &lt Z &lt Z2 )       <img src="gCs.gif">  </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="D"> Pr( Z1 &lt Z &lt Z2 )       <img src="gDs.gif">  </TD>
<TR>     <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="E"> Pr( Z &gt Z1 )              <img src="gEs.gif">  </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="F"> Pr( Z &lt -Z1 or Z &gt Z1 ) <img src="gFs.gif">  </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="G"> Pr( Z &lt Z1 or Z &gt Z2 )  <img src="gGs.gif">  </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="H"> Pr( Z &lt Z1)               <img src="gHs.gif">  </TD>
     
     
   
       
      
     </td>
  </tr>
</table>

</FORM>



</body>
</html>

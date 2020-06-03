
<html>


<body>

<h2> Finding Probabilities Based on the Cumulative Binomial Table</h2>
<Form action="cumbinb.asp" method="post">
<% 
'**************************************
   'Initialize starting values
   Randomize 
   intN = int(rnd()*24+5)
   randomize
   intX2 = int(rnd()*intN)
   randomize
   intX1 = int(rnd()*intX2)
   randomize
   dblP = round(rnd(),2)

  'be sure that 0 < X < X2 <= N
   if intX1 = intX2 then
     if intX2 = intN then intx1=intx1-1
     if intX1 = 0 then 
       intX2=2
       intX1=1
     end if  
   end if   
'****************************************

 

    'strMessage1 = "Given N = " & intN & " P = " & dblP & " x= " & intX
    'response.write strmessage1

%>
<input Name="GivenN" value="<%=intN%>" Type = "Hidden">
<input Name="GivenX1" value="<%=intX1%>" Type = "Hidden">
<input Name="GivenX2" value="<%=intX2%>" Type = "Hidden">
<input Name="GivenP" value="<%=dblP%>" Type  = "Hidden">
 Click on one of the following option buttons and then click go.  
<INPUT TYPE=submit VALUE="GO">  <Input type=reset value="Cancel"> 
<p>
<TAble border=2>
<tr>
   
</tr>
  
<TR>     <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="A"> Pr ( X &lt= x1 ) </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="B"> Pr ( X &lt  x1 )  </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="C"> Pr ( X &gt  x1 ) </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="D"> Pr ( X &gt= x1 ) </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="E"> Pr ( X = x1 ) </TD>
<TR>     <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="F"> Pr ( x1 &lt= X &lt= x2 )  </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="G"> Pr ( x1 &lt= X &lt  x2 )   </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="H"> Pr ( x1 &lt  X &lt= x2 )  </TD>
         <TD>   <INPUT TYPE=RADIO NAME=GO VALUE="I"> Pr ( x1 &lt  X &lt  x2 )   </TD>
     
     
   
       
      
     </td>
  </tr>
</table>


</FORM>



</body>
</html>

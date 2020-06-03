<html>
<head>

</head>


<body>
<form action=VarCalcSol.asp method=POST enctype="application/x-www-form-urlencoded">
 <%   dim sumX, i, sngmean
   SSXX = Array(0.1,1.1,2.1,3.1,4.1,5.1)
   SSDevX = Array(0.1,1.1,2.1,3.1,4.1,5.1)
   X = array(0.1,1.1,2.1,3.1,4.1,5.1)
   numAttempts=0
   sumX=0.0 
  randomize
  
     x(1)=round(rnd()*5+.5,0)  
     x(2)=round(rnd()*5+.5,0)  
     x(3)=round(rnd()*5+.5,0)  
     x(4)=round(rnd()*5+.5,0)  
     x(5)=round(rnd()*5+.5,0)  

     Session("X1") = x(1)  
     Session("X2") = x(2)  
     Session("X3") = x(3)
     Session("X4") = x(4)
     Session("X5") = x(5)
     Session("NumAttempts")=numAttempts
     Name=""
     if Session("Name")<>"" then name=replace(session("Name")," ","-")
 %><!--  -->
<h1 align="center"> Calculate the Variance and Standard Deviation</h1>
<p align="center"> Round final answers to three decimal places</p>

<table align="center" border="1" >
  <tr>
     <td align="center">Values</td>
  </tr>
   <tr>
     <td align="center"> 
       <%response.write(x(1))%>  <br>
       <%response.write(x(2))%>  <br>
       <%response.write(x(3))%>  <br>
       <%response.write(x(4))%>  <br>
       <%response.write(x(5))%>  <br>
     </td>
   </tr>	
  <tr>
     <td>
      1. Enter your name, last name first
		<input name="name" value=<%=Name%>>
      	        
     </td>
    </tr>
   <tr>
     <td>
      2. Enter the average
		<input name="mean"  >
      	        
     </td>
    </tr>
   <tr> 
     <td>
      3. Assuming its a population <br>
          Enter the population variance value <input name="varP" >
      	        <br>
          Enter the population standard deviation value <input name="SD_P" >
      	        

     </td>
    </tr>

    <tr> 
     <td>
      4. Assuming its a sample<br>
          Enter the sample variance value <input name="varS" >
      	        <br>
          Enter the sample standard deviation value <input name="SD_S" >
      	        

     </td>
    </tr>
</table>
<br><br>
<input type="submit" value="Please click here to check answers" name="END">

</body>

</html>
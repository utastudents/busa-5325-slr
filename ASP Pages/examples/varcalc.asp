<html>
<head>

</head>


<body lang=EN-US style='tab-interval:.5in'>
 <%   dim sumX, i, sngmean
   SSXX = Array(0.1,1.1,2.1,3.1,4.1,5.1)
   SSDevX = Array(0.1,1.1,2.1,3.1,4.1,5.1)
   X = array(0.1,1.1,2.1,3.1,4.1,5.1)
  
   sumX=0.0 
  randomize
   
     x(1)=round(rnd()*5+.5,0)
     x(2)=round(rnd()*5+.5,0)
     x(3)=round(rnd()*5+.5,0)
     x(4)=round(rnd()*5+.5,0)
     x(5)=round(rnd()*5+.5,0)
     sumx=x(1)+x(2)+x(3)+x(4)+x(5)
   

     sngMean=sumx/5.0
   
  'Calculate deviations from mean, squared deviations, and statistics
   sumXX = 0.0 
   
   for i=1 to 5
     SSDevX(i) = x(i)-sngmean
     SSXX(i) = round(SSDevX(i)^2.,2)
     SumXX = SumXX + SSXX(i)
   next 
     
   VarP=SumXX/5.0
   VarS=SumXX/4.0
   SD_P = round(sqr(varp),8)
   SD_S= round(sqr(vars),8)    
 %><!--  -->
<h1 align="center"> Calculating Variance and Standard Deviation</h1>
<table align="center" border="1" >
  <tr>
     <th></th>
     <th>Step b</th>
     <th>Step c</th>
  </tr>
  <tr>
     <td align="center">Values</td>
     <td title="Take each value and subtract the mean.">Distance to Average</td>
     <td title="Square the values in previous column.">Square the Distances</td>
  </tr>
   <tr>
     <td align="center" Title="First step is to average the values"> <!--sup is added to line up the values-->
       <%response.write(x(1))%>       <sup> </sup><br>
       <%response.write(x(2))%>       <sup> </sup><br>
       <%response.write(x(3))%>       <sup> </sup><br>
       <%response.write(x(4))%>       <sup> </sup><br>
       <%response.write(x(5))%>       <sup> </sup><br>
       <b>Step a:</b> mean=<%response.write(sngmean)%><br>
     </td>
     
     <td align="center" title="Subtract the mean from each value.">
        <%response.write( "(" & x(1) & "-" & sngmean & ") = " & ssdevx(1))%>       <sup> </sup><br>
        <%response.write( "(" & x(2) & "-" & sngmean & ") = " & ssdevx(2))%>       <sup> </sup><br>
        <%response.write( "(" & x(3) & "-" & sngmean & ") = " & ssdevx(3))%>       <sup> </sup><br>
        <%response.write( "(" & x(4) & "-" & sngmean & ") = " & ssdevx(4))%>       <sup> </sup><br>
        <%response.write( "(" & x(5) & "-" & sngmean & ") = " & ssdevx(5))%>       <sup> </sup><br> 
        <br>
     </td>

      <td align="center" title="Square the values in previous column.">
        <%response.write( "(" & ssdevx(1) & ")")%>       <sup>2 </sup> <%Response.write(" = " & SSXX(1))%> <br>
        <%response.write( "(" & ssdevx(2) & ")")%>       <sup>2 </sup> <%Response.write(" = " & SSXX(2))%> <br>
        <%response.write( "(" & ssdevx(3) & ")")%>       <sup>2 </sup> <%Response.write(" = " & SSXX(3))%> <br>
        <%response.write( "(" & ssdevx(4) & ")")%>       <sup>2 </sup> <%Response.write(" = " & SSXX(4))%> <br>
        <%response.write( "(" & ssdevx(5) & ")")%>       <sup>2 </sup> <%Response.write(" = " & SSXX(5))%> <br>
        <br>
      </td>  

   <tr>
     <td><b>Step d:</b></td>
     <td>Sum of squares =</td>
     <td align="center" title="Sum the squared values"> <%response.write(sumxx)%> </td>
   </tr>

  
  <tr>
     <td><b>Step e.</b><br>
            Variance      
     </td>
     <td>
         Population: &#963<sup>2</sup> =<br>
         Sample: s<sup>2</sup> =<br> 
     </td>
     <td align="center" title="If population, divide the Sum by n. If sample, divide it by n-1.">
         <%response.write(sumxx)%>/5=<%response.write(varp)%><sup> </sup><br>
         <%response.write(sumxx)%>/4=<%response.write(vars)%><sup> </sup><br> 
         
     </td>
   <tr>
  <tr>
     <td><b>Step f.</b><br>
        Standard Deviation<br>
        
     </td>

     <td>
         
         Population: &#963 =<br>
         Sample: s =<br> 
     </td>
     <td align="center" title="Take the square root of the above variances.">
       
          &#8730<font style="text-decoration: overline;"><%response.write(varp)%></font>=<%response.write(SD_P)%><br>
          &#8730<font style="text-decoration: overline;"><%response.write(vars)%></font>=<%response.write(SD_S)%><br>
     </td>
 

</table>		
Click <a href="<%= request.servervariables("script_name") %>">here</a> for a new example
</body>

</html>
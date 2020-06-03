<html>
<head>

</head>


<body lang=EN-US style='tab-interval:.5in'>
 <%   dim sumX, i, sngmean
   x1=session("X1")
   x2=session("X2")
   x3=session("X3")
   x4=session("X4")
   x5=session("X5")

   NumAttempts=session("NumAttempts")
   NumAttempts=NumAttempts+1
   session("numattempts")=numattempts

   SSXX = Array(0.1,1.1,2.1,3.1,4.1,5.1)
   SSDevX = Array(0.1,1.1,2.1,3.1,4.1,5.1)
   X = array(0.1,1.1,2.1,3.1,4.1,5.1)
   x(1) = x1
   x(2) = x2
   x(3) = x3
   x(4) = x4
   x(5) = x5

   sumX=0.0 
 
    MeanAns = Request.Form("mean")
    VarpAns =  Request.Form("varp")
    SD_PAns = Request.Form("SD_P")
    VarsAns =  Request.Form("varS")
    SD_SAns = Request.Form("SD_S")
    Name = request.form("name")
    session("Name")=replace(name," ","-")   

     sumx=x(1)+x(2)+x(3)+x(4)+x(5)
   

     sngMean=sumx/5.0
   
  'Calculate deviations from mean, squared deviations, and statistics
   sumXX = 0.0 
   
   for i=1 to 5
     SSDevX(i) = x(i)-sngmean
     SSXX(i) = round(SSDevX(i)^2.,2)
     SumXX = SumXX + SSXX(i)
   next 
     
   VarP=round(SumXX/5.0,3)
   
   VarS=round(SumXX/4.0,3)
   
   SD_P = round(sqr(varp),3)
   SD_S= round(sqr(vars),3) 

   response.write("<h1 align=" & chr(34) & "center" & chr(34) & _
           "> Calculating Variance and Standard Deviation</h1>")
   response.write("<p align=" & chr(34) & "center" & chr(34) & _
           ">Number of Attempts = " & NumAttempts & "</p>")
  
   function ChkAns(a,b)
      if abs(a-b) < 0.002 then
        chkans=1
      else
        chkans=0
      end if
    end function
  %>


<table align="center">
  <tr> <td>
   <b><%response.write("Your name is: " & name)%></b> <br>
   <% 
     if isnumeric(meanans) then 
        if chkans(meanans,sngmean) = 1 then
       'if round(csng(meanans),1)=round(sngmean,1) then  
           response.write("Mean: Correct")
       else
           response.write("Mean: Incorrect")
       end if
     else
         response.write("Your value of the mean was not a number.")
     end if
     %>

   <br>
 
   <% 
    if isnumeric(varpans) then 
        if chkans(varpans,varp) = 1 then
        'if round(csng(varpans),1) =round(varp,1) then 
          response.write("Population variance: Correct")
        else
          response.write("Population variance: Incorrect. ")
        end if
     else
         response.write("Your value of the population variance was not a number.")
     end if
   %>
   
   <br>

  <% 
    if isnumeric(sd_Pans) then 
        if chkans(sd_pans,sd_p) = 1 then
       'if round(csng(sd_Pans),1)=round(sd_p,1) then
          response.write("Population standard deviation:Correct")
       else
          response.write("Population standard deviation: Incorrect. ")
       end if
     else
         response.write("Your population standard deviation value was not a number.")
     end if
   %>

   <br>

   <%  
    if isnumeric(varsans) then 
       if round(csng(varsans),1)=round(vars,1) then
          response.write("Sample variance: Correct")
       else
          response.write("Sample variance: Incorrect. ")
       end if
     else
         response.write("Your sample standard deviation value was not a number.")
     end if
   %>

   <br>

     <% 
    if isnumeric(sd_sans) then 
       if round(csng(sd_sAns),1)=round(sd_s,1) then
          response.write("Sample standard deviation: Correct")
       else
          response.write("Sample standard deviation: Incorrect. ")
       end if
     else
         response.write("Your value for the sample variance was not a number.")
     end if
   %>

   <br> 

  </td></tr>
</table> <br> <br>

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
         <%response.write(sumxx)%>/5=<%vp=round(varp,2):response.write(vp)%><sup> </sup><br>
         <%response.write(sumxx)%>/4=<%vs=round(vars,2):response.write(vs)%><sup> </sup><br> 
         
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
       
          &#8730<font style="text-decoration: overline;"><%response.write(vp)%></font>=<%sp=round(sd_p,3):response.write(sp)%><br>
          &#8730<font style="text-decoration: overline;"><%response.write(vs)%></font>=<%ss=round(sd_s,3):response.write(SS)%><br>
     </td>
 

</table>
		
Click <a href="https://wweb.uta.edu/faculty/eakin/asps/examples/varcalcques.asp">here</a> for a new question
</body>

</html>
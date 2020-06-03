<html>
<head>

</head>


<body lang=EN-US style='tab-interval:.5in'>
 <%   dim sumX, i, sngmean
   place=Session("place")
   obsr=Session("obsr")
   obsrs=Session("obsrs")
   Cat1L1=Session("Cat1L1")
   Cat1L2=Session("Cat1L2")
   Cat1L1obsr=Session("Cat1L1obsr")
   Cat2L1=Session("Cat2L1")
   Cat2L2=Session("Cat2L2")
   Cat2L1obsr=Session("Cat2L1obsr")
   Paragraph= Session("Paragraph") 
   Sentence2= Session("Sentence2") 
   Sentence3= Session("Sentence3") 

   Cond1 = csng(session("Cond1"))
   Cond2 = csng(session("Cond2"))
   Margin1 = csng(session("Margin1"))
   NumAttempts = session("NumAttempts")
   NumAttempts = numattempts+1
   session("Numattempts") = numattempts

   sumX=0.0 
  randomize
    StrNumProblem="None"
    StudentNum=request.form("ansnum")
    if len(trim(studentNum))=0 then strNumProblem="No data"
    studentNum = replace(studentNum,"%","")
    if isnumeric(studentNum)=false then 
       strNumproblem = "Your numerator of " & studentNum &  " is not a number" 
    else 
       NumAns = csng(studentNum)
    end if

    StrDenProblem="None"
    StudentDen = request.form("ansden")
    if len(trim(studentDen))=0 then strDenProblem="No data"
    StudentDen = replace(studentDen,"%","")  
    if isnumeric(studentden)=false then 
       strDenproblem = "Your denominator " & studentDen &  " is not a number" 
    else 
       DenAns = csng(studentden)
       if Denans=0 then strDenProblem="The Division by zero"
    end if
    

    name=request.form("Name")
    
    session("Name")=replace(name,",","-")
    session("Name")=replace(name," ","-")     
 
    if strNumProblem="None" and strDenProblem="None" then
      StudentAnswer = csng(NumAns)/csng(DenAns)
   end if    
  

   Numerator = margin1*cond1/100
   Denominator =(numerator+(100-Margin1)*Cond2/100)
   Answer = csng(numerator)/csng(denominator)


   %>
<h1 align="center"> Answer to Bayes Exercise</h1>


<table align="center">
  <tr> <td>
   <b><%response.write("Your name is: " & name)%></b> <br>
   <% 
     if strNumProblem="None" and strDenProblem="None" then 

       if round(csng(studentanswer),1)=round(answer,1) then
         response.write("Your answer  is Correct")
       else
       response.write("Your answer of " & numans & " / " & denans & " is Incorrect")
       end if

     else
       response.write("There was a problem with the way you entered your answer.<br>")
       response.write("Numerator Error: " & strNumproblem & "<br>")
       response.write("Denominator Error: " & strDenProblem) 
     end if
       response.write("<br>Number of attempts = " & numattempts)
     %>
   <br>

  
  </td></tr>
</table> <br> <br>
  
  
<table align="center" border="1" width="600">

   <tr>
     <td align="left" style="word-wrap: break-word"> 
       <%response.write(paragraph)%>  <br>

     </td>
   </tr>
   <tr>
     <td>
      <b> Solution </b>
      	        
     </td>
    </tr>	
  <tr>
     <td>
      <%Response.write("a. Determine the percent in the first major category, those that " & Cat1L1  &"."    )%><br><br>
      <%Response.write("To start lets assume there are 10,000 " & obsrs & ".")%><br>
      <%Response.write( Margin1 & "% or " & Margin1*100 & " " & obsrs & " " & Cat1L1)%><br>
      <%Response.write("therefore")%><br>
      <%Response.write( (100-Margin1) & "% or " & (100-Margin1)*100 & " "  & Cat1L2 & ".")%><br><br>
     </td>
    </tr>
   <tr>
     <td>
      <%Response.write("b. Of the levels of the first major category, what percent of population belong to the second?" )%><br>
      <br>
      <%Response.write(sentence2)%> <br>
      and <br>
      <%Response.write(sentence3)%> <br>
    	        
     </td>
    </tr>
    
    <tr>
     <td>
      <%Response.write("c. Determine the percent of all " & obsrs & " that have the combined characteristics." )%><br>
      <br>
      <%Response.write("(" & Cond1 & "% of " & Margin1*100 & ") or  " & Cond1*Margin1 & " " & obsrs & " both " & Cat1L1 & " and " & Cat2L1)%> <br>
      while <br>
      <%Response.write("(" & Cond2 & "% of " & (100-Margin1)*100 & ") or  " & (100-Margin1)*Cond2 & " " & obsrs & " both " & Cat1L2 & " and " & Cat2L1)%> <br>
    	        
     </td>
    </tr>

    <tr>
     <td>
      <%Response.write("d. Determine what is the total percent of all " & obsr & " that " & Cat2L1 )%><br>
      <br>
      <%Response.write("Of the total number that " & Cat2L1 & " some " & Cat1L1 & " and some " & Cat1L2)%> <br>
      <%Response.write("This is the sum of the previous two values: " & Cond1*Margin1 & " and " & Cond2*(100-Margin1) & ".")%> <br>
      <%Response.write("Therefore " & (Cond1*Margin1+(100-Margin1)*Cond2) & " " & obsrs & " "  & Cat2L1)%> <br>
      	        
     </td>
    </tr>

    <tr>
     <td>
      <%Response.write("e. To find your answer to the problem, determine among all " & obsr & " that " & Cat2L1 & " what proportion " & Cat1L1)%><br><br>
      
      <%Response.write("Of the " & (Margin1*Cond1+(100-Margin1)*Cond2) & "  that " & Cat2L1 & ", " & Margin1*Cond1 & " " & Cat1L1)%> <br>
      <%Response.write("Calculate the fraction " & Cond1*Margin1 & " divided by " & ((100-Margin1)*Cond2+Margin1*Cond1) & "")%> <br>
      <%Response.write("The answer to the exercise is " & Numerator*100 & "/" & Denominator*100 & " = " & ROUND(Numerator/Denominator,4)*100 & "% (rounding to two decimal places.)")%> <br>
      	        
     </td>
    </tr>

</table>
Click <a href="Bayesques.asp">here</a> for a new question
</body>

</html>
<html>
<head>

</head>


<body lang=EN-US style='tab-interval:.5in'>
<h1 align="center">Probability of Sample Proportion Exercise </h1>
<br>
 <%   
   Ans1=request.form("AnsQues1")
   Ans2=request.form("AnsQues2")
   Ans3=request.form("AnsQues3")
   Ans4L=request.form("AnsQues4Lower")
   Ans4U=request.form("ansques4Upper")

   x =   Session("x")	  
   mea = Session("mea")	
   obsr = Session("obsr")	
   mu = Session("mu")	
   sig = Session("sig")	
   n = Session("n")
   SE = Session("SE")	
   bounds = Session("bounds")
   Prob = Session("Prob")	
   prless = Session("prless")
   prGreater = Session("prGreater")
   LessThanProb = Session("LessThanProb")
   GrThanProb = Session("GrThanProb")
   BetweenProb= Session("BetweenProb")
   LessThanValue=Session("LessThanValue") 
   GrThanValue =Session("GrThanValue")
   WithinValue =Session("WithinValue")

   UpperValue=Session("UpperValue") 
   LowerValue=Session("LowerValue")
   CenterProb=Session("CenterProb")
   PostFix = Session("PostFix")
   X1_for_pctile = Session("X1_for_pctile")
   X2_for_pctile = session("x2_for_pctile")
   paragraph= Session("paragraph")


    name=request.form("Name")
    session("Name")=replace(name,",","-")
    session("Name")=replace(name," ","-")     

   NumAttempts = session("NumAttempts")
   NumAttempts = numattempts+1
   session("Numattempts") = numattempts

   prhtml("<p align=" & chr(34) & "center" & chr(34) & "> Number of attempts = " & numattempts & "<p>")
   
    sub PrHtml(strHtml)
       response.write(strhtml)
    end sub
   
   %>
   <p align="center"> <b><%response.write("Your name is: " & name)%></b> </p>
   <table align="center" border="1" width="600">
     <tr>
       <td align="left" style="word-wrap: break-word"> 
         <%response.write(paragraph)%>  <br>
       </td>
     </tr>
     </table>	
  
   <br>

      
   <%
    Prhtml("<table align=" & chr(34) & "center" & chr(34) & "><tr><td>")

    prhtml("<br>In case of an error, if you are using Internet Explorer, then " & _
              "go Back to the previous page <br> and try again. The program will " & _
              "keep track of the number of attempts.<br><br>")
 
       prhtml("</td></tr><tr><td>")  

    Prhtml("<p>A. Identify: This is asking questions about sample proportions.</p>")
   
    Prhtml("<p>B. Determine if the normal approximation to the probability is adequate: Since <br>")

    if mu>1.-mu then minProb = (1.-mu) else minProb = mu

    if n*minprob > 5 then
      Prhtml("both " & n & "*" & mu & " and " & n & "*(1-" & mu & ") are greater than 5, probabilities from the<br>")
      Prhtml("standard normal table can be used. </p>")
    else
       Prhtml(n & "*" & minprob & " = " & n*minprob & " is not greater than 5, the probabilities from the standard <br>")
       Prhtml("normal table used in the answers below may be far " & _ 
       "from the actual answers. </p>")
    end if
      
    Prhtml("<p>C. Calculate the standard error: This will be the square root of [" & _
             "the population proportion <br> of successes " & _
             "times the population proportion of failures divided by the sample size]<br> = " &  _
             " the square root of (" &  mu & "*" & (1-mu) & "/" & n & ") = " & se & "</p>")
      



  'Question 1
    PrHtml("Question 1: ") 
      if isnumeric(ans1) = true then
    
        If chkans(ans1,LessThanProb) =1 then
         
           PrHtml("Correct: The probability of finding " & obsr & " with a value of " & x & _
                  " less than " & LessThanValue & " " & mea & " is " & round(LessThanProb,4) & "<br>")
           
         Else
            Prhtml("Incorrect: Your answer of " & ans1 & " is wrong. <br>")
         end if
      elseif len(ans1) =0 then
         prhtml("You did not enter an answer to Question 1. <br>")
      else
            Prhtml("Your answer of " & ans1 & " is not a number between 0 and 1.<br>")
      end if
     prhtml("<br></td></tr><tr><td>")

    PrHtml("Question 2: ") 
      if isnumeric(ans2) = true then
    
        If chkans(ans2,GrThanProb) =1 then
         
           PrHtml("Correct: The probability of finding " & obsr & " with a value of " & x & _
                   " greater than " & GrThanValue & " " & mea & " is " & round(GrThanProb,4) & "<br>")
           
         Else
            Prhtml("Incorrect: Your answer of " & ans2 & " is wrong. <br>")
         end if
      elseif len(ans2) =0 then
         prhtml("You did not enter an answer to Question 2. <br> ")
      else
            Prhtml("Your answer of " & ans2 & " is not a number between 0 and 1.<br>")
      end if
     prhtml("<br></td></tr><tr><td>")
    
     PrHtml("Question 3: ") 
      if isnumeric(ans3) = true then
    
        If chkans(ans3,BetweenProb) =1 then
         
           PrHtml("Correct: The probability of finding a sample average within " & withinValue & _
          " " & mea & " of the population mean is " & round(BetweenProb,4) & "<br>")
           
         Else
            Prhtml("Incorrect: Your answer of " & ans3 & " is wrong. <br>")
         end if
      elseif len(ans3) =0 then
         prhtml("You did not enter an answer to Question 3. <br> ")
      else
            Prhtml("Your answer of " & ans3 & " is not a number between 0 and 1. <br>")
      end if
     prhtml("<br></td></tr><tr><td>")
     
     PrHtml("Question 4 Lower Value: ") 
      if isnumeric(ans4L) = true then
    
        If abs(ans4l-X1_for_pctile) < 0.01*x1_for_pctile then
          
         
           PrHtml("Correct: " & round(X1_for_pctile,2) & " is the correct lower value. <br>")
  
         Else
            Prhtml("Incorrect: Your answer for the lower value:  " & ans4L & " is wrong. <br>")
         end if
      elseif len(ans4l) =0 then
         prhtml("You did not enter a lower value answer to Question 4 . <br>")
      else
            Prhtml("Your answer of " & ans4l & " for the lower value is not a number.<br>")
      end if
     
     PrHtml("Question 4 Upper Value: ") 
      if isnumeric(ans4U) = true then
    
        If abs(ans4u-X2_for_pctile) < 0.01*x1_for_pctile then
          
         
           PrHtml("Correct: " & round(X2_for_pctile,2) & " is the correct upper value. <br>")
  
         Else
            Prhtml("Incorrect: Your answer for the upper value:  " & ans4u & " is wrong. <br>")
         end if
      elseif len(ans4u) =0 then
         prhtml("You did not enter an upper value answer to Question 4 . <br>")
      else
            Prhtml("Your answer of " & ans4u & " for the upper value is not a number.<br>")
      end if

     prhtml("<br></td></tr><tr><td>")
     prhtml("</td></tr></table>")

   function ChkAns(a,b)
      if abs(a-b) < 0.0015 then
        chkans=1
      else
        chkans=0
      end if
    end function
    %>
Click <a href="ProbOfPQues.asp">here</a> for a new question
</body>

</html>